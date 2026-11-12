#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サンプル第4弾「インバウンド観光ダッシュボード」— 整形スクリプト

data/raw/ の公開ファイルを読み、data/processed/ に tidy 形式CSVを出力する。
出典は3レイヤー (混在させない。各CSVに source 列で明示):
  L1 主レイヤー: 観光庁「宿泊旅行統計調査」(政府統計コード 00601020)
  L2 補助レイヤー: JNTO「訪日外客数」(国籍/月別)
  L3 補助レイヤー: 観光庁「インバウンド消費動向調査」暦年確報 (2019・2024・2025)

出力:
  guest_nights_monthly.csv      L1 月次×都道府県 総/外国人 延べ宿泊者数 (2015-01〜2026-09)
  guest_nights_annual.csv       L1 年次×都道府県 同上
  nationality_nights_monthly.csv L1 月次×都道府県×国籍 外国人延べ宿泊者数
                                 (従業者数10人以上の施設が対象の参考第1表ベース)
  nationality_nights_annual.csv L1 年次×都道府県×国籍 同上
  occupancy_monthly.csv         L1 月次×都道府県 客室稼働率 (総計)
  occupancy_annual.csv          L1 年次×都道府県 客室稼働率 (総計)
  occupancy_by_type_annual.csv  L1 年次×都道府県×施設タイプ 客室稼働率
  nights_by_type_annual.csv     L1 年次×都道府県×施設タイプ 延べ宿泊者数 (総/外国人)
  arrivals_monthly.csv          L2 月次×国籍 訪日外客数 (2015-01〜2026-09、総数含む)
  spend_per_person.csv          L3 年次×国籍 1人1回当たり旅行消費単価・平均泊数
  spend_by_category.csv         L3 年次×国籍×費目 1人当たり消費単価と構成比
  spend_total_official.csv      L3 年次 訪日外国人旅行消費額 総額 (報道発表の確報値)

数値の扱い: 公表値の '*' (秘匿等) は数値として読み取り、'-'・空欄は欠損 (空欄出力)。
外国人比率は算出値 (外国人÷総数×100) で、CSVには derived であることを列名で示す。
"""
import csv
import pathlib
import re
from collections import defaultdict

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed"
OUT.mkdir(parents=True, exist_ok=True)

SRC_L1 = "観光庁 宿泊旅行統計調査"
SRC_L2 = "JNTO 訪日外客数（国籍/月別）"
SRC_L3 = "観光庁 インバウンド消費動向調査"

PREFS = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
    "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
    "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
    "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
]
PREF_SET = set(PREFS)
TYPES = ["旅館", "リゾートホテル", "ビジネスホテル", "シティホテル",
         "簡易宿所", "会社・団体の宿泊所"]


def num(v):
    """公表値セルの数値化。'-'・空欄は None、'*123' は 123。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).replace(",", "").replace("*", "").strip()
    if s in ("", "-", "－", "***", "…", "..."):
        return None
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except ValueError:
        return None


def norm_label(v):
    if v is None:
        return ""
    s = str(v).replace("　", "").replace(" ", "").replace("\n", "")
    s = re.sub(r"^\d{2}", "", s)
    return s


def norm_nat(name):
    s = norm_label(name)
    return {"アメリカ": "米国", "イギリス": "英国"}.get(s, s)


def pref_rows(ws):
    """シートから (全国行, {都道府県: 行}) を返す。行はセル値タプル (0-index)。"""
    total_row = None
    rows = {}
    for row in ws.iter_rows(values_only=True):
        lab = norm_label(row[0])
        if lab in PREF_SET:
            rows[lab] = row
        elif total_row is None and not rows and lab and num(row[1]) is not None:
            # 全国行: 年計は「…1～12月計」、月次は「令和7年1月」のように現れる
            total_row = row
    return total_row, rows


def sheet_map(wb):
    m = {}
    for n in wb.sheetnames:
        mm = re.match(r"^(第\d+表|参考第\d+表)\((\d+)月\)$", n)
        if mm:
            m[(mm.group(1), int(mm.group(2)))] = n
        elif re.match(r"^(第\d+表|参考第\d+表)\(年計\)$", n):
            m[(n.split("(")[0], 0)] = n
    return m


# ---------------------------------------------------------------- L1 宿泊統計
def parse_accommodation():
    gn_m, gn_a, nat_m, nat_a = [], [], [], []
    occ_m, occ_a, occ_t, nbt = [], [], [], []

    def handle_sheet(ws, ym, table, annual=False):
        total_row, rows = pref_rows(ws)
        targets = [("全国", total_row)] + [(p, rows.get(p)) for p in PREFS]

        if table == "R1":
            hdr = None
            for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
                if row[2] and norm_nat(row[2]) in ("韓国", "中国"):
                    hdr = row
                    break
            nats = []
            if hdr:
                for c in hdr[2:]:
                    if c:
                        nats.append(norm_nat(c))
            for pref, row in targets:
                if row is None:
                    continue
                for i, nat in enumerate(nats):
                    v = num(row[2 + i]) if 2 + i < len(row) else None
                    rec = dict(year_month=ym, prefecture=pref, nationality=nat,
                               foreign_nights=v, source=SRC_L1,
                               table="参考第1表", scope="従業者数10人以上の施設")
                    (nat_a if annual else nat_m).append(rec)
            return

        for pref, row in targets:
            if row is None:
                continue
            if table == "T2":
                total, foreign = num(row[1]), num(row[16])
                share = round(foreign / total * 100, 2) if total and foreign is not None else None
                rec = dict(year_month=ym, prefecture=pref, total_nights=total,
                           foreign_nights=foreign, foreign_share_pct_calc=share,
                           source=SRC_L1, table="第2表")
                (gn_a if annual else gn_m).append(rec)
            elif table == "T8":
                rec = dict(year_month=ym, prefecture=pref, occupancy_pct=num(row[1]),
                           source=SRC_L1, table="第8表")
                (occ_a if annual else occ_m).append(rec)
                if annual:
                    for i, t in enumerate(TYPES):
                        occ_t.append(dict(year=int(ym[:4]), prefecture=pref,
                                          facility_type=t,
                                          occupancy_pct=num(row[16 + i]),
                                          source=SRC_L1, table="第8表"))
            elif table == "T4" and annual:
                for i, t in enumerate(TYPES):
                    nbt.append(dict(year=int(ym[:4]), prefecture=pref, facility_type=t,
                                    total_nights=num(row[2 + i]),
                                    foreign_nights=num(row[9 + i]),
                                    source=SRC_L1, table="第4表"))

    for year in range(2015, 2026):
        wb = openpyxl.load_workbook(RAW / f"annual_{year}.xlsx",
                                    read_only=True, data_only=True)
        sm = sheet_map(wb)
        for m in range(1, 13):
            ym = f"{year}-{m:02d}"
            for tname, key in [("第2表", "T2"), ("参考第1表", "R1"), ("第8表", "T8")]:
                sn = sm.get((tname, m))
                if sn:
                    handle_sheet(wb[sn], ym, key)
        for tname, key in [("第2表", "T2"), ("参考第1表", "R1"), ("第8表", "T8"), ("第4表", "T4")]:
            sn = sm.get((tname, 0))
            if sn:
                handle_sheet(wb[sn], f"{year}", key, annual=True)
        wb.close()
        print(f"L1 annual {year} done")

    for m in range(1, 10):
        f = RAW / (f"m2026_{m:02d}.xlsx" if m <= 8 else "m2026_09_prelim.xlsx")
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        sm = sheet_map(wb)
        ym = f"2026-{m:02d}"
        for tname, key in [("第2表", "T2"), ("参考第1表", "R1"), ("第8表", "T8")]:
            sn = sm.get((tname, m))
            if sn:
                handle_sheet(wb[sn], ym, key)
        wb.close()
        print(f"L1 2026-{m:02d} done")

    return gn_m, gn_a, nat_m, nat_a, occ_m, occ_a, occ_t, nbt


# ---------------------------------------------------------------- L2 JNTO
JNTO_COUNTRIES = {
    "韓国", "中国", "台湾", "香港", "タイ", "シンガポール", "マレーシア",
    "インドネシア", "フィリピン", "ベトナム", "インド", "米国", "アメリカ",
    "カナダ", "英国", "ドイツ", "フランス", "イタリア", "スペイン",
    "ロシア", "オーストラリア", "その他",
}


def parse_jnto():
    wb = openpyxl.load_workbook(RAW / "jnto_arrivals_monthly_by_nationality.xlsx",
                                read_only=True, data_only=True)
    out = []
    for year in range(2015, 2027):
        if str(year) not in wb.sheetnames:
            continue
        ws = wb[str(year)]
        # 「1月」見出しの列位置は年によってずれる (2019年以前は col2、2020年以降は col3)
        jan_col = 1
        for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
            if row and "1月" in [str(c).strip() if c else "" for c in row]:
                jan_col = [str(c).strip() if c else "" for c in row].index("1月")
                break
        for row in ws.iter_rows(values_only=True):
            lab = norm_nat(row[0]) if row and row[0] else ""
            if lab == "総数" or lab in JNTO_COUNTRIES:
                for m in range(1, 13):
                    ci = jan_col + 2 * (m - 1)
                    v = num(row[ci]) if len(row) > ci else None
                    if v:
                        out.append(dict(year_month=f"{year}-{m:02d}",
                                        nationality=("総数" if lab == "総数" else lab),
                                        arrivals=v, source=SRC_L2))
        print(f"L2 {year} done")
    wb.close()
    return out


# ---------------------------------------------------------------- L3 消費動向
SPEND_TOTALS = {
    2019: (48135, "観光庁 報道発表 2019年暦年確報"),
    2024: (81257, "観光庁 報道発表 2024年暦年確報"),
    2025: (94549, "観光庁 報道発表 2025年暦年確報"),
}
MAJOR_CATS = ["宿泊費", "飲食費", "交通費", "娯楽等サービス費", "買物代", "その他"]


def nat_header(ws):
    """国籍ヘッダ (5行目、回答数/値の2列組) を [(国籍名, 値列index)] で返す。"""
    for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
        if row and len(row) > 5 and isinstance(row[4], str) and "全国籍" in row[4]:
            pairs = []
            c = 4
            while c + 1 < len(row) and row[c] is not None:
                pairs.append((norm_nat(row[c]), c + 1))
                c += 2
            return pairs
    return []


def find_row(ws, keyword, col=1):
    for row in ws.iter_rows(values_only=True):
        lab = norm_label(row[col]) if len(row) > col else ""
        if keyword in lab:
            return row
    return None


def parse_spend():
    per_person, by_cat, totals = [], [], []
    for year in (2019, 2024, 2025):
        wb = openpyxl.load_workbook(RAW / f"consumption_{year}_converted.xlsx",
                                    read_only=True, data_only=True)
        spend_map = {}
        for sheet, field in [("参考1", "per_person_spend_yen"), ("表4-1", "avg_nights")]:
            ws = wb[sheet]
            pairs = nat_header(ws)
            row = find_row(ws, "全体")
            for name, ci in pairs:
                spend_map.setdefault(name, {})[field] = num(row[ci]) if row else None
        for nat, vals in spend_map.items():
            per_person.append(dict(year=year, nationality=nat,
                                   per_person_spend_yen=vals.get("per_person_spend_yen"),
                                   avg_nights=vals.get("avg_nights"),
                                   source=SRC_L3,
                                   note="参考1（パッケージツアー参加費内訳含む）・表4-1【A1全国調査】試算値"))
        ws = wb["参考2"]
        # 参考2 は (消費単価, 構成比) の2列組で、国籍名の列がそのまま消費単価列
        pairs2 = []
        for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
            if row and len(row) > 5 and isinstance(row[4], str) and "全国籍" in row[4]:
                c = 4
                while c + 1 < len(row) and row[c] is not None:
                    pairs2.append((norm_nat(row[c]), c))
                    c += 2
                break
        for row in ws.iter_rows(values_only=True):
            cat = norm_label(row[2]) if len(row) > 2 and row[2] else ""
            if cat in MAJOR_CATS:
                for name, ci in pairs2:
                    by_cat.append(dict(year=year, nationality=name, category=cat,
                                       per_person_spend_yen=num(row[ci]),
                                       share_pct=num(row[ci + 1]) if ci + 1 < len(row) else None,
                                       source=SRC_L3, table="参考2"))
        wb.close()
        v, src = SPEND_TOTALS[year]
        totals.append(dict(year=year, total_spend_oku_yen=v, source=src,
                           note="訪日外国人旅行消費額（確報・推計値）"))
        print(f"L3 {year} done")
    return per_person, by_cat, totals


def write_csv(name, rows, fields):
    with open(OUT / name, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k) for k in fields})
    print(f"{name}: {len(rows)} rows")


def main():
    gn_m, gn_a, nat_m, nat_a, occ_m, occ_a, occ_t, nbt = parse_accommodation()
    write_csv("guest_nights_monthly.csv", gn_m,
              ["year_month", "prefecture", "total_nights", "foreign_nights",
               "foreign_share_pct_calc", "source", "table"])
    write_csv("guest_nights_annual.csv", gn_a,
              ["year_month", "prefecture", "total_nights", "foreign_nights",
               "foreign_share_pct_calc", "source", "table"])
    write_csv("nationality_nights_monthly.csv", nat_m,
              ["year_month", "prefecture", "nationality", "foreign_nights",
               "source", "table", "scope"])
    write_csv("nationality_nights_annual.csv", nat_a,
              ["year_month", "prefecture", "nationality", "foreign_nights",
               "source", "table", "scope"])
    write_csv("occupancy_monthly.csv", occ_m,
              ["year_month", "prefecture", "occupancy_pct", "source", "table"])
    write_csv("occupancy_annual.csv", occ_a,
              ["year_month", "prefecture", "occupancy_pct", "source", "table"])
    write_csv("occupancy_by_type_annual.csv", occ_t,
              ["year", "prefecture", "facility_type", "occupancy_pct", "source", "table"])
    write_csv("nights_by_type_annual.csv", nbt,
              ["year", "prefecture", "facility_type", "total_nights",
               "foreign_nights", "source", "table"])

    arr = parse_jnto()
    write_csv("arrivals_monthly.csv", arr,
              ["year_month", "nationality", "arrivals", "source"])

    pp, bc, tt = parse_spend()
    write_csv("spend_per_person.csv", pp,
              ["year", "nationality", "per_person_spend_yen", "avg_nights", "source", "note"])
    write_csv("spend_by_category.csv", bc,
              ["year", "nationality", "category", "per_person_spend_yen",
               "share_pct", "source", "table"])
    write_csv("spend_total_official.csv", tt,
              ["year", "total_spend_oku_yen", "source", "note"])

    print("\n==== 検証: 外国人延べ宿泊者数 (全国・年計, 全施設) ====")
    for r in gn_a:
        if r["prefecture"] == "全国" and r["year_month"] in ("2015", "2019", "2023", "2024", "2025"):
            print(r["year_month"], f"total={r['total_nights']:,} foreign={r['foreign_nights']:,}")
    print("公表値 (国交省資料): 2023年 11,775万人泊 / 2024年 16,446万人泊")
    print("\n==== 検証: 月次の年合計 vs 年計 (外国人, 全国) ====")
    s = defaultdict(int)
    for r in gn_m:
        if r["prefecture"] == "全国" and r["foreign_nights"]:
            s[r["year_month"][:4]] += r["foreign_nights"]
    a = {r["year_month"]: r["foreign_nights"] for r in gn_a if r["prefecture"] == "全国"}
    for y in sorted(a):
        print(y, f"monthly_sum={s[y]:,} annual={a[y]:,} diff={s[y] - a[y]:,}")
    print("\n==== 検証: JNTO 訪日外客数 総数 (年計) ====")
    s2 = defaultdict(int)
    for r in arr:
        if r["nationality"] == "総数":
            s2[r["year_month"][:4]] += r["arrivals"]
    for y in sorted(s2):
        print(y, f"{s2[y]:,}")


if __name__ == "__main__":
    main()
