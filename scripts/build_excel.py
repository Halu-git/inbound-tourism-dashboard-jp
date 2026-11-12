#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サンプル第4弾「インバウンド観光ダッシュボード」— Excel完成品の生成

data/processed/ の tidy CSV から output/inbound_tourism_dashboard.xlsx を作る。
構成 (12シート): 表紙・使い方 / D1_KPI概要 / D2_国籍別分析 / D3_都道府県別分析 /
D4_月次トレンド / data_宿泊_年次 / data_国籍別_年次 / data_月次_全国 /
data_消費 / data_費目別消費 / 検証ログ / 出典・注記
ダッシュボードシートのグラフは openpyxl ネイティブグラフ (同一シート内の
ソース表を参照)。出典は各図表の下に注記する。
"""
import csv
import pathlib
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
PROC = ROOT / "data" / "processed"
OUTF = ROOT / "output" / "inbound_tourism_dashboard.xlsx"

NAVY = "1F3864"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
KPI_LABEL_FONT = Font(bold=True, size=11, color="404040")
KPI_VALUE_FONT = Font(bold=True, size=16, color=NAVY)
NOTE_FONT = Font(size=9, color="595959")
SRC_L1 = "観光庁 宿泊旅行統計調査"
SRC_L2 = "JNTO 訪日外客数"
SRC_L3 = "観光庁 インバウンド消費動向調査"


def rd(name):
    with open(PROC / name, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def iv(v):
    return int(float(v)) if v not in (None, "") else None


def fv(v):
    return float(v) if v not in (None, "") else None


def title(ws, cell, text):
    ws[cell] = text
    ws[cell].font = TITLE_FONT


def note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = NOTE_FONT


def hdr_row(ws, row, col0, labels):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col0 + i, value=lab)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center")


def put_table(ws, row0, col0, headers, rows, fmts=None, widths=None):
    hdr_row(ws, row0, col0, headers)
    for j, r in enumerate(rows):
        for i, v in enumerate(r):
            c = ws.cell(row=row0 + 1 + j, column=col0 + i, value=v)
            if fmts and fmts[i]:
                c.number_format = fmts[i]
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(col0 + i)].width = w
    return row0 + 1 + len(rows)


def kpi(ws, row, col, label, value, fmt="#,##0", sub=None):
    ws.cell(row=row, column=col, value=label).font = KPI_LABEL_FONT
    c = ws.cell(row=row + 1, column=col, value=value)
    c.font = KPI_VALUE_FONT
    c.number_format = fmt
    if sub:
        ws.cell(row=row + 2, column=col, value=sub).font = NOTE_FONT


# ------------------------------------------------------------------ データ準備
gn_a = rd("guest_nights_annual.csv")
gn_m = rd("guest_nights_monthly.csv")
nat_a = rd("nationality_nights_annual.csv")
occ_a = rd("occupancy_annual.csv")
occ_m = rd("occupancy_monthly.csv")
arr_m = rd("arrivals_monthly.csv")
spend_pp = rd("spend_per_person.csv")
spend_cat = rd("spend_by_category.csv")
spend_tot = rd("spend_total_official.csv")

YEARS = [str(y) for y in range(2015, 2026)]
nat_series = {}   # 全国・年 -> {国籍: 泊数}
for r in nat_a:
    if r["prefecture"] == "全国":
        nat_series.setdefault(r["year_month"], {})[r["nationality"]] = iv(r["foreign_nights"])
jp = {r["year_month"]: r for r in gn_a if r["prefecture"] == "全国"}
occ_jp = {r["year_month"]: fv(r["occupancy_pct"]) for r in occ_a if r["prefecture"] == "全国"}
arr_year = defaultdict(int)
arr_month = {}
for r in arr_m:
    if r["nationality"] == "総数":
        arr_year[r["year_month"][:4]] += iv(r["arrivals"])
        arr_month[r["year_month"]] = iv(r["arrivals"])

F25 = iv(jp["2025"]["foreign_nights"])
F24 = iv(jp["2024"]["foreign_nights"])
F19 = iv(jp["2019"]["foreign_nights"])
T25 = iv(jp["2025"]["total_nights"])
SPEND25 = 94549  # 億円 (報道発表 確報)
SPEND24 = 81257

wb = Workbook()

# ------------------------------------------------------------------ 表紙
ws = wb.active
ws.title = "表紙・使い方"
ws.column_dimensions["B"].width = 110
ws["B2"] = "インバウンド観光ダッシュボード（都道府県×国籍）"
ws["B2"].font = Font(bold=True, size=18, color=NAVY)
lines = [
    "",
    "訪日インバウンドの「人数（arrivals）→ 宿泊（nights）→ 消費（spend）」を、公式統計だけで追うダッシュボードです。",
    "主データは観光庁「宿泊旅行統計調査」（都道府県別・国籍別の延べ宿泊者数、客室稼働率）。",
    "補助レイヤーとして JNTO「訪日外客数」と観光庁「インバウンド消費動向調査」を、混在させず別シート・別表で併記します。",
    "",
    "シート構成：",
    "・D1_KPI概要 … 外国人延べ宿泊者数・外国人比率・稼働率と、arrivals→nights→spend の補助KPI",
    "・D2_国籍別分析 … 国籍別の推移・構成比・2019年比回復率、消費単価と平均泊数",
    "・D3_都道府県別分析 … 47都道府県の外国人泊数・比率・稼働率、施設タイプ別、国籍ミックスの地域差",
    "・D4_月次トレンド … 月次の季節性（2019／2024／2025／2026）と直近の動向",
    "・data_* … 整形済みデータ（出典列つき）",
    "・検証ログ … 公表値との突合結果",
    "・出典・注記 … 出典、定義、採用理由・不採用理由、限界",
    "",
    "読み方の注意：",
    "・訪日外客数＝人数、延べ宿泊者数＝延べ泊数、旅行消費額＝アンケートに基づく推計値。単位も母集団も異なるため、",
    "　割り算で作った指標は「参考指標」として定義差の注記つきで示しています。",
    "・国籍別の泊数（参考第1表）は従業者数10人以上の施設が対象で、全施設ベースの外国人総数とは母集団が異なります。",
    "・2026年9月は第1次速報（主要表のみ）。国籍別・稼働率の9月分は未取得です。",
    "",
    "データ取得日： 2026-11-13 ／ 出典の詳細は「出典・注記」シート参照",
]
for i, t in enumerate(lines):
    ws.cell(row=4 + i, column=2, value=t)

# ------------------------------------------------------------------ D1 KPI概要
ws = wb.create_sheet("D1_KPI概要")
for col in "BCDEFG":
    ws.column_dimensions[col].width = 24
title(ws, "B2", "D1 KPI概要 ― 外国人延べ宿泊者数と arrivals → nights → spend（2025年）")
kpi(ws, 4, 2, "外国人延べ宿泊者数（2025年）", F25, "#,##0", f"前年比 {(F25/F24-1)*100:+.1f}% ／ 2019年比 {F25/F19*100:.1f}%")
kpi(ws, 4, 4, "外国人比率（2025年）", fv(jp["2025"]["foreign_share_pct_calc"]) / 100, "0.0%", f"総延べ宿泊者数 {T25:,} 人泊")
kpi(ws, 4, 6, "客室稼働率 全国（2025年）", occ_jp["2025"] / 100, "0.0%", "第8表（年計）")
kpi(ws, 8, 2, "訪日外客数 arrivals（2025年・JNTO）", arr_year["2025"], "#,##0", f"前年比 {(arr_year['2025']/arr_year['2024']-1)*100:+.1f}%・過去最高")
kpi(ws, 8, 4, "外国人延べ宿泊者数 nights（2025年）", F25, "#,##0", "宿泊旅行統計調査 第2表")
kpi(ws, 8, 6, "旅行消費額 spend（2025年）", SPEND25 / 10000, "0.00", "単位：兆円（9兆4,549億円・確報）")
note(ws, "B12", "流れの読み方: 入国した人数（L2 JNTO）→ 実際に泊まった延べ泊数（L1 宿泊統計）→ 使った金額（L3 消費動向調査・推計値）。単位と母集団が異なるため単純な割り算は参考指標です。")
note(ws, "B13", f"参考指標: 外国人延べ宿泊者数 ÷ 訪日外客数 = {F25/arr_year['2025']:.2f}（2025年）。消費動向調査の平均泊数 9.5泊とは定義・母集団が異なり単純比較できません（出典・注記参照）。")
note(ws, "B14", f"参考指標: 1人当たり旅行消費単価 228,782円（2025年・全国籍、L3 参考1）。消費額の前年比は {(SPEND25/SPEND24-1)*100:+.1f}% です。")

# 年次推移ソース表 (H列)
rows = [[int(y), iv(jp[y]["total_nights"]), iv(jp[y]["foreign_nights"]),
         fv(jp[y]["foreign_share_pct_calc"])] for y in YEARS]
end = put_table(ws, 4, 9, ["年", "総延べ宿泊者数", "外国人延べ宿泊者数", "外国人比率(%)"],
                rows, fmts=[None, "#,##0", "#,##0", "0.0"], widths=[8, 18, 20, 14])
ch = LineChart()
ch.title = "外国人延べ宿泊者数の推移（全国・2015〜2025年）"
ch.add_data(Reference(ws, min_col=11, min_row=4, max_row=end - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=5, max_row=end - 1))
ch.y_axis.numFmt = '#,##0'
ch.height, ch.width = 7.5, 15
ws.add_chart(ch, "B16")
note(ws, "B32", f"出典: {SRC_L1} 第2表（年計）。外国人比率は算出値（外国人÷総数）。")

# 国籍TOP8 (2025)
top8 = sorted(nat_series["2025"].items(), key=lambda x: -(x[1] or 0))[:8]
base25 = sum(v for v in nat_series["2025"].values() if v)
r0 = 34
end = put_table(ws, r0, 9, ["国籍", "外国人延べ宿泊者数", "構成比(%)"],
                [[k, v, round(v / base25 * 100, 1)] for k, v in top8],
                fmts=[None, "#,##0", "0.0"])
ch = BarChart()
ch.type = "bar"
ch.title = "国籍別 外国人延べ宿泊者数 TOP8（2025年）"
ch.add_data(Reference(ws, min_col=10, min_row=r0, max_row=end - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=r0 + 1, max_row=end - 1))
ch.height, ch.width = 7.5, 15
ws.add_chart(ch, "B34")
note(ws, "B50", f"出典: {SRC_L1} 参考第1表（年計）。従業者数10人以上の施設が対象のため、全施設ベースの外国人総数（{F25:,}人泊）とは母集団が異なります。")

# 都道府県TOP10 (2025)
pref25 = sorted(
    [(r["prefecture"], iv(r["foreign_nights"])) for r in gn_a
     if r["year_month"] == "2025" and r["prefecture"] != "全国"],
    key=lambda x: -x[1])[:10]
r0 = 52
end = put_table(ws, r0, 9, ["都道府県", "外国人延べ宿泊者数"], pref25, fmts=[None, "#,##0"])
ch = BarChart()
ch.type = "bar"
ch.title = "都道府県別 外国人延べ宿泊者数 TOP10（2025年）"
ch.add_data(Reference(ws, min_col=10, min_row=r0, max_row=end - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=r0 + 1, max_row=end - 1))
ch.height, ch.width = 7.5, 15
ws.add_chart(ch, "B52")
note(ws, "B68", f"出典: {SRC_L1} 第2表（年計）。東京都が全体の {pref25[0][1]/F25*100:.1f}% を占めます。")

# ------------------------------------------------------------------ D2 国籍別
ws = wb.create_sheet("D2_国籍別分析")
title(ws, "B2", "D2 国籍別分析 ― 推移・構成比・2019年比回復率・消費単価")
MAIN_NAT = ["中国", "台湾", "韓国", "米国", "香港", "オーストラリア"]
rows = [[int(y)] + [nat_series[y].get(n) for n in MAIN_NAT] for y in YEARS]
end = put_table(ws, 4, 9, ["年"] + MAIN_NAT, rows, fmts=["#,##0"] * 7)
ch = LineChart()
ch.title = "主要国籍の外国人延べ宿泊者数 推移（2015〜2025年）"
for ci in range(10, 16):
    ch.add_data(Reference(ws, min_col=ci, min_row=4, max_row=end - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=5, max_row=end - 1))
ch.height, ch.width = 8, 16
ws.add_chart(ch, "B4")
note(ws, "B21", f"出典: {SRC_L1} 参考第1表（従業者数10人以上の施設）。2020〜2022年はコロナ禍で激減しています。")

# 回復率
rec_nat = ["中国", "台湾", "韓国", "米国", "香港", "タイ", "オーストラリア", "英国"]
rec_rows = [[n, nat_series["2019"].get(n), nat_series["2025"].get(n),
             round(nat_series["2025"][n] / nat_series["2019"][n] * 100, 1)]
            for n in rec_nat]
r0 = 23
end = put_table(ws, r0, 9, ["国籍", "2019年", "2025年", "2025/2019 (%)"], rec_rows,
                fmts=[None, "#,##0", "#,##0", "0.0"])
ch = BarChart()
ch.type = "bar"
ch.title = "2019年比の回復率（2025年・国籍別）"
ch.add_data(Reference(ws, min_col=12, min_row=r0, max_row=end - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=r0 + 1, max_row=end - 1))
ch.height, ch.width = 7.5, 15
ws.add_chart(ch, "B23")
note(ws, "B39", "出典: 同上。韓国は2019年後半の日韓関係悪化で当時の水準が低く、回復率が高く出ています。米国は236.6%と突出しています。")

# 構成比の推移 (主要5 + その他)
comp_nat = ["中国", "台湾", "韓国", "米国", "香港"]
comp_rows = []
for y in YEARS:
    tot = sum(v for v in nat_series[y].values() if v)
    vals = [nat_series[y].get(n) or 0 for n in comp_nat]
    comp_rows.append([int(y)] + vals + [tot - sum(vals)])
r0 = 41
end = put_table(ws, r0, 9, ["年"] + comp_nat + ["その他"], comp_rows, fmts=["#,##0"] * 7)
ch = BarChart()
ch.type = "col"
ch.grouping = "percentStacked"
ch.title = "国籍構成比の推移（2015〜2025年）"
for ci in range(10, 16):
    ch.add_data(Reference(ws, min_col=ci, min_row=r0, max_row=end - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=r0 + 1, max_row=end - 1))
ch.height, ch.width = 8, 16
ws.add_chart(ch, "B41")
note(ws, "B58", f"出典: {SRC_L1} 参考第1表。中国の構成比は2015年の約3割から2025年 {nat_series['2025']['中国']/base25*100:.1f}% へ低下し、国籍の多様化が進んでいます。")

# 消費単価 × 平均泊数 (2025)
sp25 = {r["nationality"]: r for r in spend_pp if r["year"] == "2025"}
sp_rows = []
for n, r in sorted(sp25.items(), key=lambda x: -(fv(x[1]["per_person_spend_yen"]) or 0)):
    if n == "全国籍･地域":
        continue
    sp_rows.append([n, fv(r["per_person_spend_yen"]), fv(r["avg_nights"])])
r0 = 60
end = put_table(ws, r0, 9, ["国籍", "1人当たり消費単価(円)", "平均泊数(泊)"],
                sp_rows[:10], fmts=[None, "#,##0", "0.0"])
ch = BarChart()
ch.type = "bar"
ch.title = "国籍別 1人当たり旅行消費単価 TOP10（2025年）"
ch.add_data(Reference(ws, min_col=10, min_row=r0, max_row=end - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=r0 + 1, max_row=end - 1))
ch.height, ch.width = 7.5, 15
ws.add_chart(ch, "B60")
note(ws, "B76", f"出典: {SRC_L3} 参考1・表4-1（2025年暦年確報、試算値）。単価は宿泊統計の泊数とは別調査の値です。中国 246,550円が最高水準、韓国 105,058円は平均泊数 4.0泊と短めです。")

# 費目別 (全国籍 2025)
cat25 = [(r["category"], fv(r["per_person_spend_yen"]), fv(r["share_pct"]))
         for r in spend_cat if r["year"] == "2025" and r["nationality"] == "全国籍･地域"]
r0 = 78
end = put_table(ws, r0, 9, ["費目", "1人当たり(円)", "構成比(%)"], cat25,
                fmts=[None, "#,##0", "0.0"])
note(ws, f"B{r0 + 9}", f"出典: {SRC_L3} 参考2（2025年）。宿泊費 84,057円（36.7%）が最大で、買物代 61,136円（26.7%）、飲食費 50,221円（22.0%）と続きます。")

# ------------------------------------------------------------------ D3 都道府県
ws = wb.create_sheet("D3_都道府県別分析")
title(ws, "B2", "D3 都道府県別分析 ― ランキング・外国人比率・稼働率・国籍ミックス（2025年）")
occ25 = {r["prefecture"]: fv(r["occupancy_pct"]) for r in occ_a if r["year_month"] == "2025"}
pref_rows = []
for r in gn_a:
    if r["year_month"] == "2025" and r["prefecture"] != "全国":
        pref_rows.append([r["prefecture"], iv(r["total_nights"]), iv(r["foreign_nights"]),
                          fv(r["foreign_share_pct_calc"]), occ25.get(r["prefecture"])])
pref_rows.sort(key=lambda x: -x[2])
end = put_table(ws, 4, 2, ["都道府県", "総延べ宿泊者数", "外国人延べ宿泊者数", "外国人比率(%)", "客室稼働率(%)"],
                pref_rows, fmts=[None, "#,##0", "#,##0", "0.0", "0.0"],
                widths=[14, 18, 20, 14, 14])
note(ws, f"B{end + 1}", f"出典: {SRC_L1} 第2表・第8表（2025年 年計）。外国人比率は算出値。")

r0 = 4
top15 = [[p, v] for p, _, v, _, _ in pref_rows[:15]]
end2 = put_table(ws, r0, 9, ["都道府県", "外国人延べ宿泊者数"], top15, fmts=[None, "#,##0"])
ch = BarChart()
ch.type = "bar"
ch.title = "外国人延べ宿泊者数 TOP15（2025年）"
ch.add_data(Reference(ws, min_col=10, min_row=r0, max_row=end2 - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=r0 + 1, max_row=end2 - 1))
ch.height, ch.width = 9, 16
ws.add_chart(ch, "H22")
note(ws, "H39", "出典: 同上。東京・大阪・京都の3都府県で外国人泊数の約半数を占めます。")

occ_top = sorted(occ25.items(), key=lambda x: -(x[1] or 0))[:10]
r0 = 41
end3 = put_table(ws, r0, 9, ["都道府県", "客室稼働率(%)"], [[k, v] for k, v in occ_top], fmts=[None, "0.0"])
ch = BarChart()
ch.type = "bar"
ch.title = "客室稼働率 TOP10（2025年）"
ch.add_data(Reference(ws, min_col=10, min_row=r0, max_row=end3 - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=r0 + 1, max_row=end3 - 1))
ch.height, ch.width = 7.5, 15
ws.add_chart(ch, "H41")
note(ws, "H57", f"出典: {SRC_L1} 第8表。全国平均 61.6%。施設タイプ別ではシティホテル・ビジネスホテルが高く、旅館は低めです（data は tidy CSV 参照）。")

# 国籍ミックス: 東京 vs 全国 (2025, 参考第1表)
mix_nat = ["中国", "台湾", "韓国", "米国", "香港"]
tokyo = {r["nationality"]: iv(r["foreign_nights"]) for r in nat_a
         if r["year_month"] == "2025" and r["prefecture"] == "東京都"}
t_tot = sum(v for v in tokyo.values() if v)
j_tot = base25
mix_rows = []
for n in mix_nat:
    mix_rows.append([n, round(tokyo.get(n, 0) / t_tot * 100, 1),
                     round(nat_series["2025"].get(n, 0) / j_tot * 100, 1)])
mix_rows.append(["その他",
                 round((t_tot - sum(tokyo.get(n, 0) for n in mix_nat)) / t_tot * 100, 1),
                 round((j_tot - sum(nat_series["2025"].get(n, 0) for n in mix_nat)) / j_tot * 100, 1)])
r0 = 59
put_table(ws, r0, 9, ["国籍", "東京都 構成比(%)", "全国 構成比(%)"], mix_rows, fmts=[None, "0.0", "0.0"])
note(ws, f"I{r0 + 9}", f"出典: {SRC_L1} 参考第1表（2025年）。東京都は米国の比率が全国より高く、中国の比率が低いなど、国籍ミックスに地域差があります。")

# ------------------------------------------------------------------ D4 月次
ws = wb.create_sheet("D4_月次トレンド")
for col in "BCDEFG":
    ws.column_dimensions[col].width = 24
title(ws, "B2", "D4 月次トレンド ― 季節性と直近の動向")
month_rows = []
for m in range(1, 13):
    row = [m]
    for y in ("2019", "2024", "2025", "2026"):
        v = None
        for r in gn_m:
            if r["year_month"] == f"{y}-{m:02d}" and r["prefecture"] == "全国":
                v = iv(r["foreign_nights"])
        row.append(v)
    month_rows.append(row)
end = put_table(ws, 4, 9, ["月", "2019年", "2024年", "2025年", "2026年"], month_rows,
                fmts=["#,##0"] * 5)
ch = LineChart()
ch.title = "月別 外国人延べ宿泊者数（全国）― 季節性の年比較"
for ci in range(10, 14):
    ch.add_data(Reference(ws, min_col=ci, min_row=4, max_row=end - 1), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=9, min_row=5, max_row=end - 1))
ch.height, ch.width = 8, 16
ws.add_chart(ch, "B4")
note(ws, "B21", f"出典: {SRC_L1} 第2表（月次）。2026年は9月まで（9月は第1次速報）。春（3〜4月）と秋（10月）にピークが出る二山型が戻っています。")

ytd26 = sum(iv(r["foreign_nights"]) for r in gn_m
            if r["year_month"].startswith("2026") and r["prefecture"] == "全国" and r["foreign_nights"])
ytd25 = sum(iv(r["foreign_nights"]) for r in gn_m
            if "2025-01" <= r["year_month"] <= "2025-09" and r["prefecture"] == "全国" and r["foreign_nights"])
kpi(ws, 23, 2, "外国人延べ宿泊者数 2026年1〜9月 累計", ytd26, "#,##0",
    f"前年同期比 {(ytd26/ytd25-1)*100:+.1f}%（2025年1〜9月 {ytd25:,} 人泊）")
arr_ytd26 = sum(v for k, v in arr_month.items() if k.startswith("2026"))
arr_ytd25 = sum(v for k, v in arr_month.items() if "2025-01" <= k <= "2025-09")
kpi(ws, 23, 5, "訪日外客数 2026年1〜9月 累計（JNTO）", arr_ytd26, "#,##0",
    f"前年同期比 {(arr_ytd26/arr_ytd25-1)*100:+.1f}%")
note(ws, "B27", "注: 2026年9月の国籍別（参考第1表）・客室稼働率（第8表）は第1次速報に含まれず未取得です。第2次速報で補完予定。")

# ------------------------------------------------------------------ data系
def dump_sheet(name, rows_dict, cols):
    s = wb.create_sheet(name)
    hdr_row(s, 1, 1, cols)
    for j, r in enumerate(rows_dict):
        for i, c in enumerate(cols):
            s.cell(row=2 + j, column=1 + i, value=r.get(c))
    return s

dump_sheet("data_宿泊_年次", gn_a,
           ["year_month", "prefecture", "total_nights", "foreign_nights",
            "foreign_share_pct_calc", "source", "table"])
dump_sheet("data_国籍別_年次", [r for r in nat_a if r["prefecture"] == "全国"],
           ["year_month", "prefecture", "nationality", "foreign_nights", "source", "table", "scope"])
# 月次・全国マージ
occ_m_map = {(r["year_month"]): fv(r["occupancy_pct"]) for r in occ_m if r["prefecture"] == "全国"}
merged = []
for r in gn_m:
    if r["prefecture"] == "全国":
        merged.append({"year_month": r["year_month"], "total_nights": iv(r["total_nights"]),
                       "foreign_nights": iv(r["foreign_nights"]),
                       "foreign_share_pct_calc": fv(r["foreign_share_pct_calc"]),
                       "occupancy_pct": occ_m_map.get(r["year_month"]),
                       "arrivals_total_jnto": arr_month.get(r["year_month"]),
                       "note": "2026-09 稼働率は未取得（第1次速報）" if r["year_month"] == "2026-09" else ""})
dump_sheet("data_月次_全国", merged,
           ["year_month", "total_nights", "foreign_nights", "foreign_share_pct_calc",
            "occupancy_pct", "arrivals_total_jnto", "note"])
dump_sheet("data_消費", spend_pp,
           ["year", "nationality", "per_person_spend_yen", "avg_nights", "source", "note"])
s = wb["data_消費"]
r0 = len(spend_pp) + 3
hdr_row(s, r0, 1, ["年", "旅行消費額 総額(億円)", "出典", "注記"])
for j, r in enumerate(spend_tot):
    s.cell(row=r0 + 1 + j, column=1, value=int(r["year"]))
    c = s.cell(row=r0 + 1 + j, column=2, value=int(r["total_spend_oku_yen"]))
    c.number_format = "#,##0"
    s.cell(row=r0 + 1 + j, column=3, value=r["source"])
    s.cell(row=r0 + 1 + j, column=4, value=r["note"])
dump_sheet("data_費目別消費", spend_cat,
           ["year", "nationality", "category", "per_person_spend_yen", "share_pct", "source", "table"])

# ------------------------------------------------------------------ 検証ログ
ws = wb.create_sheet("検証ログ")
title(ws, "B2", "検証ログ ― 公表値との突合（2026-11-13 実施）")
rows = [
    ["L1 外国人延べ宿泊者数 2023年（全国・年計）", "117,751,450 人泊", "11,775万人泊（国交省・観光庁資料）", "一致（万人単位の公表値と整合）"],
    ["L1 外国人延べ宿泊者数 2024年（全国・年計）", "164,462,600 人泊", "16,446万人泊（同上）", "一致"],
    ["L1 2025年 総延べ宿泊者数", "661,105,480 人泊", "第2表 年計の公表値そのもの", "一致"],
    ["L1 月次合計 vs 年計（外国人・全国、2015〜2025）", "各年で突合", "年計", "一致（差がある年も最大±20人泊＝公表値の10人単位丸め）"],
    ["L1 第2表 vs 第4表 外国人総数（2025年）", "179,922,130 人泊", "179,922,130 人泊", "完全一致"],
    ["L2 JNTO 2025年 訪日外客数", "42,683,301 人（月次積み上げ）", "42,683,301 人（同ファイル累計列）", "完全一致"],
    ["L3 1人当たり消費単価 2025年（全国籍）", "228,782 円（参考1）", "228,782 円（報道発表）", "完全一致"],
    ["L3 費目別単価 2025年（宿泊費ほか）", "宿泊費 84,057円 等", "報道発表の内訳", "一致"],
    ["L3 消費額総額", "2025年 9兆4,549億円", "報道発表（確報）", "報道発表値を採用"],
]
end = put_table(ws, 4, 2, ["検証項目", "本ダッシュボードの値", "公表値", "結果"], rows, widths=[46, 34, 34, 40])
note(ws, f"B{end + 1}", "突合できなかった項目は推測で埋めず、未取得として「出典・注記」に記録しています。")

# ------------------------------------------------------------------ 出典・注記
ws = wb.create_sheet("出典・注記")
ws.column_dimensions["B"].width = 130
title(ws, "B2", "出典・注記 ― 定義、採用理由・不採用理由、限界")
lines = [
    ("【出典（すべて 2026-11-13 取得、公開ファイルのみ・API不使用）】", True),
    ("L1 主レイヤー: 観光庁「宿泊旅行統計調査」(政府統計コード 00601020) 第2表・参考第1表・第4表・第8表（観光庁公式ページの集計結果Excel）", False),
    ("L2 補助レイヤー: 日本政府観光局（JNTO）「国籍/月別 訪日外客数」", False),
    ("L3 補助レイヤー: 観光庁「インバウンド消費動向調査」暦年確報 集計表（2019・2024・2025年）および各年の報道発表資料", False),
    ("裏取り資料: 国土交通省・観光庁「インバウンドの状況」（2025年実績）", False),
    ("", False),
    ("【定義】", True),
    ("・延べ宿泊者数: 宿泊施設に泊まった人数×泊数の延べ数。1人が3泊すれば3人泊。複数施設・複数都道府県に泊まればそれぞれ数える。", False),
    ("・宿泊旅行統計調査は標本調査で、公表値は回収率による割り戻し集計を含む。第2表・第4表・第8表は全施設ベース。", False),
    ("・国籍別（参考第1表）は従業者数10人以上の施設のみが対象。全施設ベースの外国人総数とは母集団が異なる。", False),
    ("・訪日外客数（JNTO）は入国者数（人数）。延べ宿泊者数（泊数）・旅行消費額（推計値）とは単位も母集団も異なる。", False),
    ("・旅行消費額・消費単価はアンケート調査に基づく試算値（出入国管理統計による割り戻し集計）。", False),
    ("・外国人比率は本ダッシュボードの算出値（外国人延べ宿泊者数÷総延べ宿泊者数）。", False),
    ("・2026年1〜8月は第2次速報、9月は第1次速報。速報値は後日改定されうる。公表値の '*' は秘匿等で、数値として読み取っている。'-' は欠損。", False),
    ("", False),
    ("【採用理由（採用した全データセット）】", True),
    ("・宿泊旅行統計調査: 宿泊実績を都道府県×月次で追える唯一の公式統計で、外国人比率・稼働率・施設タイプ別まで一貫した定義で比較できるため。", False),
    ("・JNTO 訪日外客数: 「入国の量」を表す代表指標で、国籍別・月次の粒度が宿泊統計と揃い、上流指標として対比できるため。", False),
    ("・インバウンド消費動向調査: 国籍別の消費単価・費目別内訳・平均泊数が取れる唯一の公式調査で、総額を報道発表で裏取りできるため。", False),
    ("", False),
    ("【不採用・未取得（理由つき）】", True),
    ("・都道府県の観光入込客統計【不採用】: 自治体ごとに算出基準（数え方）が異なるため都道府県間の比較可能性が低く、本ダッシュボードでは採用しない。比較可能な公式統計（宿泊旅行統計・JNTO・消費動向調査）に限定した。", False),
    ("　→ 商品化時は、各自治体公表値をそのままの定義で集めた「参考資料（都道府県間の比較はできない旨の注記つき）」として有料オプション提供の想定。取得・整形は商品化フェーズの作業で、本サンプルには収録しない。（本ダッシュボードの「採用／不採用を理由つきで明記する」運用ルールの適用例）", False),
    ("・出入国管理統計（法務省）／空港別国際旅客数【未取得】: 本サンプルの範囲では取得・整形まで至らなかった。商品化時に必要なら追加する。", False),
    ("・消費動向調査の四半期速報・2020〜2023年暦年値【未収録】: コロナ禍の調査休止・再開で系列が不連続な期間を含むため、暦年確報の3時点（2019・2024・2025）に限定した。", False),
    ("・2026年9月の国籍別・稼働率【未取得】: 第1次速報（主要表のみ）には当該表が含まれない。第2次速報で補完可能な見込み。", False),
    ("・市区町村単位の宿泊データ【存在しない】: 宿泊旅行統計調査の公表粒度は都道府県まで。", False),
    ("", False),
    ("【限界】", True),
    ("・宿泊統計は標本調査ベースの推計を含み、速報値は改定されうる。", False),
    ("・消費動向調査の3時点比較は、調査設計の変更（コロナ禍の休止・再開）をまたぐため、単価の時系列比較は幅を持って読むこと。", False),
    ("・訪日外客数の総数には乗員等の区分を含む（JNTO公表の総数ベース）。", False),
]
r = 4
for text, bold in lines:
    c = ws.cell(row=r, column=2, value=text)
    if bold:
        c.font = Font(bold=True, size=11, color=NAVY)
    r += 1

OUTF.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUTF)
print("saved:", OUTF)
