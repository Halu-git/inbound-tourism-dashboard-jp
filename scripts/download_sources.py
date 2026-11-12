#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サンプル第4弾「インバウンド観光ダッシュボード」— 公開データ取得スクリプト

すべての出典は公開ファイルのみ（APIアプリケーションIDは不使用）。
取得日: 2026-11-13（初回取得）。再実行すると data/raw/ に同じファイルを取得し直す。

出典一覧:
  主レイヤー: 観光庁「宿泊旅行統計調査」(政府統計コード 00601020、e-Stat登録統計。
    集計結果Excelは観光庁公式ページ https://www.mlit.go.jp/kankocho/tokei_hakusyo/shukuhakutokei.html
    で公表されている正本ファイル)
    - 年の確定値 集計結果 (2015〜2025年): 月次12か月分＋年計の全統計表を含む
    - 2026年 月次 第2次速報 (1〜8月分)、第1次速報 (9月分)
    - 推移表、調査の概要 (対象範囲の確認用)
  補助レイヤー1: 日本政府観光局 (JNTO)「国籍/月別 訪日外客数 (2003年〜2026年)」
    https://www.jnto.go.jp/statistics/data/visitors-statistics/ から取得
  補助レイヤー2: 観光庁「インバウンド消費動向調査」(旧・訪日外国人消費動向調査)
    暦年確報 集計表 (2019・2024・2025年、旧形式 .xls は LibreOffice で .xlsx 化した
    *_converted.xlsx を整形に使用) と、各年の報道発表資料 (総消費額の裏取り用)
  裏取り用: 国土交通省・観光庁「インバウンドの状況」資料 (2025年実績の概数確認)
"""
import pathlib
import urllib.request

RAW = pathlib.Path(__file__).resolve().parent.parent / "data" / "raw"
KANKO = "https://www.mlit.go.jp/kankocho/content/"

FILES = [
    # 宿泊旅行統計調査 年の確定値
    (KANKO + "002026815.xlsx", "annual_2015.xlsx"),
    (KANKO + "002026813.xlsx", "annual_2016.xlsx"),
    (KANKO + "002026811.xlsx", "annual_2017.xlsx"),
    (KANKO + "002026795.xlsx", "annual_2018.xlsx"),
    (KANKO + "002026792.xlsx", "annual_2019.xlsx"),
    (KANKO + "002026786.xlsx", "annual_2020.xlsx"),
    (KANKO + "002026777.xlsx", "annual_2021.xlsx"),
    (KANKO + "002026772.xlsx", "annual_2022.xlsx"),
    (KANKO + "002026766.xlsx", "annual_2023.xlsx"),
    (KANKO + "002026764.xlsx", "annual_2024.xlsx"),
    (KANKO + "002026762.xlsx", "annual_2025.xlsx"),
    # 宿泊旅行統計調査 2026年 月次速報
    (KANKO + "002026744.xlsx", "m2026_01.xlsx"),
    (KANKO + "002026742.xlsx", "m2026_02.xlsx"),
    (KANKO + "002026740.xlsx", "m2026_03.xlsx"),
    (KANKO + "002026733.xlsx", "m2026_04.xlsx"),
    (KANKO + "002026728.xlsx", "m2026_05.xlsx"),
    (KANKO + "002026747.xlsx", "m2026_06.xlsx"),
    (KANKO + "002028777.xlsx", "m2026_07.xlsx"),
    (KANKO + "002028775.xlsx", "m2026_08.xlsx"),
    (KANKO + "002028774.xlsx", "m2026_09_prelim.xlsx"),  # 第1次速報 (主要表のみ)
    # 推移表・調査概要
    (KANKO + "002028776.xlsx", "trend_table.xlsx"),
    (KANKO + "002025596.pdf", "survey_overview.pdf"),
    # JNTO 訪日外客数 (国籍/月別)
    ("https://www.jnto.go.jp/statistics/data/_files/20261021_1615-5.xlsx",
     "jnto_arrivals_monthly_by_nationality.xlsx"),
    # インバウンド消費動向調査 暦年確報 集計表 (旧形式 .xls)
    (KANKO + "002020401.xls", "consumption_2019.xls"),
    (KANKO + "002022649.xls", "consumption_2024.xls"),
    (KANKO + "002025515.xls", "consumption_2025.xls"),
    # 消費額 総額の裏取り用 報道発表 (暦年確報の概要)
    (KANKO + "002020396.pdf", "consumption_press_2019.pdf"),
    (KANKO + "002020210.pdf", "consumption_press_2024.pdf"),
    (KANKO + "002020046.pdf", "consumption_press_2025.pdf"),
    # 2025年実績の概数 (国交省資料)
    ("https://www.mlit.go.jp/policy/shingikai/content/001979608.pdf",
     "mlit_2025_inbound_summary.pdf"),
]


def main() -> None:
    RAW.mkdir(parents=True, exist_ok=True)
    for url, name in FILES:
        dest = RAW / name
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=300) as r:
            dest.write_bytes(r.read())
        print(f"{name}: {dest.stat().st_size:,} bytes")
    print("注意: consumption_YYYY_converted.xlsx は consumption_YYYY.xls を")
    print("LibreOffice ヘッドレス変換したもので、build_tidy.py の入力に使う。")


if __name__ == "__main__":
    main()
