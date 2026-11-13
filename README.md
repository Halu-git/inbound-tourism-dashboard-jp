# インバウンド観光ダッシュボード（都道府県×国籍）

訪日インバウンドの「人数（arrivals）→ 宿泊（nights）→ 消費（spend）」を、公式統計だけで都道府県別・国籍別に追うダッシュボードのサンプルです。Excel完成品、Power BI設計パック、整形済みCSVを含みます。

解説記事（Zenn）： https://zenn.dev/halu_git/articles/07614fe48a03d4

## 主な数値（2025年）

- 外国人延べ宿泊者数： 179,922,130人泊（前年比 +9.4%、2019年比 155.6%）／ 外国人比率 27.2%
- 国籍別（従業者数10人以上の施設）： 中国 19.7%・台湾 12.9%・韓国 11.5%・米国 11.2%
- 都道府県別： 東京都 60,115,730人泊（外国人比率 55.7%）、大阪府、京都府が続く
- 客室稼働率（全国）： 61.6%
- 訪日外客数（JNTO）： 42,683,301人（+15.8%）
- 旅行消費額： 9兆4,549億円（+16.4%）／ 1人当たり 228,782円

## 構成

- `excel/inbound_tourism_dashboard.xlsx` ― Excel完成品（12シート、ネイティブグラフ）
- `excel/inbound_tourism_dashboard_en.xlsx` ― 英語版Excel（数値は日本語版と同一）
- `powerbi/` ― Power BI設計パック（model.md / measures.dax / pages.md）
- `data/processed/` ― tidy形式CSV 12ファイル（出典列つき）
- `scripts/` ― 取得スクリプト（download_sources.py）と整形スクリプト（build_tidy.py）、Excel生成（build_excel.py）
- `design.md` ― 設計メモ（目的・検証結果・ページ設計・限界）
- 生データ（約40MB）は容量の都合で同梱していません。`scripts/download_sources.py` で公式ページから再取得できます

## データソースと採用理由

| データ | 採用理由 |
|---|---|
| 観光庁「宿泊旅行統計調査」(政府統計コード 00601020) | 宿泊実績を都道府県×月次で追える唯一の公式統計。外国人比率・稼働率・施設タイプ別まで一貫した定義で比較できる |
| JNTO「訪日外客数（国籍/月別）」 | 「入国の量」を表す代表指標。国籍別・月次の粒度が宿泊統計と揃い、上流指標として対比できる |
| 観光庁「インバウンド消費動向調査」暦年確報（2019・2024・2025年） | 国籍別の消費単価・費目別内訳・平均泊数が取れる唯一の公式調査。総額は報道発表で裏取りできる |

取得日： いずれも 2026-11-13。いずれも公開ファイルのみを使用（APIキー不使用）。

### 定義の注意

- 訪日外客数＝人数、延べ宿泊者数＝延べ泊数、旅行消費額＝アンケートに基づく推計値。単位も母集団も異なります
- 国籍別泊数（参考第1表）は従業者数10人以上の施設のみが対象で、全施設ベースの外国人総数とは母集団が異なります
- 2026年9月は第1次速報（主要表のみ）。国籍別・稼働率の9月分は未取得です

## 採用しなかったデータ

- **都道府県の観光入込客統計**： 自治体ごとに算出基準（数え方）が異なるため都道府県間の比較可能性が低く、採用していません。比較可能な公式統計に限定しました。商品化の際は、各自治体公表値をそのままの定義で集めた「参考資料（都道府県間の比較はできない旨の注記つき）」として有料オプションで提供する想定です（本リポジトリには収録していません）
- 出入国管理統計／空港別国際旅客数： 未取得（サンプル範囲外）
- 消費動向調査の四半期速報・2020〜2023年暦年値： コロナ禍の調査休止・再開で系列が不連続なため未収録
- 市区町村単位の宿泊データ： 公表自体がありません（都道府県が最細粒度）

## 検証

公表値との突合を実施し、全項目で一致を確認しています（外国人延べ宿泊者数 2023年 11,775万人泊・2024年 16,446万人泊との整合、JNTO月次積み上げと累計列の完全一致、消費単価・費目別の報道発表との一致など）。詳細は `design.md` と Excel の「検証ログ」シートを参照してください。

## ライセンス・出典の扱い

データの出典は各統計の公表機関（観光庁・日本政府観光局）に帰属します。政府統計のデータは各機関の利用条件に従ってください。本リポジトリのコードとドキュメントはサンプルとして公開しています。

## English summary

**Inbound Tourism Dashboard — Japan, by prefecture and nationality.** A sample dashboard that follows the flow *arrivals → nights → spend* for inbound visitors to Japan, built only from official statistics. Includes a finished Excel dashboard, a Power BI design pack (data model, ~25 DAX measures, 4 page layouts), and tidy CSVs with source columns.

Key figures (2025, cross-checked against published totals):
- Foreign guest nights: **179,922,130** (+9.4% YoY; 155.6% of 2019); foreign share of all guest nights **27.2%**
- By nationality (facilities with 10+ employees): China 19.7%, Taiwan 12.9%, South Korea 11.5%, USA 11.2%
- Top prefecture: Tokyo 60,115,730 guest nights (foreign share 55.7%); national room occupancy 61.6%
- Visitor arrivals (JNTO): 42,683,301 (+15.8%); travel spend: **JPY 9.45 trillion** (+16.4%), JPY 228,782 per visitor

Data sources and why they were chosen:
- Japan Tourism Agency, *Accommodation Travel Statistics Survey* — the only official source tracking guest nights monthly by prefecture and nationality under one consistent definition
- JNTO visitor arrivals — the upstream "volume" indicator, same nationality/month granularity
- Japan Tourism Agency, *International Visitor Survey* (consumption) — the only official source for per-visitor spend and spend by category

Deliberately excluded: prefectural "visitor count" statistics published by local governments. Each municipality counts differently, so cross-prefecture comparison is not valid; the dashboard uses comparable official statistics only. (In client work, those local figures can be supplied separately as clearly-labelled reference material.)
