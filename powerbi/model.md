# Power BI モデル設計 ― インバウンド観光ダッシュボード

入力： `data/processed/` の tidy CSV 群（12ファイル）。Excel版と同じCSVをそのまま使う。

## テーブル構成（レイヤー分離が原則）

### L1 宿泊統計（主レイヤー）
| テーブル | 元CSV | 粒度 | 主な列 |
|---|---|---|---|
| FactNightsMonthly | guest_nights_monthly.csv | 月×都道府県（全国含む48行/月） | total_nights, foreign_nights |
| FactNightsAnnual | guest_nights_annual.csv | 年×都道府県 | total_nights, foreign_nights（年ページ・ランキング用） |
| FactNationalityMonthly | nationality_nights_monthly.csv | 月×都道府県×国籍（21区分） | foreign_nights ※従業者数10人以上の施設が対象 |
| FactNationalityAnnual | nationality_nights_annual.csv | 年×都道府県×国籍 | foreign_nights |
| FactOccupancyMonthly | occupancy_monthly.csv | 月×都道府県 | occupancy_pct（比率は値として保持、合計しない） |
| FactOccupancyAnnual | occupancy_annual.csv | 年×都道府県 | occupancy_pct |
| FactOccupancyTypeAnnual | occupancy_by_type_annual.csv | 年×都道府県×施設タイプ | occupancy_pct |
| FactNightsTypeAnnual | nights_by_type_annual.csv | 年×都道府県×施設タイプ | total_nights, foreign_nights |

### L2 JNTO（補助レイヤー・L1とリレーションで混ぜない）
| FactArrivalsMonthly | arrivals_monthly.csv | 月×国籍（「総数」行を含む） | arrivals |

### L3 消費動向（補助レイヤー・年次3時点のみ）
| FactSpendPerson | spend_per_person.csv | 年×国籍 | per_person_spend_yen, avg_nights |
| FactSpendCategory | spend_by_category.csv | 年×国籍×費目 | per_person_spend_yen, share_pct |
| FactSpendTotal | spend_total_official.csv | 年 | total_spend_oku_yen（報道発表の確報値） |

## ディメンション

- DimDate： FactNightsMonthly の year_month から生成（年月・年・月・四半期）。FactOccupancyMonthly・FactNationalityMonthly・FactArrivalsMonthly とリレーション可（月次粒度の共通軸としてのみ使う）
- DimYear： 2015〜2026（年次ファクト用）
- DimPrefecture： 全国＋47都道府県（並び順は都道府県コード順。「全国」は合計行として扱い、ランキングからは除外できるフラグ is_total を持たせる）
- DimNationality： 21区分（L1・L2・L3で名称統一済み： 「アメリカ」→「米国」、「イギリス」→「英国」に正規化済み）。FactArrivalsMonthly の「総数」はディメンションに含めず、メジャー側で分離する
- DimFacilityType： 旅館／リゾートホテル／ビジネスホテル／シティホテル／簡易宿所／会社・団体の宿泊所

## リレーション方針

- L1 内： Fact* — DimDate / DimYear / DimPrefecture / DimNationality / DimFacilityType は単方向フィルタ
- L2・L3 は L1 ファクトと**直接リレーションを張らない**。共通ディメンション（DimDate・DimNationality）経由の見た目の連動のみとし、ページ側で「別調査・別定義」であることを注記する
- occupancy_pct・share_pct は平均系メジャーで扱い、SUM しない（モデル側で集計方法を「平均」に固定するか、メジャーで明示する）

## 注意（実装時に落とし穴になる点）

- FactNationality* の母集団は「従業者数10人以上の施設」のみ。全施設ベースの FactNights* の外国人総数とは一致しない（2025年： 154,254,100 vs 179,922,130 人泊）。国籍構成比は FactNationality 内でのみ計算する
- 2026年9月は第1次速報で、FactNationalityMonthly・FactOccupancyMonthly に行が存在しない（未取得）。月次ページでは欠損として扱い、補間しない
- JNTO の「総数」は国籍別合計と一致する行として別に存在する。二重計上を避けるため、国籍別グラフでは「総数」を除外する
- L3 は 2019・2024・2025 の3時点のみ。時系列の線で結ぶ場合は点の間隔が不均等であることを注記する
