# Rail Price Import Conventions

How `extract.py` reads the rail price sheets, and the decisions agreed for each source type.
Read this before importing new files. Add every new source-specific decision here.

Last updated: 2026-09-30

## 1. General Rules (All Types)

### Source Types

Files with the same naming convention share a format and data logic.

| Type | File Name Pattern | Family Label |
|---|---|---|
| A | 欧洲自营价格表 (中欧报价表) | ZY 自营 |
| B | 欧洲非自营价格表 | FZY 非自营 |
| C | 成都（城厢站）欧线 / 重庆非自营 | CY 成渝 |
| D | NEP跨里海报价表 | KLH 跨里海 |

The 东南亚 sheets are excluded.

### Scope (Test Pairing)

- Destination: Budapest 布达佩斯.
- Origins: Chengdu 城厢, Chongqing 团结村, Xi'an 国际港, Wuhan, Changsha.
- No Jiangxi origin has a Budapest price. Wuhan and Changsha are the nearest origins to Yichun 宜春 (Jiangxi).
- More destinations are added in the script's `DESTINATIONS` list.

### Quote Date and Versions

- The quote date comes from the file name, e.g. `-9.29`.
- Copies with the same name, such as `(1)` and `(2)`, are usually updates or corrections made on the same day.
- Each quote date uses one copy per type: the copy with the **latest internal last-saved time**.
  - .xlsx: `docProps/core.xml` → `dcterms:modified`
  - .xls: OLE summary `last_saved_time` (read with `olefile`)
  - If neither is available, the file's modification time is used.
- Ties are broken by: 更新 in the name, then file modification time, then copy number.
- The download or modification time alone is not reliable. Browser re-downloads change it.
- `extract_log.txt` lists every used and ignored copy with its saved time.

### Departure Months

- A price listed for a range, such as 9-10月, counts under both months.
- If both a range listing and a single-month listing exist for the same month, the single-month listing wins.

### Departure Dates (ETD)

ETDs are worked out in this order:

1. Dates listed in the schedule text.
2. Dates taken from a weekly schedule.
3. Otherwise, the 15th of the month, marked as assumed. It is shown as a hollow dot.

### Currency

- USD comes from the sheet where given.
- For RMB-only prices, the RMB/USD rate implied by that day's own sheets is used (about 6.68–6.85 in Jun–Sep 2026).

### SOC and COC

- SOC and COC are always kept separate.
- A header that shows only "SOC", with no size, means **40HQ**.

### Container Sizes

- Headers that don't name a size default to 40HQ when a 40HQ column exists.
- 2×20ft sets are rare in rail. Keep them only when explicitly labelled.

### All-In Price

All-In = base rate + COC box pickup charge in the origin city + fees marked 不含.

- The pickup charge is parsed from the Budapest clause, with its own currency.
- The 不绕区 surcharge is optional. It is added only when that toggle is on.

### Capacity (舱位) Notices

- This is a price tracker, not a capacity tracker. All prices are kept.
- Capacity notices are **not** shown as flags on rate cards. They appear only in remarks: card Sheet Remark and Price List Details.
- Matching words: 已无舱位, 无舱位, 无仓位, 暂无舱, 售罄, 已收齐, 已收满, 暂已收满, 已受理满, 受理满, 爆仓.
- Only the schedule, price-note and destination cells set sold-out date ranges. Remark text only adds notes.
- "不绕区（无舱位）" means the no-detour option is unavailable. It does not create a sold-out range.

### Departure Date Rule

This rule is used by Lookup and by the Timeline's Departure Date mode.

- A departure is priced at the latest quote published on or before its ETD.
- If a sold-out notice for that departure was published first, the last quote published before the notice is used.

### Publish Date Mode (Timeline)

- The x-axis is the quote date.
- There is one line per route per departure month.

## 2. Type A: 欧洲自营价格表 (ZY)

| # | Topic | Decision | Status |
|---|---|---|---|
| A1 | Duplicate copies | Use the latest internal saved time (see General Rules). | Done |
| A2 | Wuhan "自营已收齐" row | Awaiting the user's comment. | Open |
| A3 | Capacity notices | Keep all prices. Note capacity dates in remarks only. No flags. | Done |
| A4 | 华南门到站 | Keep it as a separate product. Its price makeup isn't comparable with station-to-station products. | Done |
| A5 | Headers with only "SOC" | Mean 40HQ. | Done |
| A6+ | More Type A comments | The user has more comments. | Pending |

## 3. Type B: 欧洲非自营价格表 (FZY)

To review.

## 4. Type C: 成都（城厢站）欧线 / 重庆非自营 (CY)

To review.

## 5. Type D: NEP跨里海报价表 (KLH)

To review.
