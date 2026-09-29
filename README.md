# Box Rate Desk — container price tool

Turns Jeff's weekly quotation emails into a searchable rate desk.

    Rail & Container Price Portal\Container\
    ├── Container Price Emails\        <- the source emails
    │   ├── 20260908 Export\           (446 .eml + Jeff's index csv)
    │   └── 20260929 Export\           (hand-exported, no id prefix in the filenames)
    └── container-price-tool\          <- this folder

`extract.py` walks `Container Price Emails\` recursively, so a new dated export folder is
picked up with no change needed — just drop it in and re-run.

## Refreshing after new emails arrive

Drop the new `.eml` files into a folder under `Container Price Emails\`, then run, from
this folder:

    python3 extract.py     # parses new .eml files, writes prices.csv / china.csv / notices.csv
    python3 build.py       # rebuilds rate-desk.html and artifact-body.html

`extract.py` caches each parsed email in `.cache/`, so it only reads files that are new or
changed — a normal weekly refresh takes seconds. A cold run (or one after the parser
changes) exceeds the shell's time limit and stops with `PARTIAL: ... run again`; just run it
again until it prints the summary. Set `EXTRACT_BUDGET` to change that cut-off.

Then ask Claude to republish, or open `rate-desk.html` directly. Adding older emails needs
nothing special — they slot into the timeline by the date inside each quotation table.

## How the Page Is Organised

Three sections, split by what each one is scoped to — they are not four views of the same
thing, so they no longer sit in one flat row:

| Section | Scope | Filters it takes |
|---|---|---|
| **Route Rates** | one selected route | corridor · on-hire · off-hire · box · grade · week |
| **Corridor** | every route in a corridor, one week | corridor · box · grade · week |
| **China Depots** | domestic buy/sell by depot — a different dataset | box · week |

Route Rates holds two views of the same selected route, switched inside the section:
**Lookup** (that route, that week) and **Timeline** (pinned routes over time). The filter
rail belongs to whichever section is open, which is why it changes shape between them.

Clicking a cell in the Corridor matrix loads that route and jumps to Route Rates → Lookup.
That is the one deliberate link between sections.

## Real Terms (Inflation Adjustment)

The Timeline and China Depots charts can restate every week in one year's money:

    real = nominal × (1 + r) ^ (years from that week to the base anchor)

Years are fractional, measured from the actual dates — adjusting in whole calendar steps
would put false jumps on 1 January that read as rate moves.

- **Your Assumed Rate** is yours to set. There is no default index behind it: these are
  USD prices on a largely Chinese cost base, so US CPI, Chinese PPI and your own cost
  inflation each give a different answer. The 3%/yr the field opens on is a starting
  point, not a recommendation.
- **Base Year** anchors on that year's *last quoted week* — 2026 on 28 Sep 2026, 2025 on
  29 Dec 2025, and so on. The anchor date is always shown beside the control and printed
  on the chart, so it is never ambiguous what the figures are relative to.
- The faint line behind each adjusted series is the nominal one.
- **Lookup is never adjusted.** It is where a rate gets copied into a customer quote, so
  it always shows Jeff's figure exactly as quoted. The Corridor matrix is not adjusted
  either — within a single week every cell would scale by the same factor, which changes
  nothing except the chance of misreading it.

A flat compounding rate is an approximation. Adjusting against a real published index
would mean loading an actual CPI series — a separate piece of work if the flat rate
turns out too blunt.

## Files

| File | What it is |
|---|---|
| `extract.py` | Email parser. Handles both layouts: HTML tables (2024–25, and Europe/China throughout) and XLSX attachments (Russia/Central Asia from 2026). |
| `build.py` | Packs the CSVs into the page (change-encoded, ~0.5 MB). |
| `template.html` | The page itself — markup, styles, charts, navigation. `__DATA__` is replaced at build time. |
| `rate-desk.html` | Offline copy. Open it in a browser, no server needed. |
| `artifact-body.html` | The version published online. |
| `prices.csv` | Every route quote: date, corridor, on-hire, off-hire, box, grade, price, status. |
| `notices.csv` | Jeff's mid-week follow-ups: date, corridor, kind (`void` / `revised` / `note`), wording. |
| `china.csv` | Domestic depot prices: purchase (last/this week), retail. |
| `extract_log.txt` | Row counts and anything the parser could not read. Check this after a refresh. |

## Decisions the parser makes

- **Prices vs states.** `case by case` and `/` (not offered) are kept as states, not
  dropped and not treated as zero — they show as `c/c` and `—`.
- **Condition grades.** Columns like `塔什干(CW)` / `塔什干(IICL)` become the same
  destination with a `grade` of CW or IICL, so both quotes are kept.
- **Corrections supersede.** Keyed on the update date inside the table, so a quoted table
  lands on its own week; where two emails cover the same route and week, the email with the
  higher id prefix wins. This matters — replies like "更正报价，以此为准" (corrected quotation, use this as the
  standard) carry the authoritative numbers, and 10 weeks are corrected this way.
  Hand-exported batches have no id prefix; that is fine because they carry no follow-up
  corrections. If one ever does, give the correction a higher id than the original.
- **Mid-week notices.** Reply emails are also read for follow-ups that change the standing
  of a week's rates: withdrawals (`失效` / `作废` / `暂停`, e.g. the 2 Mar 2026 "no longer
  valid, confirm case by case" notice across all four series) and revisions. These show as a
  strip on the Lookup, Matrix and China tabs, and as ▲ markers on the timeline.
- **Wrong attachments.** The corridor is read from each table's own title, not the email
  subject — the 15 Jun 2026 "Russia" email carried the Central Asia workbook and is
  filed correctly.
- **Column layouts vary.** Column roles (on-hire, box, available stock, container
  conditions, destinations) are located per table rather than assumed, so one-off shapes
  parse — e.g. the 24 Aug 2026 Baku correction, which has no stock column and puts the
  condition grade in the last column.
- **Spelling drift** is unified: KARAGANDA→KARAGANDY, MALASZEWICAE→MALASZEWICZE,
  KHAVAROVSK/HABAROVSK→KHABAROVSK, RYBACHYE→BALYKCHY (same city, renamed in later sheets).

## Backups

`..\container-price-tool-backup-<date>\` holds a full copy of the tool as it stood before a
significant change. See the `RESTORE.txt` in each.

## Coverage

102 weekly quotations, 23 Sep 2024 – 28 Sep 2026, continuous. Four emails are
notice-only replies with no table (they appear as `NO TABLE` in the log — that is expected,
their wording is captured in `notices.csv`).

After a refresh, the only benign lines in `extract_log.txt` are those four `NO TABLE`
entries and two `SERIES OVERRIDE` lines for the 15 Jun 2026 Russia email. Anything else
usually means Jeff changed a table layout.
