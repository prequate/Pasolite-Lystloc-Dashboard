# Pasolite Field Visit & Lead Dashboard

A single-file dashboard built from Pasolite's weekly LystLoc check-in exports. HR and management use it to see what each sales executive is doing: new leads, repeat visits, whether promised follow-ups were kept, and what is overdue.

Built and maintained by Prequate Advisory for Pasolite Electricals Pvt. Ltd.

## Read this first: the data is private

`index.html` and every file in `data/cleaned/` contain customer names, site addresses and the reps' own visit notes.

- **Keep this repository private.**
- **Do not turn on GitHub Pages.** A Pages site is public on the internet even when the repository is private, unless the organisation is on GitHub Enterprise Cloud with private Pages.
- To share the dashboard with HR, send them `index.html` directly. It opens by double-clicking and works offline.

## What is in this repository

| Path | What it is |
|---|---|
| `index.html` | The dashboard. Open it in any browser. Everything it needs is inside the file. |
| `build_dashboard.py` | Rebuilds `index.html` from the weekly files. |
| `data/cleaned/` | One cleaned LystLoc export per week, named `Pasolite_LystLoc_Checkins_Cleaned_DD-Mon-YYYY.xlsx` (the Sunday the week ends). |
| `assets/pasolite_logo.png` | Logo embedded in the dashboard header. |
| `vendor/` | The PDF library behind the Overdue PDF download (jsPDF and jsPDF-AutoTable, MIT). Licences in `vendor/LICENSES.md`. |
| `requirements.txt` | Python package needed to rebuild. |

## Adding a new week

1. Download the week's check-in export from LystLoc.
2. Clean it into the standard format. This is done in the Claude project "[Pasolite]_Quarterback", following the LystLoc check-in cleaning instructions kept there. The output is one file named `Pasolite_LystLoc_Checkins_Cleaned_DD-Mon-YYYY.xlsx`.
3. Put that file in `data/cleaned/`.
4. Rebuild:

   ```
   pip install -r requirements.txt
   python build_dashboard.py
   ```

   The script picks up every file in `data/cleaned/` on its own, oldest week first, and rewrites `index.html`.
5. On GitHub, use **Add file > Upload files** to upload the new weekly file into `data/cleaned/` and the new `index.html` to the top level. Then open both on GitHub and check they show the new week. A web upload can fail without any error message, so always check.

## How the numbers are worked out

- **New vs existing lead:** a lead is existing if the same name, ignoring capitals and spaces, was logged in an earlier week. Tracking started on 17 Aug 2026, so "new" means new to this tracker, not new to Pasolite. A spelling change counts as a new lead.
- **Follow-up deadline:** read from the rep's Next Follow-up entry.
  - A date (10-09-2026, 25/08/2026, 3rd September) means that day.
  - A named day means the next time that day comes round.
  - "This week" means by Sunday. "Next week" means by the end of next week.
  - "Next month" means by the end of next month. "October first week" means by 7 October.
  - A spelled-out gap such as "in 2 days" is counted exactly.
  - No usable date: the deadline comes from the Requirement Timeline, at most 30 days after the visit. Immediate is 7 days, Within 15 days is 15 days, anything longer is 30 days. Marked "from timeline".
  - No date and no timeline: the end of the week after the visit, marked "assumed".
  - The LystLoc drop-down (live from 29 Sep 2026) is read with or without the number in front ("4 - 16 to 30 days"), counted from the visit date: 0 to 3 days (3 days), 4 to 7 days (7), 8 to 15 days (15), 16 to 30 days (30), More than 30 days (45), Client will call back (15). No follow-up required sets no deadline.
  - Older free-text options are still read: Tomorrow (1 day), Within 3 days (3 days), This week (Sunday), Next week (Sunday after), Within 15 days (15 days), This month (month end), Next month (end of next month).
- **Closed leads:** a lead whose latest visit says Order confirmed, Not interested, or Met, no live requirement with No follow-up required, is closed. It never shows as overdue.
- **Conflicting entry:** the form contradicts itself, for example Not interested with a follow-up date, or an open stage such as Samples or quote asked with No follow-up required. The lead stays on the overdue watch, tagged "conflicting entry", so HR can ask the rep which is right.
- **Followed Up On Time:** the next visit to the lead, by any rep, came on or before the deadline. The credit goes to the rep who made the promise. A second form for the same lead on the same day counts as kept.
- **Overdue:** the deadline on the lead's latest visit has passed and nobody has visited since. It is measured against the day the dashboard is opened, not the last date in the data.
- **Lead Stage:** replaced Client Satisfaction on the LystLoc form from 29 Sep 2026. Weeks before that keep their old label, shown as "Interested (old form)" and similar, so no old entry is re-scored.
- **Lead groups:** Type of Lead is grouped into Specifier (Architect, Interior designer, Lighting or MEP consultant), Project (Builder or developer, Procurement or purchase manager, Contractor, Business or institution, and the old Project and Builders or Procurement Manager), Residential (Homeowner) and Channel (Dealer or electrical shop, Distributor). Lead Mix shows each group with its types under it.
- **Weeks across two months:** a selected week always shows in full, even when it runs into the next month.
- **Joint visits:** a lead leaves a rep's overdue list once any colleague visits it later. It then shows under "Followed Up by a Colleague" in that rep's view. A lead two reps visited together, with no visit since, stays on both reps' lists.

## Layout

The dashboard opens on the most recent week in the data (and its month), on the **Scoreboard** tab.

- **Tabs.** Scoreboard, Section 1 Lead Mix, Section 2 Follow-up Watch, Section 3 New Leads, Section 4 Existing Leads and Section 5 Missed Days. One shows at a time, and each tab shows its count for the current filters. Adding `#stale` (or another section id) to the address opens that tab directly. The old Overview tab is gone; an old `#overview` link opens the Scoreboard.
- **Scoreboard.** One row per sales rep, A to Z: Forms, New Leads, Existing Leads, Followed Up On Time, Overdue and Missed Days, shaded so the larger numbers stand out. Picking a name in Sales Executive moves that rep to the top with a "Selected" tag. Clicking a column header still sorts by that column. Click any number to slide in the leads behind it. Click a rep's name to open their full view.
- **Grouped lists.** With All Executives selected, every list in Sections 2 to 5 is grouped by sales rep, collapsed, most first. With All Executives, the tables in Sections 2, 3 and 4 also show the Sales Executive right after the Lead.
- **"You are here" bar.** A thin bar under the tabs names the section, the table and the rep group in view.
- **Missed Days** leaves Sundays out.

## Lead Mix layout

Three columns: Project Requirement Timeline and Lead Stage on the left, Type of Lead in the middle, Estimated Project Value on the right.

- **Type of Lead** is shown under four captions: Specifier, Project, Residential and Channel. Each caption carries its total. Captions are not clickable; click a lead type to filter.
- **Lead Stage** lists the new-form stages first. Labels from the old form (before 29 Sep 2026) sit under their own "Old form, before 29 Sep" caption.

## Finding the reps behind a lead mix

In Section 1 (Lead Mix), select All Executives and pick two or more rows or slices, for example Architect and 1 to 2 months. A panel below shows one card per sales rep with the number of matching leads, most first. Click a card to switch the whole dashboard to that rep with the same filters still on. **Back to all sales executives** returns to the cards.

## On a phone

At 640px wide and below, the dashboard switches to a phone layout. The desktop view is unchanged.

- The frozen header is one slim bar. A **Filters** button opens the Month, Week and Sales Executive drop-downs, and a line under it shows the current selection.
- Every table shows as cards, one per row, with each field as a label and value.
- Long lists show 20 cards at a time, with a **Show 20 more** button. Search and filters still cover every row.
- Sorting by column header is desktop only.

## Known caveats

- **Data cut-off:** overdue counts run to the day the file is opened. A visit made after the last weekly upload will not show until the next upload.
- **Rep entries:** Est. Project Value, Requirement Timeline, Type of Lead, Lead Stage (Client Satisfaction before 29 Sep 2026) and Remarks are what the rep typed. Nobody has audited them.
- **Names:** leads are matched on the name as typed, so two different people with the same name can merge into one history.
- **PDF characters:** the overdue PDF prints "Rs" in place of the rupee sign and drops emoji, because standard PDF fonts cannot print them.
