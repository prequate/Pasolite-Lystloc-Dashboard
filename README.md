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
  - No usable date means the end of the week after the visit, marked "assumed".
- **Followed Up On Time:** the next visit to the lead, by any rep, came on or before the deadline. The credit goes to the rep who made the promise. A second form for the same lead on the same day counts as kept.
- **Overdue:** the deadline on the lead's latest visit has passed and nobody has visited since. It is measured against the day the dashboard is opened, not the last date in the data.
- **Joint visits:** a lead leaves a rep's overdue list once any colleague visits it later. It then shows under "Followed Up by a Colleague" in that rep's view. A lead two reps visited together, with no visit since, stays on both reps' lists.

## Known caveats

- **Data cut-off:** overdue counts run to the day the file is opened. A visit made after the last weekly upload will not show until the next upload.
- **Rep entries:** Est. Project Value, Requirement Timeline, Type of Lead, Client Satisfaction and Remarks are what the rep typed. Nobody has audited them.
- **Names:** leads are matched on the name as typed, so two different people with the same name can merge into one history.
- **PDF characters:** the overdue PDF prints "Rs" in place of the rupee sign and drops emoji, because standard PDF fonts cannot print them.
