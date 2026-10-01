# Weekly routine: Pasolite field visit dashboard

This is the instruction file for the scheduled Claude Code routine on this repository. It runs every Monday at 15:00 IST.

## Steps

1. **Find the new file.** Look in the Google Drive folder "Lystloc - Weekly Checkins Report" (folder id `1l1cGoCyNlKifpMrykwBlPuuSDR5yVK-H`) and find the newest .xlsx file. This should be the GWC104 Weekly Checkins Report for the week that just ended.
   - If there is no file for the week that just ended, stop.
   - Report: "No LystLoc file in Drive yet."
2. **Download and save it.** Download it with the Google Drive connector, decode the base64 and save it as `inbox/<file name>.xlsx`.
3. **Run the pipeline.**
   - `pip install -r requirements.txt` if openpyxl is missing.
   - `python3 run_week.py inbox/<file name>.xlsx`
   - If the output says STOPPED, do not commit anything. Report the reason in plain words.
4. **Check the SUMMARY.** Flag any of the following in the report:
   - forms outside 90 to 145
   - unparsed notes above 0
   - a gap_warning
   - any new_dropdown_values. These mean LystLoc has switched to the new drop-downs and the dashboard needs its parser update. Say so first.
5. **Commit and push.**
   - Run `git add index.html data/cleaned/`.
   - Commit with the message "Week <Mon> to <Sun> <year>: dashboard update", then push to main.
   - Never commit anything in `inbox/` or `raw/`.
6. **Confirm the push.** Run `git fetch` and check that origin/main holds the new commit.
7. **Report.** Write three lines:
   - the week covered
   - forms, new leads, existing leads
   - anything flagged

   Add one WhatsApp-ready line Harsh can forward to HR.

## Rules

- Do not change `build_dashboard.py` or the dashboard logic.
- Do not send messages to anyone.
- The data is customer data. Do not copy it anywhere other than this repository.
