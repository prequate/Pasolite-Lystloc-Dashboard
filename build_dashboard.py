#!/usr/bin/env python3
"""
Pasolite - Field Visit & Lead Dashboard (Task E, independent of Task D).

Reads the already-cleaned Task D output (one row per LystLoc check-in, with
Checkout Meeting Notes already split into columns) and builds a standalone
HTML dashboard: Month / Week / Sales Executive toggles, a plain (non-
interactive) facts strip (forms submitted, new leads, existing leads),
separate New Leads and Existing Leads tables for the selected period, a
Follow-up Watch section (Section 4) with a "Followed Up On Time" table
(promises the rep actually kept) followed by one flat "Overdue Follow-ups"
table sorted by days overdue - the earlier "Never Followed Up" / "Follow-up
Lapsed Again" split was merged and then had its Status distinction dropped
entirely per Harsh's 23-Sep-2026 requests, both tables also carry a Sales
Executive column and a Not Interested tag pulled from Client Satisfaction -
a clickable timeline/lead-type breakdown (click a value to see the leads
behind it), and a No Check-ins Logged section (with a Day Type column
flagging Sundays and national public holidays) for days with a LystLoc entry
but no Checkin Time logged at all. (Conversion by Executive, the Visit Map,
the Client Satisfaction breakdown panel, the "Leads Completed" section and
its browser-local Mark Done mechanism have all been removed per Harsh's
requests across several 22/23-Sep-2026 follow-ups; the single Leads Visited
table is now split into New Leads / Existing Leads; the click-to-expand
interaction originally built onto the KPI cards was moved onto the Timeline
& Lead Type breakdown panels instead, per Harsh's correction on 23-Sep-2026.
On 24-Sep-2026 both Section 4 tables gained a "N visits" button that opens
the lead's full visit history, all reps and weeks, with each promise's
deadline and outcome; same-day second entries count as kept, per Harsh.
Also 24-Sep-2026: a Download PDF button for a selected rep's overdue list
(jsPDF, embedded from vendor/), the joint-visit rule (a lead leaves a rep's
overdue list once any colleague visits later), a Followed Up by a Colleague
table in the rep view, and Section 1 headings renamed to the field names.)

This build (23-Sep-2026, third pass) processes all 5 weeks tracked so far
in one pass, in chronological order: 17-23 Aug, 24-30 Aug, 31 Aug-06 Sep,
07-13 Sep, 14-20 Sep 2026. Designed to keep accumulating: point
EXISTING_DASHBOARD_HTML at the previously-published dashboard and add only
the new week(s) to CLEANED_FILES for future runs, and it will merge rather
than start over (de-duplicated by weekStart).
"""
import json
import os
import re
from datetime import datetime, timedelta

import openpyxl

# ---- Repo layout (GitHub version) -------------------------------------------
# Every file in data/cleaned/ named Pasolite_LystLoc_Checkins_Cleaned_DD-Mon-YYYY.xlsx
# is picked up automatically, oldest week first. To add a week, drop the new
# cleaned file into data/cleaned/ and run this script; it rewrites index.html.
REPO_DIR = os.path.dirname(os.path.abspath(__file__))
CLEANED_DIR = os.path.join(REPO_DIR, "data", "cleaned")


def _discover_cleaned_files():
    found = []
    for fn in os.listdir(CLEANED_DIR):
        m = re.match(r"Pasolite_LystLoc_Checkins_Cleaned_(\d{2}-[A-Za-z]{3}-\d{4})\.xlsx$", fn)
        if m:
            found.append((os.path.join(CLEANED_DIR, fn), m.group(1)))
    if not found:
        raise SystemExit("No cleaned weekly files found in " + CLEANED_DIR)
    return sorted(found, key=lambda x: datetime.strptime(x[1], "%d-%b-%Y"))


CLEANED_FILES = _discover_cleaned_files()
OUT_HTML = os.path.join(REPO_DIR, "index.html")
VENDOR_DIR = os.path.join(REPO_DIR, "vendor")
LOGO_PATH = os.path.join(REPO_DIR, "assets", "pasolite_logo.png")

# NOTE (22-Sep-2026): the Visit Map section was removed at Harsh's request
# ("doesn't look good"), so the India state/national boundary JSON is no
# longer loaded or embedded here. Geocoding itself (GEOCODE, extract_city_
# state, geocode_location, city/state/lat/lon on each record) is left in
# place - low risk to keep, and it may be useful again if a map or
# territory view is asked for later.

# If re-running against an existing dashboard to append a new week, point
# this at the previously-published HTML (read via Projects.project_read and
# saved locally first) and list only the NEW week(s) in CLEANED_FILES above.
EXISTING_DASHBOARD_HTML = None

HONORIFIC_TOKENS = {
    "mr", "mrs", "ms", "er", "ar", "sir", "madam", "dr", "shri", "smt",
}

# See the long comment in build_records(): tested against real data and
# found to be a dead signal (essentially every lead row has a follow-up
# date filled in, new or old contact alike). Kept as a named flag rather
# than deleted so it's a one-line flip if a future week's data ever shows
# real variation on this field.
APPLY_FOLLOWUP_PROXY = False

# Found in the 20-Sep-2026 extract: several check-ins logged the lead name
# as literally "Na Na" (placeholder / not filled in). Naively repeat-
# matching would wrongly merge those into "the same lead visited N times".
# Any normalized name in this set is never matched against history - it
# always counts as New, and is flagged separately as a data-quality item.
PLACEHOLDER_NAMES = {"", "na na", "na", "n a", "none", "nil", "unknown", "xxx", "test"}

# ---------------------------------------------------------------------------
# City / state extraction from free-text Checkin Location strings, and a
# geocoding lookup for the map view. The extraction heuristic finds the
# LAST Indian state name mentioned in the address (real addresses put the
# state near the end, right before the pincode; taking the first match can
# misfire on a road literally named e.g. "Goa Road" earlier in the string)
# and takes the text segment immediately before it as the city/locality.
# This is a heuristic over free text typed by field reps, not a validated
# geocoder, and is disclosed as such in the dashboard footer.
STATE_LIST = [
    "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chhattisgarh", "Goa",
    "Gujarat", "Haryana", "Himachal Pradesh", "Jharkhand", "Karnataka", "Kerala",
    "Madhya Pradesh", "Maharashtra", "Manipur", "Meghalaya", "Mizoram", "Nagaland",
    "Odisha", "Punjab", "Rajasthan", "Sikkim", "Tamil Nadu", "Telangana", "Tripura",
    "Uttar Pradesh", "Uttarakhand", "West Bengal", "Delhi", "Puducherry",
    "Jammu and Kashmir", "Ladakh", "Chandigarh",
]
_STATE_SORTED = sorted(STATE_LIST, key=len, reverse=True)
STATE_PATTERN = re.compile(r"\b(" + "|".join(re.escape(s) for s in _STATE_SORTED) + r")\b", re.IGNORECASE)
STATE_CANON = {s.lower(): s for s in STATE_LIST}

LATLON_PATTERN = re.compile(
    r"^\s*(-?\d{1,3}\.\d+)\s*°?\s*[NS]?\s*,\s*(-?\d{1,3}\.\d+)\s*°?\s*[EW]?\s*$",
    re.IGNORECASE,
)


def extract_city_state(location):
    if not location:
        return None, None
    matches = list(STATE_PATTERN.finditer(location))
    if not matches:
        return None, None
    m = matches[-1]
    state_title = STATE_CANON[m.group(1).lower()]
    before = location[: m.start()]
    before = re.sub(r"[-,]\s*$", "", before).strip()
    parts = [p.strip() for p in before.split(",") if p.strip()]
    if not parts:
        return None, state_title
    city = parts[-1]
    city = re.sub(r"\s*-?\s*\d{6}$", "", city).strip()  # trailing pincode
    city = re.sub(r"^\d{6}\s*", "", city).strip()  # leading pincode
    return city or None, state_title


# Geocoded once (22-Sep-2026) via the Google Places API for every distinct
# city/locality bucket found across the 3 weeks tracked so far (77 buckets).
# Real, source-backed coordinates - not hand-estimated. Re-run the geocoding
# step for any NEW city bucket a future week introduces (this dict will not
# have it, and that location will be reported as "unmapped" rather than
# guessed).
GEOCODE = {
    "Belagavi": (15.860897399999997, 74.5129177),
    "Rangareddy": (17.1999602, 78.5505481),
    "Hubballi": (15.364708299999998, 75.1239547),
    "Pune": (18.5246091, 73.8786239),
    "Mysuru": (12.295810399999999, 76.6393805),
    "Vijayapura": (16.8302397, 75.71055799999999),
    "Hyderabad": (17.406498, 78.47724389999999),
    "Coimbatore": (10.9973691, 76.9588876),
    "Kalaburagi": (17.329731, 76.8342957),
    "Kolkata": (22.5743545, 88.3628734),
    "Salem": (11.6751854, 78.1222336),
    "Dharwad": (15.4589466, 75.00785859999999),
    "Bhubaneswar": (20.295984699999998, 85.8246101),
    "Mumbai": (18.9582347, 72.8319514),
    "Ernakulam": (9.9816358, 76.2998842),
    "Navi Mumbai": (19.0330488, 73.0296625),
    "Bidar": (17.9103939, 77.51990789999999),
    "Sirsi": (14.6155109, 74.8347096),
    "Thane": (19.2122949, 72.97716609999999),
    "Jalgaon": (21.0076578, 75.5626039),
    "Jamkhandi": (16.5043199, 75.2917512),
    "Janthagalli": (12.240584199999999, 76.7293347),
    "Hutagalli": (12.3490848, 76.5777203),
    "Rajamahendravaram": (17.0095854, 81.7797234),
    "Sriramapura": (12.994843099999999, 77.5664029),
    "Pimpri-Chinchwad": (18.6292517, 73.8114318),
    "Kumbalgodu": (12.879400799999999, 77.44513169999999),
    "Rasipuram": (11.4614624, 78.1854772),
    "Nanded": (19.1485937, 77.31912),
    "Dhenkanal": (20.658446599999998, 85.5964935),
    "Namakkal": (11.219384800000002, 78.16784179999999),
    "Mudhol": (16.3333105, 75.2858208),
    "Salem (M.Corp.)": (11.6751854, 78.1222336),
    "Muthukalipatti": (11.4715362, 78.1689808),
    "Rammanahalli": (12.3465332, 76.70984159999999),
    "Thiruvananthapuram": (8.5241391, 76.9366376),
    "Kottayam": (9.591566799999999, 76.5221531),
    "Sangareddy": (17.6075232, 78.0798191),
    "Bhogadi": (15.860897399999997, 74.5129177),
    "Tiruchirappalli": (10.7904833, 78.7046725),
    "Bamphakuda": (20.3707684, 85.8985046),
    "Pahala": (20.347441, 85.8863585),
    "Phulnakhara": (20.364642099999998, 85.8908337),
    "Kesora": (20.273752299999998, 85.8818829),
    "Badaraghunathpur": (20.2352151, 85.7283935),
    "Padanpur": (20.1753061, 85.6807271),
    "Somayampalayam": (11.046903799999999, 76.90187759999999),
    "Erode": (11.3410364, 77.7171642),
    "Palakkad": (10.7733121, 76.6580401),
    "Puri": (19.8134554, 85.8312359),
    "Maltipatpur": (19.866902, 85.831986),
    "Chaudabatia": (19.970658699999998, 85.82174479999999),
    "Mallamooppampatti": (11.6825035, 78.0965691),
    "Olakkachinnanur": (11.5295892, 77.9138889),
    "Dowlaiswaram": (16.955761, 81.7927436),
    "Medchal Malkajgiri": (17.569841399999998, 78.52065),
    "Reddipatti": (11.209719999999999, 78.2044997),
    "Kollam": (8.8932118, 76.6141396),
    "Anagahalli": (12.3713185, 76.6022778),
    "Chennammana Kitturu": (15.5944573, 74.78929509999999),
    "Marandahalli": (12.389069399999999, 78.00327990000001),
    "Erranahalli": (12.2838622, 78.0594874),
    "Malgathi": (17.114853699999998, 76.9751074),
    "Hagarga": (17.348433699999998, 76.91460769999999),
    "Kataka": (20.462521, 85.8829895),
    "Thrissur": (10.524117600000002, 76.2120649),
    "Andagalore": (11.4467748, 78.1569468),
    "Tumbipadi": (11.8062384, 78.07802989999999),
    "Ekachalia": (20.1747072, 85.8484769),
    "Madurai": (9.9252007, 78.1197754),
    "M.Kalipatti": (11.8259639, 77.9331706),
    "Nashik": (19.9993217, 73.79001880000001),
    "Alappuzha": (9.498066699999999, 76.3388484),
    "Vijayawada": (16.506174299999998, 80.6480153),
    "Belagavi Cantonment": (15.859205899999997, 74.5023819),
    "Sittaneri": (11.5886882, 78.0922911),
    "North Dumdum": (22.662608499999997, 88.40900429999999),
    # Added 22-Sep-2026 for the 17-23 Aug and 24-30 Aug weeks, geocoded the
    # same way (Google Places API, via the places_search tool) as the
    # original 77 buckets above.
    "Kochi": (9.9312328, 76.26730409999999),
    "Kakkanad": (10.0158605, 76.3418666),
    "Bengeri": (15.370001199999997, 75.1473127),
    "Dindigul": (10.361965, 77.9735844),
    "Nagalur": (11.838424999999999, 78.21058570000001),
    "Tiruchengode": (11.379023499999999, 77.8949435),
    "Kendrapara": (20.5035436, 86.4199321),
    "Dongaon": (20.1818568, 76.7206652),
    "Kaikhali": (22.6329544, 88.4348617),
    "Calangute": (15.5456923, 73.7607543),
    "Koneripatti": (11.578991499999999, 77.7459234),
    "Sendamangalam": (11.2813371, 78.2357948),
}


def geocode_location(location):
    """Return (city, state, lat, lon) for a Checkin Location value, or
    ("", "", None, None) if it can't be placed. Falls back to parsing a raw
    GPS pin string when no address/state text is present."""
    city, state = extract_city_state(location)
    if city and city in GEOCODE:
        lat, lon = GEOCODE[city]
        return city, state or "", lat, lon
    m = LATLON_PATTERN.match(location or "")
    if m:
        try:
            return "", "", float(m.group(1)), float(m.group(2))
        except ValueError:
            pass
    return "", "", None, None


def normalize_lead_name(name):
    """Loose key for repeat-lead detection: lowercase, strip honorifics/
    punctuation, collapse whitespace. Imperfect on purpose - documented as
    a heuristic, never used to silently merge records, only to flag."""
    if not name:
        return ""
    s = re.sub(r"[.,]", " ", name.lower())
    tokens = [t for t in s.split() if t not in HONORIFIC_TOKENS]
    return " ".join(tokens).strip()


def parse_date_ddmmyyyy(s):
    return datetime.strptime(s.strip(), "%d-%m-%Y")


def week_start(dt):
    """Monday of dt's week."""
    return dt - timedelta(days=dt.weekday())


def month_label(dt):
    return dt.strftime("%B %Y")


def week_label(monday):
    sunday = monday + timedelta(days=6)
    if monday.month == sunday.month:
        return f"{monday.day} - {sunday.day} {monday.strftime('%b %Y')}"
    return f"{monday.strftime('%d %b')} - {sunday.strftime('%d %b %Y')}"


def fmt_time_cell(date_str, time_str):
    """'15 Sep 2026 06:22 PM' -> '15 Sep, 06:22 PM'. Falls back to raw."""
    if not time_str or time_str == "-":
        return ""
    try:
        dt = datetime.strptime(time_str.strip(), "%d %b %Y %I:%M %p")
        return dt.strftime("%d %b, %I:%M %p")
    except ValueError:
        return time_str


# India national public holidays for 2026 [CATEGORY RESEARCH, NOT CLIENT-
# CONFIRMED] - replaced 23-Sep-2026 (was previously a Karnataka-specific
# list; Harsh flagged that as the wrong scope since Pasolite's sales
# executives are spread across many states, not concentrated in Karnataka).
# This is the Central Government of India's gazetted holiday list for 2026
# (DoPT OM No. F.No.12/2/2023-JCA, notified for Central Govt establishments
# nationwide), cross-checked against two independent published sources
# (govtcalendar.org's DoPT-sourced calendar and bankbazaar.com's 2026 Indian
# holiday calendar, both retrieved 23-Sep-2026); both agree on every date
# below. Of these, only 3 (Republic Day, Independence Day, Gandhi Jayanthi)
# are compulsory national holidays observed by every establishment,
# government and private, everywhere in the country; the rest are the wider
# Central Government gazetted list and may not be observed as a holiday by
# every Pasolite office or by every state government - this is the best
# available NATIONAL baseline (not tied to one state), not a confirmed
# Pasolite company holiday calendar, which has not been confirmed - flagged
# in the Human Validation checklist. The two Islamic dates (Id-ul-Fitr,
# Id-ul-Zuha/Bakrid) and Muharram are fixed by moon sighting and were
# tentative by up to a day at both source sites at the time of retrieval.
# Corrected 23-Sep-2026 during an independent data-sanity pass: Id-ul-Zuha
# (Bakrid) is 27-May-2026, not 28-May as first entered - caught by cross-
# checking this whole dict against Python's own `holidays` library (a third,
# independent source) rather than only the two sites above. This date falls
# outside every week currently tracked, so it changed nothing already shown
# to Harsh - flagged here so the correction itself is on record.
# Sundays are computed directly from the date, not from this dict.
INDIA_NATIONAL_HOLIDAYS_2026 = {
    "2026-01-26": "Republic Day",
    "2026-03-04": "Holi",
    "2026-03-21": "Id-ul-Fitr",
    "2026-03-26": "Ram Navami",
    "2026-03-31": "Mahavir Jayanti",
    "2026-04-03": "Good Friday",
    "2026-04-14": "Dr. B.R. Ambedkar Jayanti",
    "2026-05-01": "Buddha Purnima",
    "2026-05-27": "Id-ul-Zuha (Bakrid)",
    "2026-06-26": "Muharram",
    "2026-08-15": "Independence Day",
    "2026-08-26": "Milad-un-Nabi (Prophet's Birthday)",
    "2026-09-04": "Janmashtami",
    "2026-10-02": "Gandhi Jayanti",
    "2026-10-20": "Dussehra (Vijaya Dashami)",
    "2026-11-08": "Diwali",
    "2026-11-24": "Guru Nanak Jayanti",
    "2026-12-25": "Christmas",
}


def day_type(dt):
    """'Sunday', a holiday name, both combined, or '' for an ordinary day."""
    parts = []
    if dt.weekday() == 6:
        parts.append("Sunday")
    holiday = INDIA_NATIONAL_HOLIDAYS_2026.get(dt.strftime("%Y-%m-%d"))
    if holiday:
        parts.append(holiday)
    return ", ".join(parts)


def load_cleaned_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["LystLoc Checkins - Clean"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        rows.append({h: r[i] for h, i in idx.items()})
    return rows


def build_records(cleaned_rows, seen_normalized_names):
    """Turn cleaned Task D rows into dashboard row records. seen_normalized_names
    is the set of normalized lead names already known from *earlier* weeks
    (empty on week 1) - mutated in place as this week's leads are processed,
    in date order, so a lead repeating within the same week is only "New"
    once."""
    records = []
    unnamed_count = 0
    unmapped_locations = 0
    for row in cleaned_rows:
        date_str = row.get("Date")
        if not date_str:
            continue
        dt = parse_date_ddmmyyyy(str(date_str))
        wk = week_start(dt)

        lead = (row.get("Lead") or "").strip()
        has_lead = bool(lead)

        is_new = None
        is_existing = None
        norm = normalize_lead_name(lead)
        if has_lead:
            # Combined 1+2 rule as specified, with one finding from testing
            # it against real data: rule 2 (follow-up date as a proxy for
            # "already known") never actually fires, because essentially
            # every lead row across all 3 tracked weeks has a Next Follow-up
            # Date filled in - reps log a next step after every visit
            # regardless of whether the contact is brand new or long-
            # standing. Rule 1 (repeat name across weeks, in chronological
            # order) is what actually does the work. Kept as a named flag
            # (APPLY_FOLLOWUP_PROXY) rather than deleted, in case a future
            # week's data genuinely varies on this field.
            has_followup = bool((row.get("Next Follow-up Date") or "").strip())
            is_placeholder = norm in PLACEHOLDER_NAMES
            seen_before = (not is_placeholder) and norm in seen_normalized_names
            if is_placeholder:
                unnamed_count += 1
                is_new = True
                is_existing = False
            elif seen_before:
                is_existing = True
                is_new = False
            elif APPLY_FOLLOWUP_PROXY and has_followup:
                is_existing = True
                is_new = False
            else:
                is_new = True
                is_existing = False
            if not is_placeholder:
                seen_normalized_names.add(norm)

        location = (row.get("Checkin Location") or "").strip()
        city, state, lat, lon = geocode_location(location)
        has_location = bool(location and location != "-")
        if has_location and lat is None:
            unmapped_locations += 1

        # Harsh (22-Sep-2026): a rep can have a placeholder LystLoc row for a
        # day with literally nothing entered (Checkin Time, Location, Notes
        # all "-") - that isn't a submitted form, it's the absence of one.
        # These should not inflate Forms Submitted; they're surfaced instead
        # in their own "No Check-ins Logged" section (see hasCheckinTime).
        checkin_time_raw = (row.get("Checkin Time") or "").strip()
        has_checkin_time = bool(checkin_time_raw) and checkin_time_raw != "-"

        records.append({
            "date": dt.strftime("%Y-%m-%d"),
            "dateDisplay": dt.strftime("%d %b %Y"),
            "weekStart": wk.strftime("%Y-%m-%d"),
            "weekLabel": week_label(wk),
            "month": month_label(dt),
            "exec": (row.get("Name") or "").strip(),
            "hasCheckinTime": has_checkin_time,
            "dayType": day_type(dt),
            "hasLead": has_lead,
            "isNew": is_new,
            "isExisting": is_existing,
            "lead": lead,
            "location": location,
            "city": city,
            "state": state,
            "lat": lat,
            "lon": lon,
            "hasLocation": has_location,
            "time": fmt_time_cell(str(date_str), str(row.get("Checkin Time") or "")),
            "value": row.get("Estimated Project Value") or "",
            "timeline": row.get("Project Requirement Timeline") or "",
            "satisfaction": row.get("Client Satisfaction (1-3)") or "",
            "followUp": row.get("Next Follow-up Date") or "",
            "leadType": row.get("Type of Lead") or "",
            "remarks": row.get("Remarks") or "",
            "unnamedLead": has_lead and normalize_lead_name(lead) in PLACEHOLDER_NAMES,
        })
    return records, unnamed_count, unmapped_locations


def extract_existing_data_and_seen(html_path):
    """Pull DASHBOARD_DATA back out of a previously-published dashboard file,
    plus rebuild the seen-name set from it (so week N+1 knows what week 1..N
    already saw)."""
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()
    m = re.search(r"const DASHBOARD_DATA = (\{.*?\});\s*\n</script>", html, re.DOTALL)
    if not m:
        raise RuntimeError("Could not find DASHBOARD_DATA in existing dashboard HTML")
    data = json.loads(m.group(1))
    seen = set()
    for rec in data["records"]:
        if rec["hasLead"]:
            seen.add(normalize_lead_name(rec["lead"]))
    return data, seen


def main():
    if EXISTING_DASHBOARD_HTML:
        existing_data, seen = extract_existing_data_and_seen(EXISTING_DASHBOARD_HTML)
        existing_records = existing_data["records"]
        existing_weeks = {r["weekStart"] for r in existing_records}
    else:
        existing_records = []
        seen = set()
        existing_weeks = set()

    all_new_records = []
    total_unnamed = 0
    total_unmapped = 0
    sources = []
    for path, label in CLEANED_FILES:
        cleaned_rows = load_cleaned_rows(path)
        recs, unnamed, unmapped = build_records(cleaned_rows, seen)
        all_new_records.extend(recs)
        total_unnamed += unnamed
        total_unmapped += unmapped
        sources.append(label)
        print(f"{label}: {len(recs)} rows, {sum(1 for r in recs if r['hasLead'])} leads, "
              f"{sum(1 for r in recs if r['isNew'])} new, {sum(1 for r in recs if r['isExisting'])} existing, "
              f"{unnamed} unnamed, {unmapped} unmapped locations")

    new_weeks = {r["weekStart"] for r in all_new_records}
    overlap = new_weeks & existing_weeks
    if overlap:
        # Don't silently double-count a week that's already in the dashboard.
        existing_records = [r for r in existing_records if r["weekStart"] not in overlap]

    all_records = existing_records + all_new_records
    all_records.sort(key=lambda r: (r["date"], r["exec"]))

    months = sorted({r["month"] for r in all_records},
                     key=lambda m: datetime.strptime(m, "%B %Y"))
    # A week can span two calendar months (e.g. 31 Aug - 06 Sep 2026). For the
    # Week dropdown, each week is filed under the month of its own Monday
    # (weekStart) - a single, unambiguous bucket - even though the KPI
    # figures themselves still group individual check-ins by their own
    # actual visit date via r["month"], so nothing about the underlying
    # counts is affected by this choice.
    weeks_by_start = {}
    for r in all_records:
        ws = r["weekStart"]
        if ws not in weeks_by_start:
            weeks_by_start[ws] = (ws, r["weekLabel"], month_label(datetime.strptime(ws, "%Y-%m-%d")))
    weeks = sorted(weeks_by_start.values(), key=lambda w: w[0])
    execs = sorted({r["exec"] for r in all_records if r["exec"]})

    dashboard_data = {
        "months": months,
        "weeks": [{"start": w[0], "label": w[1], "month": w[2]} for w in weeks],
        "execs": execs,
        "records": all_records,
        "meta": {
            "generatedFrom": ", ".join(f"Pasolite_LystLoc_Checkins_Cleaned_{s}.xlsx" for s in sources),
            "totalRecords": len(all_records),
        },
    }

    import base64
    with open(LOGO_PATH, "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode("ascii")
    logo_data_uri = f"data:image/png;base64,{logo_b64}"

    # jsPDF 4.2.1 + jsPDF-AutoTable 5.0.8 (both MIT) power the Overdue PDF
    # download. Embedded from vendor/ so the dashboard works offline.
    pdf_libs = "".join(
        f"<script>/* {name} (MIT licence) */\n" + open(os.path.join(VENDOR_DIR, fn), encoding="utf-8").read() + "\n</script>\n"
        for name, fn in [("jsPDF 4.2.1", "jspdf.umd.min.js"), ("jsPDF-AutoTable 5.0.8", "jspdf.plugin.autotable.min.js")]
    )
    html = HTML_TEMPLATE.replace(
        "__DASHBOARD_DATA__", json.dumps(dashboard_data, ensure_ascii=False)
    ).replace("__LOGO_DATA_URI__", logo_data_uri).replace("__PDF_LIBS__", pdf_libs)
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print("\nSaved:", OUT_HTML)
    print("Total records (all weeks):", len(all_records))
    print("Weeks:", [w[1] for w in weeks])
    print("Execs:", execs)
    print("Total unnamed/placeholder leads:", total_unnamed)
    print("Total unmapped locations (address present, not geocoded):", total_unmapped)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pasolite: Field Visit &amp; Lead Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,300;0,400;0,600;0,700;1,400&display=swap" rel="stylesheet">
<style>
  /* Colour tokens and component patterns ported directly from the current
     Task C build (Pasolite: Data Quality Dashboard) at Harsh's request, so
     the two dashboards read as one system rather than two different tools. */
  :root{
    color-scheme: light;
    --bg:#FFFFFF;
    --surface:#FFFFFF;
    --surface-2:#FAF8F6;
    --surface-3:#F3F1EE;
    --ink:#161616;
    --ink-secondary:#5B5B5B;
    --ink-muted:#8C8C8C;
    --border:rgba(22,22,22,0.11);
    --border-soft:rgba(22,22,22,0.07);
    --accent:#E1372B;
    --accent-deep:#B71C1C;
    --good:#2F8F5B;
    --good-bg:rgba(47,143,91,0.12);
    --watch:#8A8A8A;
    --watch-bg:rgba(112,112,112,0.13);
    --critical:#B23A2E;
    --critical-bg:rgba(178,58,46,0.13);
    --amber:#B4791E;
    --amber-bg:rgba(180,121,30,0.12);
    --teal:#2E7D8C;
    --teal-bg:rgba(46,125,140,0.12);
    --violet:#6B4C8A;
    --violet-bg:rgba(107,76,138,0.12);
    --good-deep:#1B7A46;
    --card-shadow: 0 1px 2px rgba(22,22,22,0.03), 0 10px 28px rgba(22,22,22,0.06);
  }
  *{box-sizing:border-box;}
  html,body{margin:0;padding:0;}
  body{
    background:var(--bg); color:var(--ink);
    font-family:'Product Sans','Open Sans',-apple-system,BlinkMacSystemFont,Helvetica,Arial,sans-serif;
    font-weight:400; line-height:1.45; -webkit-font-smoothing:antialiased;
    -webkit-text-size-adjust:100%; text-size-adjust:100%;
  }
  h1,h2,h3{font-family:'Product Sans','Open Sans',sans-serif; font-weight:700; margin:0;}
  .num{font-variant-numeric:tabular-nums;}

  .icon{
    width:1em; height:1em; display:inline-block; vertical-align:middle;
    /* 23-Sep-2026 fix: every icon on this dashboard was rendering as a solid
       filled blob rather than the intended stroke-line icon, because fill/
       stroke set on the <g> wrapping the <symbol> defs below never actually
       inherits into a <use> instance (a <symbol>'s effective parent for CSS
       inheritance is the referencing <use> element, not its literal parent
       in <defs> - a well-known SVG quirk). Setting it here, on the <svg
       class="icon"> itself, is what actually reaches the icon. Caught during
       the redesign mock's own browser-rendered verification, not just a
       code read - this was live and unnoticed before today. */
    fill:none; stroke:currentColor; stroke-width:1.8; stroke-linecap:round; stroke-linejoin:round;
  }
  .icon-badge{display:inline-flex; align-items:center; justify-content:center; flex:none; width:38px; height:38px; border-radius:10px;}
  .icon-badge .icon{width:19px; height:19px;}
  .icon-badge.tone-accent{background:var(--critical-bg); color:var(--accent-deep);}
  .icon-badge.tone-amber{background:var(--amber-bg); color:var(--amber);}
  .icon-badge.tone-teal{background:var(--teal-bg); color:var(--teal);}
  .icon-badge.tone-violet{background:var(--violet-bg); color:var(--violet);}
  .icon-badge.tone-good{background:var(--good-bg); color:var(--good);}
  .icon-badge.tone-watch{background:var(--watch-bg); color:var(--ink-secondary);}
  .icon-badge.tone-neutral{background:var(--surface-3); color:var(--ink-muted);}
  .icon-badge.sm{width:30px; height:30px; border-radius:8px;}
  .icon-badge.sm .icon{width:15px; height:15px;}

  /* ---------- Masthead + quicknav, frozen together while scrolling ---------- */
  .sticky-head{position:sticky; top:0; z-index:50;}
  .masthead{
    background:var(--ink); color:#fff; padding:14px clamp(20px,4vw,48px);
    box-shadow:0 8px 20px rgba(22,22,22,0.18);
  }
  .masthead-inner{display:flex; align-items:center; justify-content:space-between; gap:20px; flex-wrap:wrap;}
  .masthead-left{display:flex; align-items:center; gap:16px;}
  .wordmark-logo{height:26px; width:auto; display:block; flex:none;}
  .masthead-divider{width:1px; height:32px; background:rgba(255,255,255,0.2); flex:none;}
  .masthead-title h1{font-size:clamp(17px,2.1vw,21px); color:#fff;}
  .masthead-title p{margin:3px 0 0; font-size:12px; color:rgba(255,255,255,0.58); max-width:56ch;}
  .masthead-right{display:flex; align-items:center; gap:16px; flex-wrap:wrap;}
  .masthead-selector{display:flex; align-items:center; gap:12px; flex-wrap:wrap;}
  .control-group{display:flex; flex-direction:column; gap:4px;}
  .control-label{font-size:11px; text-transform:uppercase; letter-spacing:0.08em; font-weight:600; color:rgba(255,255,255,0.55); white-space:nowrap;}
  select#monthSelect, select#weekSelect, select#execSelect{
    appearance:none; -webkit-appearance:none;
    font-family:inherit; font-size:13.5px; font-weight:600; color:#fff;
    background:rgba(255,255,255,0.08) url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="12" height="8"><path d="M1 1l5 5 5-5" stroke="%23FFFFFF" stroke-width="1.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>') no-repeat right 14px center;
    border:1px solid rgba(255,255,255,0.28); border-radius:8px;
    padding:9px 34px 9px 14px; min-width:190px; cursor:pointer;
  }
  select#monthSelect{min-width:150px;}
  select#execSelect{min-width:180px;}
  select#monthSelect:focus-visible, select#weekSelect:focus-visible, select#execSelect:focus-visible{outline:2px solid var(--accent); outline-offset:1px;}
  select#monthSelect option, select#weekSelect option, select#execSelect option{color:var(--ink); background:#fff;}

  .quicknav{background:var(--ink); border-top:1px solid rgba(255,255,255,0.12); padding:0 clamp(20px,4vw,48px);}
  .quicknav-inner{max-width:1180px; margin:0 auto; display:flex; gap:0; flex-wrap:wrap;}
  .quicknav a{
    display:inline-block; padding:10px 16px; font-size:11.5px; font-weight:600; text-transform:uppercase;
    letter-spacing:0.05em; color:rgba(255,255,255,0.55); text-decoration:none; border-bottom:2px solid transparent;
  }
  .quicknav a:hover{color:#fff; border-bottom-color:var(--accent);}

  .wrap{max-width:1180px; margin:0 auto; padding:18px clamp(20px,4vw,48px) 32px;}

  section.block{
    margin-top:16px; background:var(--surface); border:1px solid var(--border);
    border-radius:16px; padding:18px 22px 20px; box-shadow:var(--card-shadow);
  }
  section.block:first-child{margin-top:0;}
  .block-head{display:flex; justify-content:space-between; align-items:flex-start; gap:16px; flex-wrap:wrap; margin-bottom:2px;}
  .block-head-left{display:flex; align-items:center; gap:12px;}
  .block-eyebrow{font-size:11px; text-transform:uppercase; letter-spacing:0.09em; color:var(--accent-deep); font-weight:700; margin-bottom:2px;}
  .block-head h2{font-size:16.5px;}
  .block-desc{font-size:12.5px; color:var(--ink-secondary); max-width:88ch; margin:6px 0 12px;}
  .block-hint{font-size:11.5px; color:var(--ink-muted); align-self:center; white-space:nowrap;}

  /* ---------- Overview: compact KPI tiles + a small New/Existing ring
     chart, replacing the old 3-tile-plus-long-sentence layout. The full
     reconciliation arithmetic is not gone, just moved behind "How this
     reconciles" (see .info-btn/.info-panel) so it doesn't sit in view by
     default - simplified 23-Sep-2026 per Harsh's request to cut wordiness
     and lean more graphical. ---------- */
  .overview-row{display:flex; align-items:stretch; gap:14px; flex-wrap:wrap;}
  .kpi-strip{display:flex; gap:12px; flex:1; min-width:280px;}
  .kpi-mini{
    flex:1; min-width:130px; background:var(--surface); border:1px solid var(--border); border-left:3px solid var(--border);
    border-radius:12px; padding:13px 16px; box-shadow:var(--card-shadow); display:flex; flex-direction:column; gap:2px;
  }
  .kpi-mini[data-kpi="forms"]{border-left-color:var(--teal);}
  .kpi-mini[data-kpi="new"]{border-left-color:var(--good);}
  .kpi-mini[data-kpi="existing"]{border-left-color:var(--accent-deep);}
  .kpi-mini .kv{font-size:25px; font-weight:700; color:var(--ink); letter-spacing:-0.01em; line-height:1;}
  .kpi-mini .kv.accent{color:var(--good);}
  .kpi-mini .kl{font-size:10.5px; color:var(--ink-muted); text-transform:uppercase; letter-spacing:0.06em; font-weight:700; margin-top:3px;}
  .donut-wrap{display:flex; align-items:center; gap:14px; border:1px solid var(--border); border-radius:12px; padding:12px 18px; box-shadow:var(--card-shadow);}
  .donut-legend{font-size:12px; color:var(--ink-secondary); display:flex; flex-direction:column; gap:5px;}
  .dot{width:9px; height:9px; border-radius:50%; display:inline-block; margin-right:7px;}

  /* ---------- Small "ⓘ" info toggle - used wherever a longer explanation
     used to sit inline by default (KPI reconciliation, Follow-up Watch and
     Missed Days methodology, footer disclosures). Nothing is deleted, it is
     one click away instead of always-on paragraph text. ---------- */
  .info-btn{
    display:inline-flex; align-items:center; gap:5px; font-size:11px; font-weight:600; color:var(--ink-secondary);
    border:1px solid var(--border); border-radius:999px; padding:5px 12px; cursor:pointer; background:var(--surface-2);
    white-space:nowrap; user-select:none;
  }
  .info-btn:hover{background:var(--surface-3); color:var(--ink);}
  .info-btn .icon{width:11px; height:11px;}
  .info-panel{display:none; margin-top:10px; padding:12px 16px; background:var(--surface-2); border:1px solid var(--border-soft); border-radius:10px; font-size:12px; color:var(--ink-secondary); line-height:1.6;}
  .info-panel.open{display:block;}
  .info-panel strong{color:var(--ink-secondary);}

  /* ---------- Stat pills - used for Follow-up Watch and Missed Days
     instead of a long paragraph caption. ---------- */
  .pill-row{display:flex; gap:14px; flex-wrap:wrap; margin:14px 0 4px;}
  .pill{flex:1; min-width:190px; display:flex; align-items:center; gap:12px; border:1px solid var(--border); border-radius:12px; padding:14px 16px; box-shadow:var(--card-shadow);}
  .pill .pv{font-size:23px; font-weight:700; line-height:1; color:var(--ink);}
  .pill .pl{font-size:11.5px; color:var(--ink-secondary); line-height:1.3; margin-top:3px;}
  .daytype-icon{display:inline-flex; align-items:center; justify-content:center; width:24px; height:24px; border-radius:6px; flex:none;}
  .daytype-icon .icon{width:13px; height:13px;}
  .daytype-icon.sun{background:var(--teal-bg); color:var(--teal);}
  .daytype-icon.holiday{background:var(--amber-bg); color:var(--amber);}
  .daytype-icon.gap{background:var(--critical-bg); color:var(--critical);}

  /* Shared expand-below-on-click detail panel - used by the Timeline &
     Lead Type breakdown panels (Section 4). Originally built for the KPI
     cards, moved here 23-Sep-2026 per Harsh's correction: the KPI cards are
     plain, non-interactive tiles again (see .kpi above). */
  .detail-panel{margin-top:16px; border:1px solid var(--border); border-radius:12px; background:var(--surface-2); overflow:hidden;}
  .detail-panel-head{display:flex; align-items:center; justify-content:space-between; gap:12px; padding:11px 16px; font-size:12px; font-weight:700; color:var(--ink); border-bottom:1px solid var(--border-soft); background:var(--surface);}
  .detail-panel-close{background:none; border:none; font-size:18px; line-height:1; cursor:pointer; color:var(--ink-muted); padding:2px 6px;}
  .detail-panel-close:hover{color:var(--ink);}
  .detail-panel-scroll{max-height:320px; overflow-y:auto;}
  .detail-panel-table{width:100%; min-width:0; font-size:12px;}
  .detail-panel-table th, .detail-panel-table td{padding:8px 14px;}
  /* Combined-filter chips (added 23-Sep-2026) - one chip per active value
     across Timeline / Lead Type / Value, each individually removable. See
     activeFilters / leadMatchesAll / leadMatchesOthers in the script. */
  .filter-chip-row{display:flex; flex-wrap:wrap; gap:6px; padding:10px 16px 2px 16px;}
  .filter-chip{display:inline-flex; align-items:center; gap:4px; background:var(--surface); border:1px solid var(--accent-deep); color:var(--ink); font-size:11.5px; font-weight:700; padding:4px 4px 4px 10px; border-radius:999px;}
  .filter-chip-x{background:none; border:none; cursor:pointer; color:var(--ink-muted); font-size:14px; line-height:1; padding:0 5px;}
  .filter-chip-x:hover{color:var(--ink);}

  /* Section 1 rep chips (24-Sep-2026): with All Executives selected and
     two or more Lead Mix filters on, the drill-down lists the sales reps
     behind the matching leads. Clicking one switches the whole dashboard to
     that rep and keeps the filters. */
  .rep-row{display:flex; flex-wrap:wrap; align-items:center; gap:6px; padding:10px 16px 4px 16px;}
  .rep-row-label{font-size:10.5px; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; color:var(--ink-muted); margin-right:4px;}
  .rep-chip{display:inline-flex; align-items:center; gap:7px; font-family:inherit; font-size:12px; font-weight:700; color:var(--ink);
    background:var(--surface); border:1px solid var(--border); border-radius:999px; padding:4px 5px 4px 11px; cursor:pointer;}
  .rep-chip:hover{border-color:var(--accent-deep); background:var(--critical-bg);}
  .rep-chip:focus-visible{outline:2px solid var(--accent); outline-offset:1px;}
  .rep-chip .rep-n{min-width:20px; padding:1px 7px; border-radius:999px; background:var(--ink); color:#fff; font-size:10.5px; text-align:center;}
  .rep-hint{padding:8px 16px 2px 16px; font-size:11.5px; color:var(--ink-muted);}
  .rep-back-row{display:flex; flex-wrap:wrap; align-items:center; gap:8px; padding:10px 16px 4px 16px; font-size:12px; color:var(--ink-secondary);}
  .rep-back-row strong{color:var(--ink);}
  .rep-back{display:inline-flex; align-items:center; gap:6px; font-family:inherit; font-size:11.5px; font-weight:700; color:#fff;
    background:var(--ink); border:1px solid var(--ink); border-radius:999px; padding:4px 12px; cursor:pointer;}
  .rep-back:hover{background:#2a2a2a;}
  .reconcile{font-size:12px; color:var(--ink-secondary); margin:14px 0 0; padding-top:2px;}
  .reconcile strong{color:var(--ink); font-weight:700;}

  .table-toolbar{display:flex; align-items:center; justify-content:space-between; gap:16px; margin:18px 0 12px; flex-wrap:wrap;}
  .table-toolbar input[type=search]{
    border:1px solid var(--border); border-radius:8px; padding:8px 12px; font-size:13px; font-family:inherit;
    min-width:240px; color:var(--ink); background:var(--surface-2);
  }
  .table-toolbar input[type=search]:focus{outline:none; border-color:var(--accent);}
  .row-count{font-size:12px; color:var(--ink-muted);}

  .table-wrap{overflow-x:auto; border-radius:10px; border:1px solid var(--border); margin-bottom:4px;}
  table.leads{border-collapse:collapse; width:100%; min-width:1100px; font-size:12.5px;}
  table.leads th, table.leads td{padding:10px 14px; text-align:left; border-bottom:1px solid var(--border-soft); vertical-align:top;}
  table.leads thead th{
    background:var(--ink); color:#fff; font-weight:700; font-size:11px;
    text-transform:uppercase; letter-spacing:0.05em; white-space:nowrap; cursor:pointer; user-select:none;
  }
  table.leads thead th:hover{background:#2a2a2a;}
  table.leads thead th .arrow{color:var(--accent); font-size:10px; margin-left:4px;}
  table.leads tbody tr:nth-child(odd){background:var(--surface-2);}
  table.leads td.lead-cell{font-weight:700; color:var(--ink); white-space:nowrap;}
  table.leads td.time-cell{white-space:nowrap;}
  table.leads td.remarks-cell{min-width:220px; max-width:320px; white-space:normal;}
  table.leads td.loc-cell{min-width:220px; max-width:280px; white-space:normal; color:var(--ink-secondary); font-size:12px;}

  .tag{display:inline-flex; align-items:center; gap:5px; padding:3px 11px; font-size:11px; font-weight:700; border-radius:999px; white-space:nowrap;}
  .tag.new{background:var(--good-bg); color:var(--good);}
  .tag.existing{background:var(--surface-3); color:var(--ink-secondary);}
  .tag.overdue{background:var(--critical-bg); color:var(--critical);}
  .tag.unnamed{background:var(--critical-bg); color:var(--critical);}
  .tag.assumed{background:var(--watch-bg); color:var(--ink-secondary);}
  .tag.notinterested{background:var(--surface-3); color:var(--ink-secondary); border:1px solid var(--border);}
  /* Visit history drop-down (Section 4, 24-Sep-2026) */
  .hist-btn{display:inline-flex; align-items:center; gap:5px; margin-top:6px; padding:3px 10px; font-size:11px; font-weight:700; border-radius:999px; border:1px solid var(--border); background:var(--surface); color:var(--ink-secondary); cursor:pointer; white-space:nowrap; font-family:inherit;}
  .hist-btn:hover{background:var(--surface-3); color:var(--ink);}
  .hist-btn .chev{display:inline-block; font-size:10px; transition:transform .15s ease;}
  .hist-btn[aria-expanded="true"]{background:var(--ink); color:#fff; border-color:var(--ink);}
  .hist-btn[aria-expanded="true"] .chev{transform:rotate(90deg);}
  table.leads tr.history-row > td{background:var(--surface-3); padding:12px 14px 16px 14px; border-bottom:2px solid var(--border);}
  .hist-title{font-size:11.5px; color:var(--ink-secondary); margin:0 0 8px 0;}
  .hist-title strong{color:var(--ink);}
  .hist-swatch{display:inline-block; width:10px; height:10px; border-radius:2px; vertical-align:-1px; margin:0 3px 0 6px;}
  .hist-scroll{overflow-x:auto; border-radius:8px; border:1px solid var(--border);}
  table.leads table.hist-table{border-collapse:collapse; width:100%; font-size:12px; background:var(--surface);}
  table.leads table.hist-table thead th{background:var(--surface-2); color:var(--ink-secondary); font-size:10.5px; cursor:default; padding:7px 10px; border-bottom:1px solid var(--border);}
  table.leads table.hist-table thead th:hover{background:var(--surface-2);}
  table.leads table.hist-table td{padding:7px 10px; background:var(--surface);}
  table.leads table.hist-table tr.hl td{background:var(--good-bg);}
  table.leads table.hist-table tr.hl-over td{background:var(--critical-bg);}
  table.leads table.hist-table td.visit-no{font-weight:700; color:var(--ink-muted); width:28px;}
  .tag.kept{background:var(--good-bg); color:var(--good);}
  .tag.sameday{background:var(--amber-bg); color:var(--amber);}
  .tag.late{background:var(--amber-bg); color:var(--amber);}
  .tag.pending{background:var(--surface-3); color:var(--ink-secondary);}
  .pdf-btn{background:var(--ink); color:#fff; border-color:var(--ink);}
  .pdf-btn:hover{background:#2a2a2a; color:#fff;}
  .pdf-btn:disabled{background:var(--surface-2); color:var(--ink-muted); border-color:var(--border); cursor:not-allowed;}
  .pdf-btn .icon{width:13px; height:13px;}
  .empty-state{padding:40px; text-align:center; color:var(--ink-muted); font-size:13px;}

  /* Lead Mix row (restructured 23-Sep-2026 per Harsh's request): Timeline,
     Lead Type and the Value pie all in one line instead of two bar panels
     with the pie stacked below - removes the need to scroll down to reach
     the pie before filtering. breakdown-grid keeps its own 2-col layout for
     the two bar panels; lead-mix-row just places it beside the pie. The two
     children are ordinary siblings (not nested into one grid), so the pie's
     SVG node is never touched by breakdownGrid's own re-render - clicking a
     bar or slice can't wipe the pie out. Stacks back to one column below
     980px so it doesn't squeeze on a phone. */
  .lead-mix-row{display:flex; gap:24px; align-items:flex-start; margin-top:6px;}
  .lead-mix-row .breakdown-grid{flex:1 1 66%; min-width:0; margin-top:0;}
  .lead-mix-row .mix-pie-row{flex:1 1 34%; min-width:0;}
  .breakdown-grid{display:grid; grid-template-columns:repeat(2,1fr); gap:24px; margin-top:6px;}
  .breakdown-panel h3{font-size:10.5px; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; color:var(--ink-muted); margin:0 0 6px 0;}
  /* Restyled 23-Sep-2026: one compact line per row (icon + label + thin bar
     + count), replacing the old stacked label-then-bar layout, so the same
     information reads at a glance instead of across two lines. The click-
     to-expand behaviour underneath (data-field/data-value, active state,
     shared detail panel below) is unchanged. */
  .bar-row{display:flex; align-items:center; gap:10px; padding:5px 6px; margin:0 -6px; border-radius:8px; cursor:pointer; border:1px solid transparent; transition:background-color .12s ease, border-color .12s ease;}
  .bar-row:hover{background:var(--surface-3);}
  .bar-row:focus-visible{outline:2px solid var(--accent); outline-offset:1px;}
  .bar-row.bar-row-active{background:var(--critical-bg); border-color:var(--accent-deep);}
  .bar-row .bar-icon{width:26px; height:26px; border-radius:7px; flex:none; display:flex; align-items:center; justify-content:center;}
  .bar-row .bar-icon .icon{width:13px; height:13px;}
  .bar-row .bar-label{flex:0 0 148px; font-size:12.5px; color:var(--ink); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
  .bar-track{flex:1; background:var(--surface-3); height:7px; border-radius:4px; position:relative;}
  .bar-fill{display:block; background:var(--accent); height:7px; border-radius:4px;}
  .bar-row .bar-count{flex:none; width:82px; text-align:right; font-size:12px; font-weight:700; color:var(--ink-secondary);}

  /* ---------- Lead Mix pie (Estimated Project Value) - bigger chart with
     leader-line labels on the chart itself rather than a side legend, and
     wired into the same combined-filtering mechanism as the two bar panels
     above (see renderPie()/activeFilters in the script). ---------- */
  .mix-pie-row{margin-top:0;}
  .pie-slice{cursor:pointer; transition:opacity .15s ease; outline:none;}
  .pie-slice:hover{opacity:0.85;}
  .pie-slice.active{stroke:var(--ink); stroke-width:1.6;}
  .pie-slice.dim{opacity:0.35;}
  .pie-label{cursor:pointer; outline:none;}
  .pie-label:hover .pie-label-cat{text-decoration:underline;}
  .pie-label-cat{font-size:11px; font-weight:700; fill:var(--ink); font-family:'Product Sans','Open Sans',sans-serif;}
  .pie-label-val{font-size:11px; font-weight:700; font-family:'Product Sans','Open Sans',sans-serif;}
  .pie-label.dim .pie-label-cat, .pie-label.dim .pie-label-val{opacity:0.35;}

  footer{margin-top:26px; border-top:1px solid var(--border); padding-top:16px; font-size:11.5px; color:var(--ink-muted); line-height:1.6;}
  footer strong{color:var(--ink-secondary);}

  @media (max-width: 980px){
    .lead-mix-row{flex-direction:column;}
    .lead-mix-row .mix-pie-row{margin-top:20px; padding-top:18px; border-top:1px solid var(--border-soft); width:100%;}
    .breakdown-grid{grid-template-columns:1fr;}
    .overview-row{flex-direction:column;}
    .kpi-strip{min-width:0;}
    .wrap{padding:16px 18px 32px 18px;}
    .masthead{padding:14px 18px;}
    .quicknav{padding:0 18px;}
    section.block{padding:16px 16px 18px;}
  }

  /* ---------- Phone layout (24-Sep-2026, per Harsh). Everything below is
     phone-only (640px and narrower); the desktop view is unchanged.
     1. Frozen header cut from about 400px to about 115px: logo and title on
        one line, a Filters button that opens the three drop-downs, a one-
        line summary of the current selection, and a single swipeable line
        of section links.
     2. Every table turns into cards: one card per row, each field shown as
        label and value. Labels come from the table's own header cells
        (see labelTable in the script), so a new column needs no CSS change.
     3. Long lists show 20 cards, then a "Show 20 more" button. Search and
        filters still run over every row. ---------- */
  .m-filters-btn, .m-filter-summary, .m-more{display:none;}
  @media (max-width: 640px){
    section.block{scroll-margin-top:124px;}
    .masthead{padding:10px 14px 9px;}
    .masthead-inner{gap:6px 10px; flex-wrap:wrap; align-items:center;}
    .masthead-left{flex:1 1 0; min-width:0; gap:10px;}
    .wordmark-logo{height:18px;}
    .masthead-divider{height:24px;}
    .masthead-title h1{font-size:13.5px; line-height:1.25;}
    .m-filters-btn{
      display:inline-flex; align-items:center; gap:6px; flex:none; font-family:inherit; font-size:12.5px; font-weight:700;
      color:#fff; background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.28); border-radius:8px; padding:7px 11px; cursor:pointer;
    }
    .m-filters-btn .icon{width:13px; height:13px;}
    .m-filters-btn .m-count{display:none; min-width:17px; height:17px; padding:0 5px; border-radius:999px; background:var(--accent); font-size:10.5px; line-height:17px; text-align:center;}
    .m-filters-btn.has-active .m-count{display:inline-block;}
    .m-filters-btn[aria-expanded="true"]{background:#fff; color:var(--ink); border-color:#fff;}
    .m-filter-summary{display:block; flex-basis:100%; font-size:11.5px; color:rgba(255,255,255,0.62); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
    .masthead-right{display:none; flex-basis:100%;}
    .masthead.filters-open .masthead-right{display:block; padding:4px 0 2px;}
    .masthead.filters-open .m-filter-summary{display:none;}
    .masthead-selector{display:grid; grid-template-columns:minmax(0,1fr); gap:8px;}
    select#monthSelect, select#weekSelect, select#execSelect{min-width:0; width:100%; font-size:16px; padding:9px 32px 9px 12px;}

    .quicknav{padding:0 4px;}
    .quicknav-inner{flex-wrap:nowrap; overflow-x:auto; scrollbar-width:none; -webkit-overflow-scrolling:touch;}
    .quicknav-inner::-webkit-scrollbar{display:none;}
    .quicknav a{flex:none; white-space:nowrap; padding:10px 11px; font-size:11px;}

    .wrap{padding:12px 12px 28px;}
    section.block{padding:14px 12px 16px; border-radius:14px; margin-top:12px;}
    .block-head{gap:10px;}
    .block-head-left{gap:10px;}
    .icon-badge{width:32px; height:32px; border-radius:9px;}
    .icon-badge .icon{width:16px; height:16px;}
    .block-head h2{font-size:15.5px;}
    .block-hint{white-space:normal; align-self:flex-start;}
    .block-desc{font-size:12.5px; margin:6px 0 10px;}

    .kpi-strip{display:grid; grid-template-columns:repeat(3, minmax(0,1fr)); gap:8px; min-width:0; flex:none; width:100%;}
    .kpi-mini{min-width:0; padding:10px 10px 11px;}
    .kpi-mini .icon-badge{width:28px; height:28px; border-radius:8px;}
    .kpi-mini .icon-badge .icon{width:14px; height:14px;}
    .kpi-mini .kv{font-size:21px;}
    .kpi-mini .kl{font-size:9.5px; letter-spacing:0.04em;}
    .donut-wrap{width:100%; padding:10px 14px;}

    .lead-mix-row{align-items:stretch;}
    .lead-mix-row .breakdown-grid{width:100%;}
    .bar-row .bar-label{flex:0 0 118px; font-size:12px;}
    .bar-row .bar-count{width:auto; min-width:66px; font-size:11.5px;}

    .pill-row{gap:8px;}
    .pill{min-width:0; flex:1 1 140px; padding:11px 12px; gap:10px;}
    .pill .pv{font-size:20px;}
    .pill .pl{font-size:11px;}

    .table-toolbar{margin:12px 0 10px; gap:8px;}
    .table-toolbar input[type=search]{min-width:0; width:100%; font-size:16px;}

    /* Cards */
    .table-wrap{overflow:visible; border:none; border-radius:0; margin-bottom:0;}
    table.leads{display:block; min-width:0; width:100%; font-size:12.5px;}
    table.leads > thead{display:none;}
    table.leads > tbody{display:block;}
    table.leads > tbody > tr{
      display:block; background:var(--surface) !important; border:1px solid var(--border); border-radius:12px;
      padding:10px 12px 6px; margin:0 0 10px; box-shadow:0 1px 2px rgba(22,22,22,0.04);
    }
    table.leads > tbody > tr > td{
      display:flow-root; padding:5px 0 5px 112px; border-bottom:1px solid var(--border-soft);
      min-width:0 !important; max-width:none !important; white-space:normal !important; overflow-wrap:anywhere; font-size:12.5px;
    }
    table.leads > tbody > tr > td:last-child{border-bottom:none;}
    table.leads > tbody > tr > td::before{
      content:attr(data-label); float:left; width:102px; margin-left:-112px; padding-top:1px;
      font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; line-height:1.4; color:var(--ink-muted);
      overflow-wrap:normal;
    }
    table.leads > tbody > tr > td:first-child{padding:0 0 8px; margin-bottom:2px; font-size:14.5px; font-weight:700; color:var(--ink); border-bottom:1px solid var(--border);}
    table.leads > tbody > tr > td:first-child::before{content:none;}
    table.leads > tbody > tr > td:empty{display:none;}
    /* Check-in time sits under the lead name as a subtitle, not as a row */
    table.leads > tbody > tr > td:first-child:has(+ td[data-label="Time"]){border-bottom:none; padding-bottom:1px; margin-bottom:0;}
    table.leads > tbody > tr > td[data-label="Time"]:nth-child(2){padding:0 0 8px; margin-bottom:2px; font-size:11.5px; color:var(--ink-muted); border-bottom:1px solid var(--border);}
    table.leads > tbody > tr > td[data-label="Time"]:nth-child(2)::before{content:none;}
    table.leads > tbody > tr > td.empty-state{padding:22px 6px; border:none; font-size:13px; font-weight:400; color:var(--ink-muted);}
    table.leads > tbody > tr > td.empty-state::before{content:none;}
    table.leads > tbody > tr.m-hidden{display:none;}
    table.leads > tbody > tr > td .hist-btn{margin-top:8px;}

    /* Visit history under a card */
    table.leads > tbody > tr.history-row{padding:0; border:none; box-shadow:none; background:transparent !important; margin:-4px 0 12px;}
    table.leads > tbody > tr.history-row > td{display:block; padding:10px; border:1px solid var(--border); border-radius:12px; background:var(--surface-3); font-size:12px; font-weight:400;}
    table.leads > tbody > tr.history-row > td::before{content:none;}
    .hist-scroll{overflow:visible; border:none;}
    table.leads table.hist-table{display:block; background:transparent;}
    table.leads table.hist-table > thead{display:none;}
    table.leads table.hist-table > tbody{display:block;}
    table.leads table.hist-table > tbody > tr{display:block; background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:8px 10px 4px; margin-bottom:8px;}
    table.leads table.hist-table > tbody > tr.hl{background:var(--good-bg); border-color:rgba(47,143,91,0.35);}
    table.leads table.hist-table > tbody > tr.hl-over{background:var(--critical-bg); border-color:rgba(178,58,46,0.35);}
    table.leads table.hist-table > tbody > tr > td{display:flow-root; background:transparent !important; padding:4px 0 4px 100px; border-bottom:1px solid var(--border-soft); white-space:normal; overflow-wrap:anywhere; width:auto;}
    table.leads table.hist-table > tbody > tr > td:last-child{border-bottom:none;}
    table.leads table.hist-table > tbody > tr > td::before{
      content:attr(data-label); float:left; width:92px; margin-left:-100px;
      font-size:9.5px; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; line-height:1.4; color:var(--ink-muted);
    }
    table.leads table.hist-table > tbody > tr > td.visit-no{padding:0 0 6px; width:auto !important; font-size:12.5px; color:var(--ink);}
    table.leads table.hist-table > tbody > tr > td.visit-no::before{content:"Visit "; float:none; width:auto; margin:0; font-size:inherit; text-transform:none; letter-spacing:0; color:inherit; line-height:inherit;}

    /* Section 1 detail list */
    .detail-panel-scroll{max-height:70vh; padding:10px 10px 2px;}
    .detail-panel-head{padding:10px 12px;}

    .m-more{
      display:flex; align-items:center; justify-content:center; width:100%; margin:2px 0 8px; padding:11px 14px;
      font-family:inherit; font-size:12.5px; font-weight:700; color:var(--ink); background:var(--surface-2);
      border:1px solid var(--border); border-radius:10px; cursor:pointer;
    }
    .m-more:active{background:var(--surface-3);}
    .detail-panel .m-more{width:calc(100% - 20px); margin:0 10px 10px;}
  }
</style>
</head>
<body>

<svg width="0" height="0" style="position:absolute" aria-hidden="true">
<defs>
<g fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
  <symbol id="icon-clipboard-check" viewBox="0 0 24 24">
    <rect x="5" y="4" width="14" height="17" rx="2"/>
    <path d="M9 4V3a1 1 0 011-1h4a1 1 0 011 1v1"/>
    <path d="M8.5 12.6l2.3 2.3 4.7-4.8"/>
  </symbol>
  <symbol id="icon-alert-triangle" viewBox="0 0 24 24">
    <path d="M12 3.5l9.5 16.5H2.5L12 3.5z"/>
    <path d="M12 9.5v5"/>
    <circle cx="12" cy="17.2" r="0.9" fill="currentColor" stroke="none"/>
  </symbol>
  <symbol id="icon-bar-chart" viewBox="0 0 24 24">
    <rect x="4" y="12" width="3.4" height="8" rx="0.8"/>
    <rect x="10.3" y="6" width="3.4" height="14" rx="0.8"/>
    <rect x="16.6" y="9.5" width="3.4" height="10.5" rx="0.8"/>
  </symbol>
  <symbol id="icon-percent" viewBox="0 0 24 24">
    <path d="M5 19L19 5"/>
    <circle cx="7.5" cy="7.5" r="2.3"/>
    <circle cx="16.5" cy="16.5" r="2.3"/>
  </symbol>
  <symbol id="icon-inbox" viewBox="0 0 24 24">
    <path d="M3 12h5l1.6 3h4.8l1.6-3h5"/>
    <path d="M4.5 12L6 5.5A2 2 0 018 4h8a2 2 0 011.9 1.5L19.5 12"/>
    <path d="M4.5 12v6a2 2 0 002 2h11a2 2 0 002-2v-6"/>
  </symbol>
  <symbol id="icon-users" viewBox="0 0 24 24">
    <circle cx="9" cy="8" r="3.2"/>
    <path d="M3 20c0-3.8 2.7-6.4 6-6.4s6 2.6 6 6.4"/>
    <circle cx="17" cy="8.4" r="2.4"/>
    <path d="M15.5 14c2.4.7 4.2 2.9 4.5 6"/>
  </symbol>
  <symbol id="icon-repeat" viewBox="0 0 24 24">
    <path d="M4 7h13a3 3 0 013 3v1"/>
    <path d="M17 4l3 3-3 3"/>
    <path d="M20 17H7a3 3 0 01-3-3v-1"/>
    <path d="M7 20l-3-3 3-3"/>
  </symbol>
  <symbol id="icon-check-circle" viewBox="0 0 24 24">
    <circle cx="12" cy="12" r="8.5"/>
    <path d="M8.3 12.4l2.5 2.5 5-5.4"/>
  </symbol>
  <symbol id="icon-calendar-x" viewBox="0 0 24 24">
    <rect x="4" y="5" width="16" height="15" rx="2"/>
    <path d="M8 3v4M16 3v4M4 10h16"/>
    <path d="M9.5 14l5 4M14.5 14l-5 4"/>
  </symbol>
  <!-- Added 23-Sep-2026 for the graphical/less-wordy redesign: category
       icons for the Lead Mix panels (When they need it / Who they are /
       What it's worth) and the Missed Days day-type column. -->
  <symbol id="icon-clock" viewBox="0 0 24 24">
    <circle cx="12" cy="12" r="8.5"/>
    <path d="M12 7v5l3.5 2"/>
  </symbol>
  <symbol id="icon-arch" viewBox="0 0 24 24">
    <path d="M4 20h16"/>
    <path d="M6 20V10l6-6 6 6v10"/>
    <path d="M10 20v-6h4v6"/>
  </symbol>
  <symbol id="icon-briefcase" viewBox="0 0 24 24">
    <rect x="3" y="8" width="18" height="12" rx="2"/>
    <path d="M8 8V6a2 2 0 012-2h4a2 2 0 012 2v2"/>
  </symbol>
  <symbol id="icon-hardhat" viewBox="0 0 24 24">
    <path d="M4 16a8 8 0 0116 0z"/>
    <path d="M2 16h20"/>
    <path d="M12 8v-2"/>
  </symbol>
  <symbol id="icon-store" viewBox="0 0 24 24">
    <path d="M4 9l1-5h14l1 5"/>
    <path d="M4 9a2 2 0 004 0 2 2 0 004 0 2 2 0 004 0 2 2 0 004 0"/>
    <path d="M5 9v10h14V9"/>
  </symbol>
  <symbol id="icon-bulb" viewBox="0 0 24 24">
    <path d="M9 18h6"/>
    <path d="M10 21h4"/>
    <path d="M12 3a6 6 0 00-3 11.2c.6.4 1 1.1 1 1.8h4c0-.7.4-1.4 1-1.8A6 6 0 0012 3z"/>
  </symbol>
  <symbol id="icon-truck" viewBox="0 0 24 24">
    <rect x="1" y="7" width="13" height="10" rx="1"/>
    <path d="M14 10h4l4 4v3h-8z"/>
    <circle cx="6" cy="19" r="1.6"/>
    <circle cx="17" cy="19" r="1.6"/>
  </symbol>
  <symbol id="icon-sun" viewBox="0 0 24 24">
    <circle cx="12" cy="12" r="4"/>
    <path d="M12 3v2M12 19v2M4.2 4.2l1.4 1.4M18.4 18.4l1.4 1.4M3 12h2M19 12h2M4.2 19.8l1.4-1.4M18.4 5.6l1.4-1.4"/>
  </symbol>
  <symbol id="icon-star" viewBox="0 0 24 24">
    <path d="M12 3l2.6 5.6 6.1.6-4.6 4.1 1.3 6-5.4-3.1-5.4 3.1 1.3-6-4.6-4.1 6.1-.6z"/>
  </symbol>
  <symbol id="icon-question" viewBox="0 0 24 24">
    <circle cx="12" cy="12" r="9"/>
    <path d="M9.2 9.5a2.8 2.8 0 015.4.9c0 1.8-2.6 1.6-2.6 3.6"/>
    <circle cx="12" cy="17.5" r="0.6" fill="currentColor" stroke="none"/>
  </symbol>
</g>
</defs>
</svg>

<div class="sticky-head">
  <div class="masthead">
    <div class="masthead-inner">
      <div class="masthead-left">
        <img class="wordmark-logo" src="__LOGO_DATA_URI__" alt="Pasolite" height="26">
        <div class="masthead-divider"></div>
        <div class="masthead-title">
          <h1>Field Visit &amp; Lead Dashboard</h1>
        </div>
      </div>
      <button type="button" class="m-filters-btn" id="filtersBtn" aria-expanded="false" aria-controls="mastheadRight">
        <svg class="icon" viewBox="0 0 24 24"><path d="M4 6h16M7 12h10M10 18h4" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        Filters <span class="m-count" id="filtersCount"></span>
      </button>
      <div class="m-filter-summary" id="filterSummary"></div>
      <div class="masthead-right" id="mastheadRight">
        <div class="masthead-selector">
          <div class="control-group"><label class="control-label" for="monthSelect">Month</label><select id="monthSelect"></select></div>
          <div class="control-group"><label class="control-label" for="weekSelect">Week</label><select id="weekSelect"></select></div>
          <div class="control-group"><label class="control-label" for="execSelect">Sales Executive</label><select id="execSelect"></select></div>
        </div>
      </div>
    </div>
  </div>

  <div class="quicknav">
    <div class="quicknav-inner">
      <a href="#overview">Overview</a>
      <a href="#breakdown">Lead Mix</a>
      <a href="#newleads">New Leads</a>
      <a href="#existingleads">Existing Leads</a>
      <a href="#stale">Follow-up Watch</a>
      <a href="#nocheckin">Missed Days</a>
    </div>
  </div>
</div>

<div class="wrap">

  <section class="block" id="overview">
    <div class="block-head">
      <div class="block-head-left">
        <span class="icon-badge tone-teal"><svg class="icon"><use href="#icon-clipboard-check"></use></svg></span>
        <div>
          <div class="block-eyebrow">This Period at a Glance</div>
          <h2>Forms, new leads &amp; existing leads</h2>
        </div>
      </div>
    </div>
    <div class="overview-row" id="overviewRow"></div>
    <button type="button" class="info-btn" data-target="reconcilePanel" style="margin-top:12px;">
      <svg class="icon"><use href="#icon-question"></use></svg> How this reconciles
    </button>
    <div class="info-panel" id="reconcilePanel"></div>
  </section>

  <section class="block" id="breakdown">
    <div class="block-head">
      <div class="block-head-left">
        <span class="icon-badge tone-violet"><svg class="icon"><use href="#icon-bar-chart"></use></svg></span>
        <div>
          <div class="block-eyebrow">Section 1</div>
          <h2>Lead Mix &mdash; when &amp; who</h2>
        </div>
      </div>
      <span class="block-hint">Click any row or slice to filter &mdash; combine across all three</span>
    </div>
    <div class="lead-mix-row">
      <div class="breakdown-grid" id="breakdownGrid"></div>
      <div class="mix-pie-row">
        <h3 style="font-size:10.5px; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; color:var(--ink-muted); margin:0 0 6px 0;">Estimated Project Value</h3>
        <svg id="pieSvg" viewBox="-40 0 400 260" style="width:100%; height:auto; display:block; margin:0 auto;" aria-label="Estimated project value, share of leads"></svg>
      </div>
    </div>
    <div class="detail-panel" id="breakdownDetail" style="display:none;"></div>
  </section>

  <section class="block" id="newleads">
    <div class="block-head">
      <div class="block-head-left">
        <span class="icon-badge tone-good"><svg class="icon"><use href="#icon-users"></use></svg></span>
        <div>
          <div class="block-eyebrow">Section 2</div>
          <h2>New Leads</h2>
        </div>
      </div>
    </div>
    <p class="block-desc">First time this lead's name has shown up in tracked data.</p>
    <div class="table-toolbar">
      <input type="search" id="newLeadsSearch" placeholder="Search lead, location, remarks...">
      <span class="row-count" id="newLeadsRowCount"></span>
    </div>
    <div class="table-wrap">
      <table class="leads" id="newLeadsTable">
        <thead><tr id="newLeadsHeadRow"></tr></thead>
        <tbody id="newLeadsBody"></tbody>
      </table>
    </div>
  </section>

  <section class="block" id="existingleads">
    <div class="block-head">
      <div class="block-head-left">
        <span class="icon-badge tone-accent"><svg class="icon"><use href="#icon-repeat"></use></svg></span>
        <div>
          <div class="block-eyebrow">Section 3</div>
          <h2>Existing Leads</h2>
        </div>
      </div>
    </div>
    <p class="block-desc">This lead's name has appeared in an earlier tracked week.</p>
    <div class="table-toolbar">
      <input type="search" id="existingLeadsSearch" placeholder="Search lead, location, remarks...">
      <span class="row-count" id="existingLeadsRowCount"></span>
    </div>
    <div class="table-wrap">
      <table class="leads" id="existingLeadsTable">
        <thead><tr id="existingLeadsHeadRow"></tr></thead>
        <tbody id="existingLeadsBody"></tbody>
      </table>
    </div>
  </section>

  <section class="block" id="stale">
    <div class="block-head">
      <div class="block-head-left">
        <span class="icon-badge tone-accent"><svg class="icon"><use href="#icon-alert-triangle"></use></svg></span>
        <div>
          <div class="block-eyebrow">Section 4</div>
          <h2>Follow-up Watch</h2>
        </div>
      </div>
      <button type="button" class="info-btn" data-target="staleInfoPanel">
        <svg class="icon"><use href="#icon-question"></use></svg> How this works
      </button>
    </div>
    <div class="info-panel" id="staleInfoPanel"></div>
    <div class="pill-row" id="stalePillRow"></div>

    <div class="block-head" style="margin-top:20px;">
      <div class="block-head-left">
        <span class="icon-badge sm tone-good"><svg class="icon"><use href="#icon-check-circle"></use></svg></span>
        <div><h2 style="font-size:14px;">Followed Up On Time</h2></div>
      </div>
    </div>
    <p class="block-desc" id="keptCaption"></p>
    <div class="table-wrap">
      <table class="leads" id="keptTable">
        <thead><tr>
          <th>Lead</th><th>Time</th><th>Est. Project Value</th><th>Type of Lead</th>
          <th>Location</th><th>Requirement Timeline</th><th>Next Follow-up</th>
          <th>Sales Executive</th><th>Client Satisfaction</th><th>Remarks</th>
        </tr></thead>
        <tbody id="keptBody"></tbody>
      </table>
    </div>
    <p class="reconcile" id="keptNote"></p>

    <div class="block-head" style="margin-top:22px;">
      <div class="block-head-left">
        <span class="icon-badge sm tone-accent"><svg class="icon"><use href="#icon-alert-triangle"></use></svg></span>
        <div><h2 style="font-size:14px;">Overdue Follow-ups</h2></div>
      </div>
      <button type="button" class="info-btn pdf-btn" id="overduePdfBtn" disabled>
        <svg class="icon" viewBox="0 0 24 24"><path d="M12 3v12m0 0l-5-5m5 5l5-5M5 21h14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
        <span id="overduePdfLabel">Pick a sales executive to download</span>
      </button>
    </div>
    <p class="block-desc" id="staleOverdueCaption"></p>
    <div class="table-wrap">
      <table class="leads" id="staleOverdueTable">
        <thead><tr>
          <th>Lead</th><th>Time</th><th>Est. Project Value</th><th>Type of Lead</th>
          <th>Location</th><th>Requirement Timeline</th><th>Next Follow-up</th>
          <th>Due Week</th><th>Days Overdue</th>
          <th>Sales Executive</th><th>Client Satisfaction</th><th>Remarks</th>
        </tr></thead>
        <tbody id="staleOverdueBody"></tbody>
      </table>
    </div>
    <p class="reconcile" id="staleOverdueNote"></p>

    <div id="colleagueBlock" style="display:none;">
      <div class="block-head" style="margin-top:22px;">
        <div class="block-head-left">
          <span class="icon-badge sm tone-watch"><svg class="icon"><use href="#icon-users"></use></svg></span>
          <div><h2 style="font-size:14px;" id="colleagueTitle">Followed Up by a Colleague</h2></div>
        </div>
      </div>
      <p class="block-desc" id="colleagueCaption"></p>
      <div class="table-wrap">
        <table class="leads" id="colleagueTable">
          <thead><tr>
            <th>Lead</th><th>Time</th><th>Est. Project Value</th><th>Type of Lead</th>
            <th>Location</th><th>Requirement Timeline</th><th>Next Follow-up</th>
            <th>Due By</th><th>Followed Up By</th><th>Their Visit</th><th>Timing</th>
            <th>Client Satisfaction</th><th>Remarks</th>
          </tr></thead>
          <tbody id="colleagueBody"></tbody>
        </table>
      </div>
      <p class="reconcile" id="colleagueNote"></p>
    </div>
  </section>

  <section class="block" id="nocheckin">
    <div class="block-head">
      <div class="block-head-left">
        <span class="icon-badge tone-accent"><svg class="icon"><use href="#icon-calendar-x"></use></svg></span>
        <div>
          <div class="block-eyebrow">Section 5</div>
          <h2>Missed Days</h2>
        </div>
      </div>
      <button type="button" class="info-btn" data-target="noCheckinInfoPanel">
        <svg class="icon"><use href="#icon-question"></use></svg> How this works
      </button>
    </div>
    <div class="info-panel" id="noCheckinInfoPanel"></div>
    <div class="pill-row" id="noCheckinPillRow"></div>
    <div class="table-wrap">
      <table class="leads" id="noCheckinTable">
        <thead><tr><th>Sales Executive</th><th>Date</th><th>Day Type</th><th>Week</th></tr></thead>
        <tbody id="noCheckinBody"></tbody>
      </table>
    </div>
    <p class="reconcile" id="noCheckinNote"></p>
  </section>

  <footer id="footerNote"></footer>
  <button type="button" class="info-btn" data-target="footerInfoPanel" style="margin-top:10px;">
    <svg class="icon"><use href="#icon-question"></use></svg> Full methodology &amp; caveats
  </button>
  <div class="info-panel" id="footerInfoPanel"></div>

</div>

__PDF_LIBS__
<script>
const DASHBOARD_DATA = __DASHBOARD_DATA__;

(function () {
  const DATA = DASHBOARD_DATA;

  // Shared column set for the New Leads / Existing Leads tables (Sections 2
  // and 3). No Status column here - each table already holds only its own
  // kind of lead, so a New/Existing tag on every row would be redundant.
  // Client Satisfaction added 23-Sep-2026 per Harsh's request, same field
  // already audited and shown in Section 4 - Not Interested rendered as a
  // tag (see renderLeadsTableBody below), the other two values shown plain.
  //
  // Column order standardised across every lead table on the dashboard
  // (23-Sep-2026, per Harsh): Lead, Time, Est. Project Value, Type of Lead,
  // Location, Requirement Timeline, Next Follow-up - then whatever else a
  // given table carries. Lead Mix detail and both Section 4 tables follow
  // the same seven-column core; keep them in step if this ever changes.
  const LEADS_COLUMNS = [
    { key: 'lead', label: 'Lead', cls: 'lead-cell' },
    { key: 'time', label: 'Time', cls: 'time-cell' },
    { key: 'value', label: 'Est. Project Value' },
    { key: 'leadType', label: 'Type of Lead' },
    { key: 'location', label: 'Location', cls: 'loc-cell' },
    { key: 'timeline', label: 'Requirement Timeline' },
    { key: 'followUp', label: 'Next Follow-up' },
    { key: 'satisfaction', label: 'Client Satisfaction' },
    { key: 'remarks', label: 'Remarks', cls: 'remarks-cell' },
  ];

  // Independent sort state (and lead-kind filter) per table, keyed by
  // 'newLeads' / 'existingLeads' - matches the id prefixes used for each
  // table's DOM elements (e.g. newLeadsHeadRow, existingLeadsBody).
  const leadsTables = {
    newLeads: { sortKey: 'time', sortDir: 1, filterFn: (r) => r.isNew },
    existingLeads: { sortKey: 'time', sortDir: 1, filterFn: (r) => r.isExisting },
  };

  function fmtNum(n) { return n.toLocaleString('en-IN'); }
  function pct(n, d) { return d ? Math.round((n / d) * 1000) / 10 : 0; }
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  // ---- lead key helper ---------------------------------------------------
  // Normalized key used to cross-check a lead against its own visit history
  // (see leadVisitCounts()/computeStaleCandidates() below). The browser-
  // local "Mark done" / Leads Completed mechanism that used to sit alongside
  // this was removed 23-Sep-2026 per Harsh's request: a mark that resets
  // whenever the dashboard is rebuilt with a new week's data was judged not
  // useful enough to keep, so there is no persisted lead status any more.
  function leadKey(name) { return (name || '').trim().toLowerCase(); }

  // ---- Best-effort follow-up date parser -----------------------------
  // Real LystLoc data: ~90 distinct raw values for "Next Follow-up Date"
  // across 3 weeks. Most are relative phrases ("Next week", "Follow Up",
  // "Month end"...), a handful are literal dates, and a handful look like
  // the wrong field was typed in ("Dealer", "New", "Order received"). This
  // parser handles literal dates and the common relative phrases found in
  // the real data, computed from the visit date; anything it doesn't
  // recognise is reported as "not specified" rather than guessed - it is
  // never silently treated as on-track or overdue.
  const MONTH_ABBR = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  const WORDNUM = { one: 1, two: 2, three: 3, four: 4, five: 5 };

  function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
  function addMonths(d, n) { const r = new Date(d); r.setMonth(r.getMonth() + n); return r; }
  function endOfMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0); }
  function endOfWeek(d) { const day = d.getDay(); const diff = day === 0 ? 0 : 7 - day; return addDays(d, diff); }
  function toNum(tok) { return WORDNUM[tok] !== undefined ? WORDNUM[tok] : parseInt(tok, 10); }

  // Monday-start week helpers, mirroring the build script's own Python
  // week_start()/month_label()/week_label() exactly, so a week bucket
  // computed here lines up with DATA.weeks entries built at build time.
  function mondayOf(d) { const day = d.getDay(); const diff = day === 0 ? -6 : 1 - day; return addDays(d, diff); }
  function ymd(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
  function monthLabelJS(d) { return d.toLocaleString('en-US', { month: 'long' }) + ' ' + d.getFullYear(); }
  function weekLabelJS(monday) {
    const sunday = addDays(monday, 6);
    if (monday.getMonth() === sunday.getMonth()) {
      return monday.getDate() + ' - ' + sunday.getDate() + ' ' + monday.toLocaleString('en-US', { month: 'short' }) + ' ' + monday.getFullYear();
    }
    const dm = (d) => String(d.getDate()).padStart(2, '0') + ' ' + d.toLocaleString('en-US', { month: 'short' });
    return dm(monday) + ' - ' + dm(sunday) + ' ' + sunday.getFullYear();
  }

  // ---- Follow-up date parser (v2, 23-Sep-2026) -------------------------
  // Reads what the rep wrote in "Next Follow-up" and returns the DEADLINE
  // the promise implies. v2 principle, per Harsh: a period is read as a
  // period. "Next week" means by the end of next week, "next month" means
  // by the end of next month - never a single exact day, so a vaguer entry
  // is never judged more harshly than the no-date fallback (which gives
  // until the Sunday of the week after the visit). Durations the rep
  // spelled out ("in 2 days", "after 2 month") are still read as exact.
  // Anything unrecognised returns no date; callers then apply the assumed
  // week-after-visit rule and tag the row "assumed". A typo'd year (2016)
  // is kept exactly as entered - never silently corrected.
  const DAY_NAMES = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
  const DAY_RE = /\b(sunday|monday|tuesday|wednesday|thursday|friday|saturday|sun|mon|tues|tue|wed|thurs|thur|thu|fri|sat)\b/g;
  const MONTH_RE = '(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*';
  const ORD_WEEK = { first: 1, '1st': 1, '1': 1, second: 2, '2nd': 2, '2': 2, third: 3, '3rd': 3, '3': 3, fourth: 4, '4th': 4, '4': 4 };

  function sundayOf(d) { return addDays(mondayOf(d), 6); }
  function monthIndex(tok) { return MONTH_ABBR.indexOf(tok.slice(0, 3)); }
  // The named month on or after the visit month (an August visit saying
  // "October" means this October; a December visit saying "January" means
  // next January).
  function namedMonthStart(mi, visit) {
    let y = visit.getFullYear();
    if (mi < visit.getMonth()) y += 1;
    return new Date(y, mi, 1);
  }

  function parseFollowUp(raw, visitDateStr) {
    if (!raw || !raw.trim()) return { date: null, method: 'blank' };
    const visit = new Date(visitDateStr + 'T00:00:00');
    const s = raw.trim();
    const sl = s.toLowerCase().replace(/\.$/, '').trim();
    let m;

    // 1. Numeric date anywhere in the text: dd-mm-yyyy, dd.mm.yy, dd/mm/yyyy
    if ((m = s.match(/(\d{1,2})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]\s*(\d{2,4})/))) {
      const dd = parseInt(m[1], 10), mm = parseInt(m[2], 10);
      const yyyy = m[3].length === 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10);
      if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
        const d = new Date(yyyy, mm - 1, dd);
        if (!isNaN(d.getTime())) return { date: d, method: 'literal' };
      }
    }
    // 2. dd-Mon-yy, e.g. "8-Oct -26"
    if ((m = s.match(/^(\d{1,2})\s*-\s*([A-Za-z]{3,9})\s*-?\s*(\d{2,4})$/))) {
      const mi = monthIndex(m[2].toLowerCase());
      const yyyy = m[3].length === 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10);
      if (mi >= 0) return { date: new Date(yyyy, mi, parseInt(m[1], 10)), method: 'literal' };
    }

    if (/last\s*week/.test(sl)) return { date: null, method: 'unparseable' }; // contradictory as a *next* date - not guessed

    // 3. Week of a month: "October first week", "September 2 week",
    //    "Next Month First Week", "3rd week". Deadline = day 7/14/21/28.
    if ((m = sl.match(/\b(first|second|third|fourth|1st|2nd|3rd|4th|[1-4])\s*week\b/))) {
      const n = ORD_WEEK[m[1]];
      let start;
      const mon = sl.match(new RegExp('\\b' + MONTH_RE + '\\b'));
      if (mon) start = namedMonthStart(monthIndex(mon[1]), visit);
      else if (/next\s*month/.test(sl)) start = new Date(visit.getFullYear(), visit.getMonth() + 1, 1);
      else start = new Date(visit.getFullYear(), visit.getMonth(), 1);
      let d = new Date(start.getFullYear(), start.getMonth(), n * 7);
      if (!mon && d < visit) d = new Date(start.getFullYear(), start.getMonth() + 1, n * 7);
      return { date: d, method: 'relative' };
    }

    // 4. Date in words: "3rd September", "September 15", "15 sep"
    if ((m = sl.match(new RegExp('\\b(\\d{1,2})(?:st|nd|rd|th)?\\s*(?:of\\s*)?' + MONTH_RE + '\\b'))) ||
        (m = sl.match(new RegExp('\\b' + MONTH_RE + '\\s*(\\d{1,2})(?:st|nd|rd|th)?\\b')))) {
      const numFirst = /^\d/.test(m[0]);
      const dd = parseInt(numFirst ? m[1] : m[2], 10);
      const mi = monthIndex(numFirst ? m[2] : m[1]);
      if (mi >= 0 && dd >= 1 && dd <= 31) {
        const st = namedMonthStart(mi, visit);
        const d = new Date(st.getFullYear(), mi, dd);
        // "After 10th of September" names the earliest date, not a
        // deadline - give until the end of the following week, the same
        // allowance as the no-date rule.
        if (/\bafter\b/.test(sl.slice(0, sl.indexOf(m[0])))) return { date: addDays(sundayOf(d), 7), method: 'relative' };
        return { date: d, method: 'literal' };
      }
    }

    if (/\btoday|\btodays\b/.test(sl)) return { date: new Date(visit), method: 'relative' };
    if (/tomorrow/.test(sl)) return { date: addDays(visit, 1), method: 'relative' };

    // 5. Month end: "this month end" / "month end" = end of visit month;
    //    "next month end" = end of next month.
    if (/month\s*end|monthend/.test(sl)) {
      return { date: /next\s*month/.test(sl) ? endOfMonth(addMonths(new Date(visit.getFullYear(), visit.getMonth(), 1), 1)) : endOfMonth(visit), method: 'relative' };
    }

    // 6. Spelled-out durations stay exact: "in 2 days", "2 weeks after",
    //    "after 2 month", "next two months", "in one month".
    if ((m = sl.match(/\b(\d+|a|one|two|three|four|five)\s*days?\b/))) return { date: addDays(visit, toNum(m[1] === 'a' ? 'one' : m[1])), method: 'relative' };
    if ((m = sl.match(/\b(\d+|a|one|two|three|four|five)\s*weeks\b|\b(\d+|one|two|three|four|five)\s*week\b/))) {
      const tok = m[1] || m[2];
      return { date: addDays(visit, toNum(tok === 'a' ? 'one' : tok) * 7), method: 'relative' };
    }
    if ((m = sl.match(/\b(\d+|one|two|three|four|five)\s*months?\b/))) return { date: addMonths(visit, toNum(m[1])), method: 'relative' };

    // 7. Day names: next time that day comes round after the visit. With
    //    "next week" also written, that day in next week. "monday or
    //    tuesday" takes the later day, so the rep is not held to the
    //    earlier of two options they gave.
    const days = [];
    let dm;
    DAY_RE.lastIndex = 0;
    while ((dm = DAY_RE.exec(sl))) {
      const idx = DAY_NAMES.findIndex((n) => n.startsWith(dm[1].slice(0, 3)));
      if (idx >= 0) days.push(idx);
    }
    if (days.length) {
      const nextWeek = /next\s*we+k/.test(sl);
      let best = null;
      days.forEach((idx) => {
        let d;
        if (nextWeek) {
          const nextMon = addDays(mondayOf(visit), 7);
          d = addDays(nextMon, (idx + 6) % 7); // Mon=0 ... Sun=6 offset
        } else {
          const diff = ((idx - visit.getDay()) + 7) % 7 || 7;
          d = addDays(visit, diff);
        }
        if (!best || d > best) best = d;
      });
      return { date: best, method: 'relative' };
    }

    // 8. This week = by this Sunday; any other "week" (next week, coming
    //    week, next weekend) = by the Sunday of the following week.
    if (/\b(this|in this|within this|by this|current)\s*we+k/.test(sl)) return { date: sundayOf(visit), method: 'relative' };
    if (/we+kk?/.test(sl)) return { date: addDays(sundayOf(visit), 7), method: 'relative' };

    // 9. Next month = by the last day of next month.
    if (/next\s*months?/.test(sl)) return { date: endOfMonth(addMonths(new Date(visit.getFullYear(), visit.getMonth(), 1), 1)), method: 'relative' };

    return { date: null, method: 'unparseable' };
  }

  function followUpStatus(rec) {
    const parsed = parseFollowUp(rec.followUp, rec.date);
    if (parsed.method === 'blank') return { ...parsed, status: 'none' };
    if (!parsed.date) return { ...parsed, status: 'unparseable' };
    const today = new Date(); today.setHours(0, 0, 0, 0);
    return { ...parsed, status: parsed.date < today ? 'overdue' : 'ontrack', daysOverdue: Math.round((today - parsed.date) / 86400000) };
  }

  // ---- filters ---------------------------------------------------------
  function populateSelect(id, options, allLabel) {
    const el = document.getElementById(id);
    el.innerHTML = '';
    const allOpt = document.createElement('option');
    allOpt.value = ''; allOpt.textContent = allLabel;
    el.appendChild(allOpt);
    options.forEach((o) => {
      const opt = document.createElement('option');
      opt.value = o.value; opt.textContent = o.label;
      el.appendChild(opt);
    });
  }

  function refreshWeekOptions() {
    const month = document.getElementById('monthSelect').value;
    const weeks = DATA.weeks.filter((w) => !month || w.month === month);
    populateSelect('weekSelect', weeks.map((w) => ({ value: w.start, label: w.label })), 'All Weeks');
  }

  function currentFilters() {
    return {
      month: document.getElementById('monthSelect').value,
      week: document.getElementById('weekSelect').value,
      exec: document.getElementById('execSelect').value,
    };
  }

  function filteredRecords() {
    const f = currentFilters();
    return DATA.records.filter((r) => {
      // A row with no Checkin Time at all is a day nobody actually visited
      // anyone - not a submitted form. Excluded here (and everywhere this
      // function feeds), and shown instead in "No Check-ins Logged" below.
      if (!r.hasCheckinTime) return false;
      if (f.month && r.month !== f.month) return false;
      if (f.week && r.weekStart !== f.week) return false;
      if (f.exec && r.exec !== f.exec) return false;
      return true;
    });
  }

  function noCheckinRecords() {
    const f = currentFilters();
    return DATA.records.filter((r) => {
      if (r.hasCheckinTime) return false;
      if (f.month && r.month !== f.month) return false;
      if (f.week && r.weekStart !== f.week) return false;
      if (f.exec && r.exec !== f.exec) return false;
      return true;
    });
  }

  // ---- overview ----------------------------------------------------
  // Plain, non-interactive KPI tiles - the click-to-expand interaction
  // originally built here (22-Sep-2026) was based on a misreading of what
  // Harsh actually asked for; on 23-Sep-2026 he confirmed it should come off
  // the KPI cards and move onto the Timeline & Lead Type breakdown panels
  // instead (see renderBreakdown / renderBreakdownDetail below, which reuse
  // the same expand-panel mechanics).
  // Compact KPI tiles + a small ring chart showing the New/Existing split
  // at a glance, replacing the old 3-tile-plus-long-caption layout. The
  // exact reconciliation arithmetic that used to sit in view by default now
  // lives in #reconcilePanel, revealed by the "How this reconciles" button
  // (see the generic info-btn toggle near the bottom of this script) -
  // nothing was deleted, it's one click away instead of always-on.
  function renderOverview(records) {
    const forms = records.length;
    const leadsWithNotes = records.filter((r) => r.hasLead);
    const newLeads = leadsWithNotes.filter((r) => r.isNew).length;
    const existingLeads = leadsWithNotes.filter((r) => r.isExisting).length;
    const noLead = forms - leadsWithNotes.length;
    const total = leadsWithNotes.length;
    const newPct = total ? Math.round((newLeads / total) * 100) : 0;
    // Ring drawn with two overlapping circles at r=15.9 so the circumference
    // is ~100, letting the New-lead share be used directly as the dash length.
    const circumference = 2 * Math.PI * 15.9;
    const newDash = (newPct / 100) * circumference;

    const row = document.getElementById('overviewRow');
    row.innerHTML =
      '<div class="kpi-strip">' +
        '<div class="kpi-mini" data-kpi="forms">' +
          '<span class="icon-badge sm tone-teal" style="margin-bottom:6px;"><svg class="icon"><use href="#icon-inbox"></use></svg></span>' +
          '<div class="kv num">' + fmtNum(forms) + '</div><div class="kl">Forms submitted</div>' +
        '</div>' +
        '<div class="kpi-mini" data-kpi="new">' +
          '<span class="icon-badge sm tone-good" style="margin-bottom:6px;"><svg class="icon"><use href="#icon-users"></use></svg></span>' +
          '<div class="kv accent num">' + fmtNum(newLeads) + '</div><div class="kl">New leads</div>' +
        '</div>' +
        '<div class="kpi-mini" data-kpi="existing">' +
          '<span class="icon-badge sm tone-accent" style="margin-bottom:6px;"><svg class="icon"><use href="#icon-repeat"></use></svg></span>' +
          '<div class="kv num">' + fmtNum(existingLeads) + '</div><div class="kl">Existing leads</div>' +
        '</div>' +
      '</div>' +
      '<div class="donut-wrap">' +
        '<svg width="72" height="72" viewBox="0 0 42 42">' +
          '<circle cx="21" cy="21" r="15.9" fill="transparent" stroke="var(--surface-3)" stroke-width="6"></circle>' +
          '<circle cx="21" cy="21" r="15.9" fill="transparent" stroke="var(--good)" stroke-width="6" ' +
            'stroke-dasharray="' + newDash.toFixed(2) + ' ' + (circumference - newDash).toFixed(2) + '" stroke-dashoffset="' + (circumference / 4).toFixed(2) + '" transform="rotate(-90 21 21)"></circle>' +
          '<text x="21" y="24" text-anchor="middle" font-size="9" font-weight="700" fill="var(--ink)">' + newPct + '%</text>' +
        '</svg>' +
        '<div class="donut-legend">' +
          '<div><span class="dot" style="background:var(--good);"></span>' + fmtNum(newLeads) + ' new leads</div>' +
          '<div><span class="dot" style="background:var(--surface-3); border:1px solid var(--border);"></span>' + fmtNum(existingLeads) + ' existing leads</div>' +
        '</div>' +
      '</div>';

    document.getElementById('reconcilePanel').innerHTML =
      fmtNum(forms) + ' forms submitted = ' + fmtNum(leadsWithNotes.length) +
      ' with a lead logged (<strong>' + fmtNum(newLeads) + ' new</strong> + <strong>' + fmtNum(existingLeads) +
      ' existing</strong>) + ' + fmtNum(noLead) + ' check-in' + (noLead === 1 ? '' : 's') + ' with no meeting notes recorded. ' +
      'Days with no check-in time logged at all are excluded here and shown separately under &ldquo;Missed Days&rdquo; below.';
  }

  function rowValue(r, key) { return r[key] || ''; }

  function renderLeadsTableHead(kind) {
    const state = leadsTables[kind];
    const headRow = document.getElementById(kind + 'HeadRow');
    headRow.innerHTML = '';
    LEADS_COLUMNS.forEach((col) => {
      const th = document.createElement('th');
      th.textContent = col.label;
      if (col.key === state.sortKey) {
        const arrow = document.createElement('span');
        arrow.className = 'arrow';
        arrow.textContent = state.sortDir === 1 ? '▲' : '▼';
        th.appendChild(arrow);
      }
      th.addEventListener('click', () => {
        if (state.sortKey === col.key) { state.sortDir *= -1; } else { state.sortKey = col.key; state.sortDir = 1; }
        renderLeadsTable(kind, filteredRecords());
      });
      headRow.appendChild(th);
    });
  }

  function renderLeadsTableBody(kind, records) {
    const state = leadsTables[kind];
    const searchEl = document.getElementById(kind + 'Search');
    const search = searchEl ? searchEl.value.trim().toLowerCase() : '';
    let leads = records.filter((r) => r.hasLead && state.filterFn(r));

    if (search) {
      leads = leads.filter((r) =>
        (r.lead + ' ' + r.location + ' ' + r.remarks + ' ' + r.exec + ' ' + r.satisfaction).toLowerCase().includes(search)
      );
    }

    leads = leads.slice().sort((a, b) => {
      const av = rowValue(a, state.sortKey), bv = rowValue(b, state.sortKey);
      if (av < bv) return -1 * state.sortDir;
      if (av > bv) return 1 * state.sortDir;
      return 0;
    });

    const body = document.getElementById(kind + 'Body');
    body.innerHTML = '';
    document.getElementById(kind + 'RowCount').textContent = fmtNum(leads.length) + ' lead' + (leads.length === 1 ? '' : 's') + ' shown';

    if (leads.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = LEADS_COLUMNS.length;
      td.className = 'empty-state';
      td.textContent = 'No leads match this selection.';
      tr.appendChild(td);
      body.appendChild(tr);
      return;
    }

    leads.forEach((r) => {
      const tr = document.createElement('tr');
      LEADS_COLUMNS.forEach((col) => {
        const td = document.createElement('td');
        if (col.cls) td.className = col.cls;
        if (col.key === 'lead' && r.unnamedLead) {
          td.innerHTML = (r.lead || '(unnamed)') + ' <span class="tag unnamed">no name logged</span>';
        } else if (col.key === 'satisfaction' && r.satisfaction === 'Not Interested') {
          td.innerHTML = '<span class="tag notinterested">' + esc(r.satisfaction) + '</span>';
        } else {
          td.textContent = rowValue(r, col.key);
        }
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
  }

  function renderLeadsTable(kind, records) {
    renderLeadsTableHead(kind);
    renderLeadsTableBody(kind, records);
  }

  // ---- stale leads -------------------------------------------------
  // Every logged follow-up is bucketed into a specific week, so this
  // section respects the same Month/Week/Executive toggles as the rest of
  // the page: a dated follow-up ("10-09-2026", "Next week") is bucketed
  // into the week that date falls in; an undated one ("Follow Up",
  // "Dealer", or blank) has no real date to go on, so it is bucketed into
  // the week immediately AFTER the visit and clearly marked "assumed" -
  // never silently treated as due the same week it was logged. A lead
  // only shows here once its bucketed week has actually passed (today is
  // past that week's Sunday) with nothing logged since.
  // Total tracked visits per normalized lead name, across ALL records ever -
  // not filtered by the current Month/Week/Executive selection, same as the
  // New/Existing determination itself. Used to tell a lead the rep truly
  // never went back to (one visit, ever) from one that's been followed up
  // on before but has lapsed again (see renderStale below).
  function leadVisitCounts() {
    const counts = new Map();
    DATA.records.forEach((r) => {
      if (!r.hasLead || r.unnamedLead || !r.hasCheckinTime) return;
      const key = leadKey(r.lead);
      counts.set(key, (counts.get(key) || 0) + 1);
    });
    return counts;
  }

  // Check-in time of day in minutes, from the display value ("19 Sep, 03:16 PM").
  function minutesOf(r) {
    const m = /(\d{1,2}):(\d{2})\s*(AM|PM)/i.exec(r.time || '');
    if (!m) return -1;
    return ((parseInt(m[1], 10) % 12) + (/pm/i.test(m[3]) ? 12 : 0)) * 60 + parseInt(m[2], 10);
  }

  function computeStaleCandidates() {
    const f = currentFilters();
    const allNamed = DATA.records.filter((r) => r.hasLead && !r.unnamedLead && r.hasCheckinTime);
    // Joint-visit rule (24-Sep-2026, per Harsh): a lead is overdue for a rep
    // only if that rep was on the lead's most recent visit day. If any
    // colleague has visited since, it drops off this rep's list; if two reps
    // visited together and nobody has returned, it stays on both lists.
    const lastDayByLead = new Map();
    allNamed.forEach((r) => {
      const k = leadKey(r.lead), d = lastDayByLead.get(k);
      if (!d || r.date > d) lastDayByLead.set(k, r.date);
    });
    let records = allNamed;
    if (f.exec) records = records.filter((r) => r.exec === f.exec);

    const latestByLead = new Map();
    records.forEach((r) => {
      const key = r.lead.trim().toLowerCase();
      const prev = latestByLead.get(key);
      // Latest visit wins; on a same-day tie the later check-in time wins
      // (fixed 24-Sep-2026 - it used to keep the earlier same-day entry).
      if (!prev || r.date > prev.date || (r.date === prev.date && minutesOf(r) >= minutesOf(prev))) latestByLead.set(key, r);
    });

    const visitCounts = leadVisitCounts();
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const stale = [];
    latestByLead.forEach((r) => {
      if (r.date !== lastDayByLead.get(leadKey(r.lead))) return; // someone has visited since
      const parsed = parseFollowUp(r.followUp, r.date);
      let assumed = false, dueDate;
      if (parsed.date) {
        dueDate = parsed.date;
      } else {
        assumed = true;
        dueDate = addDays(mondayOf(new Date(r.date + 'T00:00:00')), 7); // Monday of the week AFTER the visit
      }
      const bucketMonday = mondayOf(dueDate);
      const assignedWeekStart = ymd(bucketMonday);
      const assignedMonth = monthLabelJS(bucketMonday);
      if (f.month && assignedMonth !== f.month) return;
      if (f.week && assignedWeekStart !== f.week) return;

      // An assumed (undated) follow-up is only "due" once its whole
      // bucketed week has passed - we only know it was due sometime that
      // week, not which exact day.
      const threshold = assumed ? addDays(bucketMonday, 6) : dueDate;
      if (threshold >= today) return;
      const days = Math.round((today - threshold) / 86400000);
      const key = leadKey(r.lead);
      // Cross-checked against all tracked history (the same data Section 2,
      // Existing Leads, draws on): has the rep ever been back to this lead
      // more than once, ever - regardless of the current filter?
      const everRevisited = (visitCounts.get(key) || 0) > 1;
      stale.push({ r, days, assumed, weekLabel: weekLabelJS(bucketMonday), key, everRevisited, due: threshold });
    });
    stale.sort((a, b) => b.days - a.days);
    return { stale, totalNamed: latestByLead.size };
  }

  // ---- stale leads ---------------------------------------------------
  // Flattened per Harsh's request (23-Sep-2026): no more Never Followed Up /
  // Lapsed Again distinction. Every lead whose promised follow-up has not
  // happened by the due date is shown in one plain list, sorted by days
  // overdue, worst first. everRevisited is still computed upstream (it
  // decides the "assumed due date" logic) but no longer surfaced as a tag.
  function renderStaleGroup(bodyId, rows) {
    const body = document.getElementById(bodyId);
    body.innerHTML = '';
    if (rows.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 12; td.className = 'empty-state';
      td.textContent = 'None in this selection.';
      tr.appendChild(td); body.appendChild(tr);
      return;
    }
    rows.forEach((s) => {
      const { r, days, assumed, weekLabel } = s;
      const tr = document.createElement('tr');
      const notInterested = r.satisfaction === 'Not Interested';
      const satCell = notInterested
        ? '<span class="tag notinterested">' + esc(r.satisfaction) + '</span>'
        : esc(r.satisfaction || '');
      // Standard seven-column core first (see LEADS_COLUMNS), then this
      // table's own metrics, then who / sentiment / notes. Time already
      // carries the visit date, so the separate Last Visited column was
      // dropped as a duplicate (23-Sep-2026, per Harsh). Next Follow-up is
      // the rep's own raw text - the thing Due Week is calculated from.
      tr.innerHTML =
        '<td class="lead-cell">' + esc(r.lead) + histButton(r.lead) + '</td>' +
        '<td class="time-cell">' + esc(r.time || '') + '</td>' +
        '<td>' + esc(r.value || '') + '</td>' +
        '<td>' + esc(r.leadType || '') + '</td>' +
        '<td class="loc-cell">' + esc(r.location || '') + '</td>' +
        '<td>' + esc(r.timeline || '') + '</td>' +
        '<td>' + esc(r.followUp || '(no date logged)') + '</td>' +
        '<td>' + weekLabel + (assumed ? ' <span class="tag assumed">assumed</span>' : '') + '</td>' +
        '<td><span class="tag overdue">' + days + ' day' + (days === 1 ? '' : 's') + '</span>' +
          (days > 365 ? ' <span style="font-size:12px;color:var(--ink-secondary);">(check the logged date - likely a typo, not corrected here)</span>' : '') + '</td>' +
        '<td>' + esc(r.exec) + '</td>' +
        '<td>' + satCell + '</td>' +
        '<td class="remarks-cell">' + esc(r.remarks || '') + '</td>';
      body.appendChild(tr);
      attachHistoryToggle(tr, r.lead, new Set([r]), 'hl-over', 'Red: the visit whose follow-up is overdue');
    });
  }

  // ---- followed up on time (added 23-Sep-2026) --------------------------
  // The mirror image of the Overdue Follow-ups list below: did the rep come
  // back on or before the date they themselves promised on the PREVIOUS
  // visit to this lead. Uses the exact same due-date + assumed-week logic as
  // computeStaleCandidates (see parseFollowUp above) so a lead is never
  // judged "on time" by a looser standard than it would be judged "overdue"
  // elsewhere on this page. Every consecutive visit pair is checked, not
  // just the latest, credited to the executive who made the promise, and
  // bucketed by the week the promise was actually kept.
  function computeFollowedUpOnTime() {
    const f = currentFilters();
    const records = DATA.records.filter((r) => r.hasLead && !r.unnamedLead && r.hasCheckinTime);
    const byLead = new Map();
    records.forEach((r) => {
      const key = leadKey(r.lead);
      if (!byLead.has(key)) byLead.set(key, []);
      byLead.get(key).push(r);
    });
    byLead.forEach((arr) => arr.sort((a, b) => (a.date < b.date ? -1 : 1)));

    const kept = [];
    byLead.forEach((visits) => {
      for (let i = 0; i < visits.length - 1; i++) {
        const prev = visits[i], next = visits[i + 1];
        const parsed = parseFollowUp(prev.followUp, prev.date);
        let assumed = false, dueDate;
        if (parsed.date) {
          dueDate = parsed.date;
        } else {
          assumed = true;
          dueDate = addDays(mondayOf(new Date(prev.date + 'T00:00:00')), 7);
        }
        const threshold = assumed ? addDays(mondayOf(dueDate), 6) : dueDate;
        const nextDate = new Date(next.date + 'T00:00:00');
        if (nextDate > threshold) continue; // came back late - not this section
        const keptMonday = mondayOf(nextDate);
        const assignedWeekStart = ymd(keptMonday);
        const assignedMonth = monthLabelJS(keptMonday);
        if (f.month && assignedMonth !== f.month) continue;
        if (f.week && assignedWeekStart !== f.week) continue;
        if (f.exec && prev.exec !== f.exec) continue;
        kept.push({ prev, next, assumed });
      }
    });
    kept.sort((a, b) => (a.next.date < b.next.date ? 1 : -1));
    return kept;
  }

  // ---- visit history drop-down (Section 4, added 24-Sep-2026) -----------
  // Every logged visit to a lead, oldest first, all reps and all weeks -
  // deliberately NOT narrowed by the Month / Week / Executive filters, so HR
  // sees the whole relationship. Visit order uses the same comparator as
  // computeFollowedUpOnTime and each promise's deadline uses the same rule,
  // so a promise shown as "Kept" here is exactly one counted in Followed Up
  // On Time, and the last visit shows "Overdue" exactly when the lead is in
  // Overdue Follow-ups.
  const HISTORY_COLS = ['#', 'Time', 'Est. Project Value', 'Type of Lead', 'Location', 'Requirement Timeline',
    'Next Follow-up', 'Due By', 'Outcome', 'Sales Executive', 'Client Satisfaction', 'Remarks'];

  function leadHistory(name) {
    const key = leadKey(name);
    return DATA.records
      .filter((r) => r.hasLead && !r.unnamedLead && r.hasCheckinTime && leadKey(r.lead) === key)
      .sort((a, b) => (a.date < b.date ? -1 : 1));
  }

  function promiseDeadline(rec) {
    const p = parseFollowUp(rec.followUp, rec.date);
    if (p.date) return { due: p.date, assumed: false };
    // No usable date: until the Sunday of the week after the visit.
    return { due: addDays(mondayOf(new Date(rec.date + 'T00:00:00')), 13), assumed: true };
  }

  function fmtDay(d) { return d.getDate() + ' ' + d.toLocaleString('en-US', { month: 'short' }) + ' ' + d.getFullYear(); }
  function plural(n, w) { return n + ' ' + w + (n === 1 ? '' : 's'); }

  function historyHtml(name, hlSet, hlClass, legend) {
    const visits = leadHistory(name);
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const reps = new Set(visits.map((v) => v.exec));
    let rows = '';
    visits.forEach((v, i) => {
      const { due, assumed } = promiseDeadline(v);
      const next = visits[i + 1];
      let outcome;
      if (next) {
        const nd = new Date(next.date + 'T00:00:00');
        if (nd <= due) {
          outcome = next.date === v.date
            ? '<span class="tag sameday" title="The next entry is on the same day. Counted as a kept follow-up.">Kept, same day</span>'
            : '<span class="tag kept">Kept</span>';
        } else {
          outcome = '<span class="tag late">Late by ' + plural(Math.round((nd - due) / 86400000), 'day') + '</span>';
        }
      } else if (due < today) {
        outcome = '<span class="tag overdue">Overdue ' + plural(Math.round((today - due) / 86400000), 'day') + '</span>';
      } else {
        outcome = '<span class="tag pending">Not yet due</span>';
      }
      const sat = v.satisfaction === 'Not Interested'
        ? '<span class="tag notinterested">' + esc(v.satisfaction) + '</span>' : esc(v.satisfaction || '');
      rows += '<tr' + (hlSet.has(v) ? ' class="' + hlClass + '"' : '') + '>' +
        '<td class="visit-no">' + (i + 1) + '</td>' +
        '<td class="time-cell">' + esc(v.time || '') + '</td>' +
        '<td>' + esc(v.value || '') + '</td>' +
        '<td>' + esc(v.leadType || '') + '</td>' +
        '<td class="loc-cell">' + esc(v.location || '') + '</td>' +
        '<td>' + esc(v.timeline || '') + '</td>' +
        '<td>' + esc(v.followUp || '(no date logged)') + '</td>' +
        '<td class="time-cell">' + fmtDay(due) + (assumed ? ' <span class="tag assumed">assumed</span>' : '') + '</td>' +
        '<td>' + outcome + '</td>' +
        '<td>' + esc(v.exec) + '</td>' +
        '<td>' + sat + '</td>' +
        '<td class="remarks-cell">' + esc(v.remarks || '') + '</td></tr>';
    });
    const swatch = hlClass === 'hl' ? 'var(--good-bg)' : 'var(--critical-bg)';
    return '<div class="hist-title"><strong>All ' + visits.length + ' visits to this lead</strong>, ' +
      (reps.size > 1 ? plural(reps.size, 'rep') + ', ' : '') + 'every week, not narrowed by the filters above.' +
      '<span class="hist-swatch" style="background:' + swatch + '"></span>' + legend + '. Outcome is whether the next visit came by the Due By date.</div>' +
      '<div class="hist-scroll"><table class="hist-table"><thead><tr>' +
      HISTORY_COLS.map((c) => '<th>' + c + '</th>').join('') + '</tr></thead><tbody>' + rows + '</tbody></table></div>';
  }

  function histButton(name) {
    const n = leadHistory(name).length;
    return n > 1 ? '<br><button type="button" class="hist-btn" aria-expanded="false"><span class="chev">&#9656;</span> ' + n + ' visits</button>' : '';
  }

  function attachHistoryToggle(tr, name, hlSet, hlClass, legend) {
    const btn = tr.querySelector('.hist-btn');
    if (!btn) return;
    btn.addEventListener('click', () => {
      const nxt = tr.nextElementSibling;
      if (nxt && nxt.classList.contains('history-row')) {
        nxt.remove(); btn.setAttribute('aria-expanded', 'false'); return;
      }
      const hr = document.createElement('tr');
      hr.className = 'history-row';
      const td = document.createElement('td');
      td.colSpan = tr.children.length;
      td.innerHTML = historyHtml(name, hlSet, hlClass, legend);
      hr.appendChild(td); tr.after(hr);
      btn.setAttribute('aria-expanded', 'true');
    });
  }

  function renderFollowedUpOnTime() {
    const kept = computeFollowedUpOnTime();
    document.getElementById('keptCaption').textContent =
      'Promised a follow-up on an earlier visit, then actually came back on or before that date - or before that week ended, when no literal date was logged. Click "visits" under a lead to see every visit to it.';
    const body = document.getElementById('keptBody');
    body.innerHTML = '';
    if (kept.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 10; td.className = 'empty-state';
      td.textContent = 'None in this selection.';
      tr.appendChild(td); body.appendChild(tr);
    } else {
      kept.forEach((k) => {
        const tr = document.createElement('tr');
        const notInterested = k.next.satisfaction === 'Not Interested';
        const satCell = notInterested
          ? '<span class="tag notinterested">' + esc(k.next.satisfaction) + '</span>'
          : esc(k.next.satisfaction || '');
        // Standard seven-column core first (see LEADS_COLUMNS). Time is the
        // visit that kept the promise (it already carries the date, so the
        // separate Kept On column was dropped as a duplicate, 23-Sep-2026).
        // Next Follow-up is the promise the rep made on the EARLIER visit -
        // the one this row proves was kept.
        tr.innerHTML =
          '<td class="lead-cell">' + esc(k.prev.lead) + histButton(k.prev.lead) + '</td>' +
          '<td class="time-cell">' + esc(k.next.time || '') + '</td>' +
          '<td>' + esc(k.next.value || '') + '</td>' +
          '<td>' + esc(k.next.leadType || '') + '</td>' +
          '<td class="loc-cell">' + esc(k.next.location || '') + '</td>' +
          '<td>' + esc(k.next.timeline || '') + '</td>' +
          '<td>' + esc(k.prev.followUp || '(no date logged)') + (k.assumed ? ' <span class="tag assumed">assumed</span>' : '') + '</td>' +
          '<td>' + esc(k.prev.exec) + '</td>' +
          '<td>' + satCell + '</td>' +
          '<td class="remarks-cell">' + esc(k.next.remarks || '') + '</td>';
        body.appendChild(tr);
        attachHistoryToggle(tr, k.prev.lead, new Set([k.prev, k.next]), 'hl', 'Green: the promise and the visit that kept it');
      });
    }
    document.getElementById('keptNote').textContent =
      fmtNum(kept.length) + ' follow-up' + (kept.length === 1 ? '' : 's') + ' kept on time in this selection.';
    return kept.length;
  }

  // ---- Overdue follow-ups PDF (added 24-Sep-2026) -------------------------
  // One click downloads the selected rep's Overdue Follow-ups, exactly the
  // rows on screen for the current Month / Week selection, as a PDF HR can
  // send to the rep. Uses the jsPDF + AutoTable libraries embedded in this
  // file, so it works offline. Standard PDF fonts have no rupee sign or
  // emoji, so the PDF prints "Rs" for the rupee sign and drops emoji.
  function pdfText(s) {
    return String(s || '').replace(/₹\s*/g, 'Rs ')
      .replace(/[☀-➿]|[\uD800-\uDBFF][\uDC00-\uDFFF]|️/g, '')
      .replace(/\s+/g, ' ').trim();
  }
  function overduePeriod(f) {
    if (f.week) {
      const w = DATA.weeks.find((x) => x.start === f.week);
      const label = w ? w.label : f.week;
      return { text: 'Follow-ups due in the week of ' + label, slug: label.replace(/\s+/g, '') };
    }
    if (f.month) return { text: 'Follow-ups due in ' + f.month, slug: f.month.replace(/\s+/g, '') };
    const dates = DATA.records.map((r) => r.date).sort();
    const d0 = new Date(dates[0] + 'T00:00:00'), d1 = new Date(dates[dates.length - 1] + 'T00:00:00');
    return { text: 'All weeks tracked (visits ' + fmtDay(d0) + ' to ' + fmtDay(d1) + ')', slug: 'AllWeeks' };
  }
  function updatePdfButton(n, f) {
    const btn = document.getElementById('overduePdfBtn'), label = document.getElementById('overduePdfLabel');
    if (!btn) return;
    if (!f.exec) { btn.disabled = true; label.textContent = 'Pick a sales executive to download'; return; }
    if (!n) { btn.disabled = true; label.textContent = 'Nothing overdue for ' + f.exec; return; }
    btn.disabled = false;
    label.textContent = 'Download PDF for ' + f.exec + ' (' + n + ')';
  }
  function exportOverduePdf() {
    const f = currentFilters();
    if (!f.exec || !window.jspdf) return;
    const { stale } = computeStaleCandidates();
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const period = overduePeriod(f);
    const doc = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    const W = doc.internal.pageSize.getWidth(), M = 28;
    doc.setFillColor(22, 22, 22); doc.rect(0, 0, W, 6, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.setTextColor(22, 22, 22);
    doc.text('Overdue follow-ups: ' + pdfText(f.exec), M, 34);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(90, 90, 90);
    doc.text(period.text + '   |   As of ' + fmtDay(today) + '   |   ' + plural(stale.length, 'lead'), M, 51);
    doc.setFontSize(9); doc.setTextColor(60, 60, 60);
    doc.text(doc.splitTextToSize('Each lead below had a follow-up promised on your last visit. That date has passed and no visit has been logged since. ' +
      'Please follow up and log each visit in LystLoc. Where no date was written, the deadline is taken as the end of the week after the visit (marked "assumed").', W - 2 * M), M, 67);
    const body = stale.map((s, i) => [
      String(i + 1), pdfText(s.r.lead), pdfText(s.r.time), pdfText(s.r.value), pdfText(s.r.leadType), pdfText(s.r.location),
      pdfText(s.r.timeline), pdfText(s.r.followUp || '(no date logged)'),
      fmtDay(s.due) + (s.assumed ? ' (assumed)' : ''),
      String(s.days) + (s.days > 365 ? ' (check date entered)' : ''),
      pdfText(s.r.satisfaction), pdfText(s.r.remarks),
    ]);
    doc.autoTable({
      startY: 92, margin: { left: M, right: M, bottom: 30 },
      head: [['#', 'Lead', 'Last Visit', 'Est. Project Value', 'Type of Lead', 'Location', 'Requirement Timeline',
              'Next Follow-up (as written)', 'Due By', 'Days Overdue', 'Client Satisfaction', 'Remarks']],
      body,
      styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 4, textColor: [30, 30, 30], valign: 'top', overflow: 'linebreak', lineColor: [225, 225, 225], lineWidth: 0.5 },
      headStyles: { fillColor: [22, 22, 22], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 7.5 },
      alternateRowStyles: { fillColor: [247, 245, 242] },
      columnStyles: { 0: { cellWidth: 16, textColor: [140, 140, 140] }, 1: { cellWidth: 82, fontStyle: 'bold' }, 2: { cellWidth: 74 }, 3: { cellWidth: 62 },
                      4: { cellWidth: 62 }, 5: { cellWidth: 96 }, 6: { cellWidth: 58 }, 7: { cellWidth: 70 }, 8: { cellWidth: 54 },
                      9: { cellWidth: 38, halign: 'center', textColor: [178, 58, 46], fontStyle: 'bold' }, 10: { cellWidth: 58 }, 11: { cellWidth: 'auto' } },
      didDrawPage: () => {
        doc.setFontSize(7.5); doc.setTextColor(140, 140, 140); doc.setFont('helvetica', 'normal');
        doc.text('Pasolite  |  Field Visit & Lead Dashboard  |  Overdue follow-ups for ' + pdfText(f.exec), M, doc.internal.pageSize.getHeight() - 14);
        doc.text('Page ' + doc.internal.getNumberOfPages() + ' of {total}', W - M, doc.internal.pageSize.getHeight() - 14, { align: 'right' });
      },
    });
    if (typeof doc.putTotalPages === 'function') doc.putTotalPages('{total}');
    const asof = fmtDay(today).replace(/\s+/g, '');
    doc.save('Overdue_Followups_' + f.exec.replace(/\s+/g, '_') + '_' + period.slug + '_asof_' + asof + '.pdf');
  }

  // ---- Followed up by a colleague (rep view only, added 24-Sep-2026) ------
  // Leads where the selected rep's last visit was followed, on a later day,
  // by a visit from another rep. These are exactly the leads the joint-visit
  // rule leaves out of the rep's Overdue list (see computeStaleCandidates),
  // shown here instead so HR can see who took the lead over. Filtered by the
  // week the rep's follow-up was due, the same way as the Overdue list.
  function computeColleagueFollowUps() {
    const f = currentFilters();
    if (!f.exec) return [];
    const byLead = new Map();
    DATA.records.filter((r) => r.hasLead && !r.unnamedLead && r.hasCheckinTime).forEach((r) => {
      const k = leadKey(r.lead);
      if (!byLead.has(k)) byLead.set(k, []);
      byLead.get(k).push(r);
    });
    const out = [];
    byLead.forEach((list) => {
      let last = null;
      list.forEach((r) => {
        if (r.exec !== f.exec) return;
        if (!last || r.date > last.date || (r.date === last.date && minutesOf(r) >= minutesOf(last))) last = r;
      });
      if (!last) return;
      const later = list.filter((r) => r.exec !== f.exec && r.date > last.date)
        .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : minutesOf(a) - minutesOf(b)));
      if (!later.length) return;
      const { due, assumed } = promiseDeadline(last);
      const bucketMonday = mondayOf(due);
      if (f.month && monthLabelJS(bucketMonday) !== f.month) return;
      if (f.week && ymd(bucketMonday) !== f.week) return;
      const first = later[0];
      out.push({ r: last, due, assumed, first, more: later.length - 1,
                 reps: [...new Set(later.map((x) => x.exec))], beforeDue: new Date(first.date + 'T00:00:00') <= due });
    });
    out.sort((a, b) => (a.first.date < b.first.date ? 1 : -1));
    return out;
  }

  function renderColleagueFollowUps() {
    const f = currentFilters();
    const block = document.getElementById('colleagueBlock');
    if (!block) return;
    if (!f.exec) { block.style.display = 'none'; return; }
    block.style.display = '';
    const rows = computeColleagueFollowUps();
    document.getElementById('colleagueCaption').textContent =
      'Leads where ' + f.exec + ' promised a follow-up and another rep has visited since. They are left out of ' + f.exec + "'s overdue list above.";
    const body = document.getElementById('colleagueBody');
    body.innerHTML = '';
    if (!rows.length) {
      const tr = document.createElement('tr'), td = document.createElement('td');
      td.colSpan = 13; td.className = 'empty-state'; td.textContent = 'None for ' + f.exec + ' in this selection.';
      tr.appendChild(td); body.appendChild(tr);
    }
    rows.forEach((c) => {
      const r = c.r, tr = document.createElement('tr');
      const sat = r.satisfaction === 'Not Interested'
        ? '<span class="tag notinterested">' + esc(r.satisfaction) + '</span>' : esc(r.satisfaction || '');
      tr.innerHTML =
        '<td class="lead-cell">' + esc(r.lead) + histButton(r.lead) + '</td>' +
        '<td class="time-cell">' + esc(r.time || '') + '</td>' +
        '<td>' + esc(r.value || '') + '</td>' +
        '<td>' + esc(r.leadType || '') + '</td>' +
        '<td class="loc-cell">' + esc(r.location || '') + '</td>' +
        '<td>' + esc(r.timeline || '') + '</td>' +
        '<td>' + esc(r.followUp || '(no date logged)') + '</td>' +
        '<td class="time-cell">' + fmtDay(c.due) + (c.assumed ? ' <span class="tag assumed">assumed</span>' : '') + '</td>' +
        '<td>' + esc(c.reps.join(', ')) + '</td>' +
        '<td class="time-cell">' + esc(c.first.time || '') + (c.more ? ' <span class="tag pending">+' + c.more + ' more</span>' : '') + '</td>' +
        '<td>' + (c.beforeDue ? '<span class="tag kept">Before due date</span>' : '<span class="tag late">After due date</span>') + '</td>' +
        '<td>' + sat + '</td>' +
        '<td class="remarks-cell">' + esc(r.remarks || '') + '</td>';
      body.appendChild(tr);
      attachHistoryToggle(tr, r.lead, new Set([r, c.first]), 'hl', "Green: " + f.exec + "'s last visit and the colleague's first visit after it");
    });
    document.getElementById('colleagueNote').textContent =
      fmtNum(rows.length) + ' lead' + (rows.length === 1 ? '' : 's') + ' followed up by a colleague in this selection.';
  }

  function renderStale() {
    const f = currentFilters();
    const { stale, totalNamed } = computeStaleCandidates();
    updatePdfButton(stale.length, f);
    renderColleagueFollowUps();

    const keptCount = renderFollowedUpOnTime();
    document.getElementById('stalePillRow').innerHTML =
      '<div class="pill">' +
        '<span class="icon-badge tone-good"><svg class="icon"><use href="#icon-check-circle"></use></svg></span>' +
        '<div><div class="pv num" style="color:var(--good);">' + fmtNum(keptCount) + '</div><div class="pl">Followed up on time<br>(promise kept)</div></div>' +
      '</div>' +
      '<div class="pill">' +
        '<span class="icon-badge tone-accent"><svg class="icon"><use href="#icon-alert-triangle"></use></svg></span>' +
        '<div><div class="pv num">' + fmtNum(stale.length) + '</div><div class="pl">Overdue for follow-up<br>(promise not kept, as of today)</div></div>' +
      '</div>';

    document.getElementById('staleInfoPanel').innerHTML =
      fmtNum(totalNamed) + ' named lead' + (totalNamed === 1 ? '' : 's') + ' tracked overall' + (f.exec ? ' for ' + f.exec : '') +
      ' &middot; ' + fmtNum(stale.length) + ' overdue for follow-up in the current Month / Week / Executive selection, cross-checked against all tracked history. ' +
      'How the Next Follow-up entry is read: a date (&ldquo;10-09-2026&rdquo;, &ldquo;3rd September&rdquo;) is due that day; a named day (&ldquo;Friday&rdquo;) is due the next time that day comes round; &ldquo;this week&rdquo; is due by this Sunday; &ldquo;next week&rdquo; by the end of next week; &ldquo;next month&rdquo; by the end of next month; &ldquo;October first week&rdquo; by 7 October; a spelled-out gap (&ldquo;in 2 days&rdquo;) is counted exactly. ' +
      'An entry with no usable date (e.g. &ldquo;Follow Up&rdquo;, &ldquo;New&rdquo;) is given until the end of the week after the visit and marked &ldquo;assumed&rdquo;.';

    // Flattened 23-Sep-2026 per Harsh's request: dropped the Never Followed Up /
    // Lapsed Again distinction entirely (both the Status column and the two
    // separate count pills). One plain "Overdue Follow-ups" list, sorted by
    // Days Overdue, worst first - everRevisited is still computed upstream
    // for the assumed-due-date logic, it's just no longer shown as a tag.
    document.getElementById('staleOverdueCaption').textContent =
      'Sorted by days overdue, worst first. Click "visits" under a lead to see its earlier visits.';
    renderStaleGroup('staleOverdueBody', stale);
    document.getElementById('staleOverdueNote').textContent =
      fmtNum(stale.length) + ' lead' + (stale.length === 1 ? '' : 's') + ' overdue for follow-up' + (f.exec ? ' for ' + f.exec : '') + '.';
  }

  // ---- timeline / lead type breakdown --------------------------------
  // Client Satisfaction panel removed 22-Sep-2026 per Harsh's request. The
  // underlying "satisfaction" field is still parsed and present on every
  // record (Task D still captures it) - only the display here was dropped.
  //
  // Click-to-filter (added 23-Sep-2026, moved here from the KPI cards per
  // Harsh's correction; extended the same day to combined filtering across
  // all three Lead Mix groupings - see activeFilters below). Clicking a
  // value row or pie slice - e.g. "Immediate Requirement" or "Architect" -
  // toggles it on. Multiple values can be active within one grouping (OR)
  // and across all three groupings at once (AND), and the shared detail
  // panel below always shows the exact leads matching the full combination.
  // Respects the current Month / Week / Executive selection throughout.
  const TIMELINE_ORDER = ['Immediate Requirement', 'Within 15 Days', '15 Days to 1 Months', '1 to 2 Months', 'Above 2 Months'];
  const VALUE_ORDER = ['Below ₹50,000', '₹50,000 to ₹2 Lakhs', '₹2 Lakhs to ₹5 Lakhs', '₹5 Lakhs to ₹10 Lakhs', '₹10 Lakhs to ₹15 Lakhs', 'Above 15 Lakhs'];
  const BREAKDOWN_FIELD_LABELS = { timeline: 'Project Requirement Timeline', leadType: 'Type of Lead', value: 'Estimated Project Value' };
  // Icon + tone per category value, added 23-Sep-2026 for the graphical
  // redesign. Anything not listed here (a lead-type/timeline value the data
  // hasn't shown yet) falls back to a neutral icon rather than breaking.
  const TIMELINE_ICONS = {
    'Immediate Requirement': { icon: 'icon-alert-triangle', tone: 'tone-accent' },
    'Within 15 Days': { icon: 'icon-clock', tone: 'tone-amber' },
    '15 Days to 1 Months': { icon: 'icon-clock', tone: 'tone-watch' },
    '1 to 2 Months': { icon: 'icon-clock', tone: 'tone-teal' },
    'Above 2 Months': { icon: 'icon-clock', tone: 'tone-good' },
  };
  const LEADTYPE_ICONS = {
    'Architect': { icon: 'icon-arch', tone: 'tone-accent' },
    'Project': { icon: 'icon-briefcase', tone: 'tone-teal' },
    'Builders or Procurement Manager': { icon: 'icon-hardhat', tone: 'tone-amber' },
    'Dealer': { icon: 'icon-store', tone: 'tone-good' },
    'Lighting consultant': { icon: 'icon-bulb', tone: 'tone-watch' },
    'Distributor': { icon: 'icon-truck', tone: 'tone-neutral' },
  };
  const FALLBACK_ICON = { icon: 'icon-bar-chart', tone: 'tone-watch' };
  // Display-only shorthands so the fixed-width bar label doesn't truncate
  // mid-word - filtering, sorting and the detail panel still use the raw
  // value untouched.
  const TIMELINE_SHORT_LABELS = {
    'Immediate Requirement': 'Immediate',
    'Within 15 Days': 'Within 15 days',
    '15 Days to 1 Months': '15 days – 1 month',
    '1 to 2 Months': '1 – 2 months',
    'Above 2 Months': '2+ months',
  };
  const LEADTYPE_SHORT_LABELS = {
    'Builders or Procurement Manager': 'Builder / procurement',
  };
  // Sequential colour ramp for the value pie, low to high. The top tier was
  // deliberately made green rather than the dashboard's usual red/accent -
  // Harsh's call (23-Sep-2026): red is used elsewhere here for overdue/
  // alert states, and the highest-value tier is the opposite of a problem.
  const VALUE_COLORS = {
    'Below ₹50,000': '#8A8A8A',
    '₹50,000 to ₹2 Lakhs': '#2E7D8C',
    '₹2 Lakhs to ₹5 Lakhs': '#B4791E',
    '₹5 Lakhs to ₹10 Lakhs': '#2F8F5B',
    '₹10 Lakhs to ₹15 Lakhs': '#B71C1C',
    'Above 15 Lakhs': '#1B7A46',
  };
  const VALUE_SHORT_LABELS = {
    'Below ₹50,000': 'Below ₹50k',
    '₹50,000 to ₹2 Lakhs': '₹50k – 2L',
    '₹2 Lakhs to ₹5 Lakhs': '₹2 – 5L',
    '₹5 Lakhs to ₹10 Lakhs': '₹5 – 10L',
    '₹10 Lakhs to ₹15 Lakhs': '₹10 – 15L',
    'Above 15 Lakhs': 'Above ₹15L',
  };
  // Combined filtering across Lead Mix (added 23-Sep-2026 per Harsh's
  // request, replacing the single expandedBreakdown/single-select model).
  // activeFilters holds a Set of raw values per category. Within one
  // category, multiple selected values are OR'd (e.g. two timeline bands
  // together); across the three categories the selections are AND'd. Each
  // panel's own counts are computed against the OTHER two categories' active
  // filters only (leadMatchesOthers), so selecting within a panel narrows
  // the other two panels and the pie, and you can still see - and add to -
  // every option inside the panel you're looking at. The shared detail
  // table below always reflects the full intersection (leadMatchesAll).
  const FILTER_FIELDS = ['timeline', 'leadType', 'value'];
  let activeFilters = { timeline: new Set(), leadType: new Set(), value: new Set() };

  function activeFilterCount() {
    return FILTER_FIELDS.reduce((n, f) => n + activeFilters[f].size, 0);
  }

  function leadMatchesOthers(lead, excludeField) {
    return FILTER_FIELDS.every((f) => {
      if (f === excludeField) return true;
      const set = activeFilters[f];
      if (!set.size) return true;
      return set.has((lead[f] || '').trim());
    });
  }

  function leadMatchesAll(lead) {
    return FILTER_FIELDS.every((f) => {
      const set = activeFilters[f];
      if (!set.size) return true;
      return set.has((lead[f] || '').trim());
    });
  }

  function toggleFilter(field, value) {
    const set = activeFilters[field];
    if (set.has(value)) set.delete(value); else set.add(value);
    renderBreakdown(filteredRecords());
  }

  function orderedCounts(records, field, preferredOrder) {
    const counts = {};
    records.forEach((r) => { const v = (r[field] || '').trim(); if (v) counts[v] = (counts[v] || 0) + 1; });
    const keys = Object.keys(counts);
    const ordered = preferredOrder.filter((k) => counts[k] !== undefined)
      .concat(keys.filter((k) => !preferredOrder.includes(k)).sort((a, b) => counts[b] - counts[a]));
    return ordered.map((k) => ({ label: k, count: counts[k] }));
  }

  function renderBreakdownPanel(field, title, rows, total, iconMap, shortLabels) {
    let html = '<div class="breakdown-panel"><h3>' + title + '</h3>';
    if (rows.length === 0) {
      html += '<p style="font-size:12.5px;color:var(--ink-secondary);margin:0;">No data in this selection.</p>';
    } else {
      const max = Math.max(...rows.map((r) => r.count));
      rows.forEach((r) => {
        const w = max ? Math.round((r.count / max) * 100) : 0;
        const active = activeFilters[field].has(r.label);
        const meta = iconMap[r.label] || FALLBACK_ICON;
        const displayLabel = (shortLabels && shortLabels[r.label]) || r.label;
        html += '<div class="bar-row' + (active ? ' bar-row-active' : '') + '" data-field="' + esc(field) + '" data-value="' + esc(r.label) + '" tabindex="0" role="button" aria-expanded="' + (active ? 'true' : 'false') + '" title="' + esc(r.label) + '">' +
          '<span class="bar-icon icon-badge ' + meta.tone + '"><svg class="icon"><use href="#' + meta.icon + '"></use></svg></span>' +
          '<span class="bar-label">' + esc(displayLabel) + '</span>' +
          '<span class="bar-track"><span class="bar-fill" style="width:' + w + '%;"></span></span>' +
          '<span class="bar-count">' + fmtNum(r.count) + ' · ' + pct(r.count, total) + '%</span>' +
        '</div>';
      });
    }
    html += '</div>';
    return html;
  }

  // ---- Estimated Project Value pie (Section 1, "What it's worth") -------
  // Bigger chart with leader-line labels placed directly on the chart
  // (rather than a side legend), and wired into the same shared detail
  // panel as the two bar breakdowns above. Recomputed on every render from
  // the current filtered `leads` array, same as the bar panels - not a
  // static snapshot, so it always reflects the live Month/Week/Executive
  // selection. A generic angular collision-avoidance pass keeps adjacent
  // thin slices' labels from overlapping regardless of how the underlying
  // numbers shift as new weeks are added.
  function pt(cx, cy, radius, deg) {
    const rad = deg * Math.PI / 180;
    return [cx + radius * Math.cos(rad), cy + radius * Math.sin(rad)];
  }

  function renderPie(leads) {
    const svgEl = document.getElementById('pieSvg');
    const total = leads.length;
    const counts = {};
    leads.forEach((r) => { const v = (r.value || '').trim(); if (v) counts[v] = (counts[v] || 0) + 1; });
    const keys = VALUE_ORDER.filter((k) => counts[k] !== undefined)
      .concat(Object.keys(counts).filter((k) => !VALUE_ORDER.includes(k)));
    const denom = keys.reduce((s, k) => s + counts[k], 0);

    if (!denom) {
      svgEl.innerHTML = '';
      svgEl.setAttribute('viewBox', '0 0 400 60');
      svgEl.insertAdjacentHTML('beforeend', '<text x="200" y="34" text-anchor="middle" font-size="12" fill="var(--ink-muted)" font-family="Open Sans, sans-serif">No estimated project value logged in this selection.</text>');
      return;
    }
    svgEl.setAttribute('viewBox', '-40 0 400 260');

    const cx = 150, cy = 128, r = 86;
    let angle = -90;
    const slices = keys.map((k) => {
      const n = counts[k];
      const frac = n / denom;
      const sweep = frac * 360;
      const s = {
        key: k, label: VALUE_SHORT_LABELS[k] || k, color: VALUE_COLORS[k] || 'var(--ink-muted)',
        n: n, pctVal: Math.round(100 * frac), a0: angle, a1: angle + sweep, mid: angle + sweep / 2,
      };
      angle += sweep;
      return s;
    });

    let prevLabel = null;
    const minGap = 30;
    slices.forEach((s) => {
      let desired = s.mid;
      if (prevLabel !== null && (desired - prevLabel) < minGap) desired = prevLabel + minGap;
      s.labelAngle = desired;
      prevLabel = desired;
    });

    let svg = '';
    slices.forEach((s) => {
      const [x0, y0] = pt(cx, cy, r, s.a0);
      const [x1, y1] = pt(cx, cy, r, s.a1);
      const large = (s.a1 - s.a0) > 180 ? 1 : 0;
      const active = activeFilters.value.has(s.key);
      const dim = activeFilters.value.size > 0 && !active;
      const d = 'M' + cx + ',' + cy + ' L' + x0.toFixed(2) + ',' + y0.toFixed(2) + ' A' + r + ',' + r + ' 0 ' + large + ' 1 ' + x1.toFixed(2) + ',' + y1.toFixed(2) + ' Z';
      svg += '<path class="pie-slice' + (active ? ' active' : '') + (dim ? ' dim' : '') + '" data-field="value" data-value="' + esc(s.key) + '" tabindex="0" role="button" d="' + d + '" fill="' + s.color + '"/>';
    });
    slices.forEach((s) => {
      const [ex, ey] = pt(cx, cy, r, s.mid);
      const [elx, ely] = pt(cx, cy, r + 22, s.labelAngle);
      const side = Math.cos(s.labelAngle * Math.PI / 180) >= 0 ? 1 : -1;
      const tx0 = elx + side * 20, ty0 = ely;
      const anchor = side > 0 ? 'start' : 'end';
      const tx = tx0 + (side > 0 ? 4 : -4);
      const dim = activeFilters.value.size > 0 && !activeFilters.value.has(s.key);
      svg += '<polyline points="' + ex.toFixed(2) + ',' + ey.toFixed(2) + ' ' + elx.toFixed(2) + ',' + ely.toFixed(2) + ' ' + tx0.toFixed(2) + ',' + ty0.toFixed(2) + '" fill="none" stroke="var(--border)" stroke-width="1.2"/>';
      svg += '<g class="pie-label' + (dim ? ' dim' : '') + '" data-field="value" data-value="' + esc(s.key) + '" tabindex="0" role="button">' +
        '<text x="' + tx.toFixed(2) + '" y="' + (ty0 - 3).toFixed(2) + '" text-anchor="' + anchor + '" class="pie-label-cat">' + esc(s.label) + '</text>' +
        '<text x="' + tx.toFixed(2) + '" y="' + (ty0 + 10).toFixed(2) + '" text-anchor="' + anchor + '" class="pie-label-val" fill="' + s.color + '">' + s.n + ' · ' + s.pctVal + '%</text>' +
      '</g>';
    });
    svgEl.innerHTML = svg;

    svgEl.querySelectorAll('.pie-slice, .pie-label').forEach((el) => {
      const toggle = () => toggleFilter('value', el.dataset.value);
      el.addEventListener('click', toggle);
      el.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); } });
    });
  }

  function renderBreakdownDetail(leads) {
    const wrap = document.getElementById('breakdownDetail');
    if (!activeFilterCount()) { wrap.style.display = 'none'; wrap.innerHTML = ''; return; }
    const rows = leads.filter((r) => leadMatchesAll(r)).slice().sort((a, b) => (a.time < b.time ? 1 : -1));

    let chips = '';
    FILTER_FIELDS.forEach((f) => {
      activeFilters[f].forEach((v) => {
        chips += '<span class="filter-chip" data-field="' + esc(f) + '" data-value="' + esc(v) + '" title="' + esc(BREAKDOWN_FIELD_LABELS[f] || f) + '">' + esc(v) +
          '<button type="button" class="filter-chip-x" aria-label="Remove ' + esc(v) + '">&times;</button></span>';
      });
    });

    let html = '<div class="detail-panel-head"><span>' + fmtNum(rows.length) + ' lead' + (rows.length === 1 ? '' : 's') + ' match' + (rows.length === 1 ? 'es' : '') + ' your selection</span>' +
      '<button type="button" class="detail-panel-close" aria-label="Clear all filters" title="Clear all filters">Clear all</button></div>' +
      '<div class="filter-chip-row">' + chips + '</div>';

    // Sales reps behind these leads (All Executives + 2 or more filters).
    const execSel = document.getElementById('execSelect');
    if (!execSel.value) {
      if (activeFilterCount() >= 2 && rows.length) {
        const byRep = {};
        rows.forEach((r) => { const e = r.exec || '(no executive)'; byRep[e] = (byRep[e] || 0) + 1; });
        const reps = Object.keys(byRep).sort((a, b) => byRep[b] - byRep[a] || a.localeCompare(b));
        html += '<div class="rep-row"><span class="rep-row-label">' + reps.length + ' sales rep' + (reps.length === 1 ? '' : 's') + ' with these leads</span>' +
          reps.map((e) => '<button type="button" class="rep-chip" data-exec="' + esc(e) + '" title="Open ' + esc(e) + '\'s view with these filters">' +
            esc(e) + '<span class="rep-n">' + fmtNum(byRep[e]) + '</span></button>').join('') + '</div>';
      } else if (activeFilterCount() === 1) {
        html += '<p class="rep-hint">Add one more filter to see which sales reps these leads belong to.</p>';
      }
    } else if (activeFilterCount() >= 2) {
      html += '<div class="rep-back-row"><span>Showing <strong>' + esc(execSel.value) + '</strong> only</span><button type="button" class="rep-back">&larr; Back to all executives</button></div>';
    }
    if (rows.length === 0) {
      html += '<p class="empty-state" style="padding:24px;">No leads match this combination.</p>';
    } else {
      html += '<div class="detail-panel-scroll"><table class="leads detail-panel-table"><thead><tr><th>Lead</th><th>Time</th><th>Est. Project Value</th><th>Type of Lead</th><th>Location</th><th>Requirement Timeline</th><th>Next Follow-up</th></tr></thead><tbody>';
      rows.forEach((r) => {
        html += '<tr><td class="lead-cell">' + esc(r.lead || '(no lead logged)') + '</td>' +
          '<td class="time-cell">' + esc(r.time || '') + '</td>' +
          '<td>' + esc(r.value || '') + '</td>' +
          '<td>' + esc(r.leadType || '') + '</td>' +
          '<td class="loc-cell">' + esc(r.location || '') + '</td>' +
          '<td>' + esc(r.timeline || '') + '</td>' +
          '<td>' + esc(r.followUp || '') + '</td></tr>';
      });
      html += '</tbody></table></div>';
    }
    wrap.innerHTML = html;
    wrap.style.display = 'block';
    wrap.querySelector('.detail-panel-close').addEventListener('click', () => {
      FILTER_FIELDS.forEach((f) => activeFilters[f].clear());
      renderBreakdown(filteredRecords());
    });
    const switchExec = (value) => {
      const sel = document.getElementById('execSelect');
      sel.value = value;
      sel.dispatchEvent(new Event('change'));
      document.getElementById('breakdown').scrollIntoView({ behavior: 'smooth', block: 'start' });
    };
    wrap.querySelectorAll('.rep-chip').forEach((btn) => btn.addEventListener('click', () => switchExec(btn.dataset.exec)));
    const back = wrap.querySelector('.rep-back');
    if (back) back.addEventListener('click', () => switchExec(''));
    wrap.querySelectorAll('.filter-chip-x').forEach((btn) => {
      btn.addEventListener('click', () => {
        const chip = btn.closest('.filter-chip');
        activeFilters[chip.dataset.field].delete(chip.dataset.value);
        renderBreakdown(filteredRecords());
      });
    });
  }

  function renderBreakdown(records) {
    const leads = records.filter((r) => r.hasLead);
    // Cross-filter (added 23-Sep-2026): each panel is computed against the
    // leads matching the OTHER two categories' active picks, so selecting in
    // one panel reshapes the other two - and the pie - rather than leaving
    // them static. See leadMatchesOthers above.
    const timelineLeads = leads.filter((r) => leadMatchesOthers(r, 'timeline'));
    const leadTypeLeads = leads.filter((r) => leadMatchesOthers(r, 'leadType'));
    const valueLeads = leads.filter((r) => leadMatchesOthers(r, 'value'));
    const grid = document.getElementById('breakdownGrid');
    grid.innerHTML =
      renderBreakdownPanel('timeline', 'Project Requirement Timeline', orderedCounts(timelineLeads, 'timeline', TIMELINE_ORDER), timelineLeads.length, TIMELINE_ICONS, TIMELINE_SHORT_LABELS) +
      renderBreakdownPanel('leadType', 'Type of Lead', orderedCounts(leadTypeLeads, 'leadType', ['Architect']), leadTypeLeads.length, LEADTYPE_ICONS, LEADTYPE_SHORT_LABELS);

    grid.querySelectorAll('.bar-row').forEach((el) => {
      const toggle = () => toggleFilter(el.dataset.field, el.dataset.value);
      el.addEventListener('click', toggle);
      el.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
      });
    });

    renderPie(valueLeads);
    renderBreakdownDetail(leads);
  }

  // ---- missed days (Section 5) -----------------------------------------
  // A record is classed "Sunday" here if its Day Type string mentions
  // Sunday at all (even on the rare day that's both a Sunday and a public
  // holiday) - "pure holiday" is a holiday that falls on a non-Sunday.
  // Matches the classification used in the 23-Sep-2026 independent data
  // audit, so these three pill counts always foot to the table below.
  function dayTypeIcon(dayType) {
    if (!dayType) {
      return '<span class="daytype-icon gap" title="No Sunday or public holiday explains this day"><svg class="icon"><use href="#icon-alert-triangle"></use></svg></span>';
    }
    if (dayType.indexOf('Sunday') !== -1) {
      return '<span class="daytype-icon sun" title="' + esc(dayType) + '"><svg class="icon"><use href="#icon-sun"></use></svg></span>';
    }
    return '<span class="daytype-icon holiday" title="' + esc(dayType) + '"><svg class="icon"><use href="#icon-star"></use></svg></span>';
  }

  function renderNoCheckin() {
    const records = noCheckinRecords();
    const unexplained = records.filter((r) => !r.dayType).length;
    const sunday = records.filter((r) => r.dayType && r.dayType.indexOf('Sunday') !== -1).length;
    const holiday = records.filter((r) => r.dayType && r.dayType.indexOf('Sunday') === -1).length;

    document.getElementById('noCheckinPillRow').innerHTML =
      '<div class="pill">' +
        '<span class="daytype-icon gap" style="width:32px;height:32px;"><svg class="icon" style="width:16px;height:16px;"><use href="#icon-alert-triangle"></use></svg></span>' +
        '<div><div class="pv num">' + fmtNum(unexplained) + '</div><div class="pl">Genuinely unexplained</div></div>' +
      '</div>' +
      '<div class="pill">' +
        '<span class="daytype-icon sun" style="width:32px;height:32px;"><svg class="icon" style="width:16px;height:16px;"><use href="#icon-sun"></use></svg></span>' +
        '<div><div class="pv num">' + fmtNum(sunday) + '</div><div class="pl">Sundays</div></div>' +
      '</div>' +
      '<div class="pill">' +
        '<span class="daytype-icon holiday" style="width:32px;height:32px;"><svg class="icon" style="width:16px;height:16px;"><use href="#icon-star"></use></svg></span>' +
        '<div><div class="pv num">' + fmtNum(holiday) + '</div><div class="pl">Public holidays</div></div>' +
      '</div>';

    document.getElementById('noCheckinInfoPanel').innerHTML =
      'Days where a sales executive had a LystLoc entry but logged no Checkin Time at all - no visit actually happened, so these are excluded from Forms Submitted above and listed here instead, for the current Month / Week / Executive selection. ' +
      'Day Type flags a Sunday or a national public holiday on that date, as a possible explanation; hover an icon in the table to see which one.';

    const body = document.getElementById('noCheckinBody');
    body.innerHTML = '';
    if (records.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 4; td.className = 'empty-state';
      td.textContent = 'No missing check-ins in this selection.';
      tr.appendChild(td); body.appendChild(tr);
    } else {
      records.slice()
        .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : (a.exec < b.exec ? -1 : 1)))
        .forEach((r) => {
          const tr = document.createElement('tr');
          tr.innerHTML =
            '<td class="lead-cell">' + esc(r.exec) + '</td>' +
            '<td>' + esc(r.dateDisplay) + '</td>' +
            '<td>' + dayTypeIcon(r.dayType) + '</td>' +
            '<td>' + esc(r.weekLabel) + '</td>';
          body.appendChild(tr);
        });
    }
    const note = document.getElementById('noCheckinNote');
    note.textContent = fmtNum(records.length) + ' day' + (records.length === 1 ? '' : 's') +
      ' with no check-in details entered in this selection.';
  }

  // ---- footer / subtitle ----------------------------------------------
  // Shortened to one line 23-Sep-2026; the full disclosure text that used
  // to sit here by default now lives in #footerInfoPanel, one click away
  // via the "Full methodology & caveats" button - nothing below was cut,
  // only moved.
  function fmtISOShort(iso) {
    const d = new Date(iso + 'T00:00:00');
    return d.getDate() + ' ' + d.toLocaleString('en-US', { month: 'short' }) + ' ' + d.getFullYear();
  }

  function renderFooter() {
    const weeks = DATA.weeks.map((w) => w.label);
    const unnamed = DATA.records.filter((r) => r.unnamedLead).length;
    const allDates = DATA.records.map((r) => r.date).sort();
    const dateRange = allDates.length ? (fmtISOShort(allDates[0]) + ' – ' + fmtISOShort(allDates[allDates.length - 1])) : '';

    document.getElementById('footerNote').innerHTML =
      esc(dateRange) +
      ' &middot; ' + DATA.execs.length + ' executive' + (DATA.execs.length === 1 ? '' : 's') +
      ' &middot; ' + fmtNum(DATA.meta.totalRecords) + ' check-ins &middot; Prequate Advisory';

    let full =
      'Source: ' + DATA.meta.generatedFrom + ' &middot; ' + fmtNum(DATA.meta.totalRecords) +
      ' check-in records across ' + DATA.weeks.length + ' week' + (DATA.weeks.length === 1 ? '' : 's') +
      ' (' + weeks.join(', ') + ') and ' + DATA.execs.length + ' field executives. ' +
      '&ldquo;New&rdquo; vs &ldquo;Existing&rdquo; is based on whether a lead\'s name has appeared in an earlier tracked week, checked against all past tracked data - not a confirmed CRM status, and every lead in the very first tracked week shows as New by definition. ' +
      'Follow-up Watch status relies on best-effort parsing of free-text fields (Next Follow-up Date, Checkin Location) typed by field reps - flagged as such, never silently guessed. ' +
      'Days with a LystLoc entry but no Checkin Time logged are treated as no visit having happened, excluded from Forms Submitted, and listed under Missed Days instead. ' +
      'Figures are drawn directly from LystLoc check-in submissions and have not been independently audited.';
    if (unnamed > 0) {
      full += ' <strong>' + fmtNum(unnamed) + ' lead' + (unnamed === 1 ? '' : 's') +
        ' this period ' + (unnamed === 1 ? 'was' : 'were') +
        ' logged with no client name at all (shown as &ldquo;no name logged&rdquo;) - each counted as its own New lead rather than matched to any other unnamed row, since there\'s nothing to match on.</strong>';
    }
    document.getElementById('footerInfoPanel').innerHTML = full;
  }

  function renderAll() {
    const records = filteredRecords();
    renderOverview(records);
    renderBreakdown(records);
    renderLeadsTable('newLeads', records);
    renderLeadsTable('existingLeads', records);
    renderStale();
    renderNoCheckin();
  }

  // ---- generic "ⓘ" info-button toggle ---------------------------------
  // One delegated listener for every .info-btn on the page (reconciliation,
  // Follow-up Watch, Missed Days, footer) - each just flips its own target
  // panel's "open" class. The buttons are static markup so this is wired
  // once, not re-attached on every render.
  document.addEventListener('click', (e) => {
    const btn = e.target.closest('.info-btn');
    if (!btn) return;
    const panel = document.getElementById(btn.dataset.target);
    if (panel) panel.classList.toggle('open');
  });

  populateSelect('monthSelect', DATA.months.map((m) => ({ value: m, label: m })), 'All Months');
  refreshWeekOptions();
  populateSelect('execSelect', DATA.execs.map((e) => ({ value: e, label: e })), 'All Executives');
  renderFooter();

  document.getElementById('monthSelect').addEventListener('change', () => { refreshWeekOptions(); renderAll(); });
  document.getElementById('weekSelect').addEventListener('change', renderAll);
  document.getElementById('execSelect').addEventListener('change', renderAll);
  document.getElementById('overduePdfBtn').addEventListener('click', exportOverduePdf);
  document.getElementById('newLeadsSearch').addEventListener('input', () => renderLeadsTable('newLeads', filteredRecords()));
  document.getElementById('existingLeadsSearch').addEventListener('input', () => renderLeadsTable('existingLeads', filteredRecords()));

  // ---- Phone layout (24-Sep-2026) ---------------------------------------
  // Cards: every table cell gets a data-label copied from its column header,
  // which the phone CSS prints beside the value. Show more: on a phone each
  // list shows 20 rows at a time. Both run from one MutationObserver, so
  // every existing renderer (and any future one) is covered without
  // touching it. A full re-render (search, filter, sort) resets the list to
  // its first 20; opening or closing a visit history does not.
  const PHONE_MQ = window.matchMedia('(max-width: 640px)');
  const PAGE_STEP = 20;
  function headLabels(table) {
    const head = table.tHead;
    if (!head || !head.rows.length) return [];
    return [...head.rows[0].cells].map((th) => {
      const t = [...th.childNodes].filter((n) => n.nodeType === 3).map((n) => n.textContent).join('').trim();
      return t || th.textContent.trim();
    });
  }
  function labelTable(table) {
    const labels = headLabels(table);
    const body = table.tBodies[0];
    if (!body || !labels.length) return;
    [...body.rows].forEach((tr) => {
      if (tr.classList.contains('history-row')) return;
      [...tr.cells].forEach((td, i) => {
        if (td.colSpan > 1) return;
        if (td.dataset.label !== labels[i]) td.dataset.label = labels[i] || '';
      });
    });
  }
  function paginate(table) {
    const body = table.tBodies[0];
    if (!body) return;
    const rows = [...body.rows].filter((tr) => !tr.classList.contains('history-row') && !tr.querySelector(':scope > td.empty-state'));
    const phone = PHONE_MQ.matches;
    const limit = table._mLimit || PAGE_STEP;
    rows.forEach((tr, i) => {
      const hide = phone && i >= limit;
      tr.classList.toggle('m-hidden', hide);
      const nx = tr.nextElementSibling;
      if (nx && nx.classList.contains('history-row')) nx.classList.toggle('m-hidden', hide);
    });
    let btn = table._mMore;
    if (!btn || !btn.isConnected) {
      btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'm-more';
      btn.addEventListener('click', () => { table._mLimit = (table._mLimit || PAGE_STEP) + PAGE_STEP; paginate(table); });
      (table.closest('.table-wrap, .detail-panel-scroll') || table).after(btn);
      table._mMore = btn;
    }
    const remaining = rows.length - limit;
    if (phone && remaining > 0) {
      btn.style.display = '';
      btn.textContent = 'Show ' + Math.min(PAGE_STEP, remaining) + ' more  ·  ' + fmtNum(limit) + ' of ' + fmtNum(rows.length) + ' shown';
    } else {
      btn.style.display = 'none';
    }
  }
  function isCardTable(t) { return t && t.tagName === 'TABLE' && (t.classList.contains('leads') || t.classList.contains('hist-table')); }
  function refreshTable(t) { labelTable(t); if (t.classList.contains('leads')) paginate(t); }
  new MutationObserver((muts) => {
    const touched = new Set();
    muts.forEach((m) => {
      m.addedNodes.forEach((n) => {
        if (n.nodeType !== 1) return;
        if (isCardTable(n)) touched.add(n);
        n.querySelectorAll('table.leads, table.hist-table').forEach((t) => touched.add(t));
      });
      const el = m.target.nodeType === 1 ? m.target : m.target.parentElement;
      const t = el && el.closest('table');
      if (!isCardTable(t)) return;
      touched.add(t);
      if (t.classList.contains('leads') && m.target === t.tBodies[0] &&
          [...m.removedNodes].some((n) => n.nodeType === 1 && n.tagName === 'TR' && !n.classList.contains('history-row'))) {
        t._mLimit = PAGE_STEP;
      }
    });
    touched.forEach(refreshTable);
  }).observe(document.body, { childList: true, subtree: true });
  const onPhoneChange = () => document.querySelectorAll('table.leads').forEach(paginate);
  if (PHONE_MQ.addEventListener) PHONE_MQ.addEventListener('change', onPhoneChange); else PHONE_MQ.addListener(onPhoneChange);

  // Filters button and one-line summary (phone header)
  const mastheadEl = document.querySelector('.masthead');
  const filtersBtn = document.getElementById('filtersBtn');
  filtersBtn.addEventListener('click', () => {
    const open = mastheadEl.classList.toggle('filters-open');
    filtersBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
  });
  function updateFilterSummary() {
    const ids = ['monthSelect', 'weekSelect', 'execSelect'];
    const parts = ids.map((id) => { const el = document.getElementById(id); const o = el.options[el.selectedIndex]; return o ? o.text : ''; });
    const active = ids.filter((id) => document.getElementById(id).value).length;
    document.getElementById('filterSummary').textContent = parts.join('  ·  ');
    document.getElementById('filtersCount').textContent = active ? String(active) : '';
    filtersBtn.classList.toggle('has-active', active > 0);
  }
  ['monthSelect', 'weekSelect', 'execSelect'].forEach((id) => document.getElementById(id).addEventListener('change', updateFilterSummary));
  updateFilterSummary();

  renderAll();
  document.querySelectorAll('table.leads, table.hist-table').forEach(refreshTable);
})();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
