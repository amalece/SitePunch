# SitePunch — Labour Attendance & Overtime Tracker

A simple, mobile-friendly web app for site supervisors to record daily labourer
attendance and automatically calculate normal + overtime hours.

## What it does

- **Daily Attendance** — pick a date and site, tap each labourer's status
  (Present / Absent / Leave / Holiday / Off), enter start/end time for anyone
  present. Hours and overtime are calculated automatically.
- **Employees** — master list: employee ID, name, designation, site, standard
  hours/day.
- **Sites** — list of projects/locations.
- **Reports** — filter by date range, site, employee, designation, and/or
  status, then view a summary or export to **Excel or a branded PDF**.
- **Company profile** (admin only) — upload your logo, and set your company
  name and address. These appear on the header of every exported PDF report.
- **Backup & Restore** (admin only) — download a single JSON file containing
  everything (sites, employees, attendance history, users, company profile),
  and re-import it later to restore or move the app to a new computer.
- **Users** (admin only) — add supervisor/admin logins.

## Filtering and exporting reports

The Reports page lets you narrow down attendance records by any combination of:

- **Date range** — from/to
- **Site**
- **Employee** — searchable by name or employee ID
- **Designation** — e.g. Mason, Electrician, Helper
- **Status** — Present, Absent, Leave, Holiday, Weekly Off

Both the on-screen summary and the Excel/PDF exports respect whatever filters
are applied, so you can pull something like "all Helpers at Site A who were
absent in March" as easily as a full month's report for everyone. Every
export includes each employee's Employee ID alongside their name.

## How overtime is calculated

```
Worked hours = End time − Start time − Break minutes
Normal hours = min(Worked hours, employee's standard hours/day)
Overtime     = max(0, Worked hours − standard hours/day)
```

Example: 7:00 AM–6:00 PM with a 1-hour break = 10 worked hours.
With an 8-hour standard day, that's 8 normal hours + 2 overtime hours.

## Running it locally

1. Install Python 3.10+ if you don't already have it.
2. Open a terminal in this folder and set up a virtual environment:

   ```bash
   python3 -m venv venv
   source venv/bin/activate        # on Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. Start the app:

   ```bash
   python app.py
   ```

4. Open **http://localhost:5000** in your browser (or from a phone on the same
   Wi-Fi, use the computer's local IP address, e.g. **http://192.168.1.20:5000**).

The first time it runs, it automatically creates the database
(`attendance.db`, a single file — back this up regularly) and one login:

- **Username:** `admin`
- **Password:** `admin123`

**Log in and change this password immediately** by going to Users (top or
bottom navigation, admin only) and editing the admin account with a new
password.

## Day-to-day use

1. Log in.
2. Go to **Sites** and add your project sites.
3. Go to **Employees** and add your labourers — give each one an Employee ID
   if you use one (e.g. from your existing HR/payroll system), assign them to
   a site, and set their designation and standard hours.
4. Each day, go to **Attendance**, pick the site, and mark everyone. Use
   "Mark all present" first, then fix only the exceptions (absences, leave,
   different timings).
5. Use **Reports** any time — filter by date, site, employee, designation, or
   status — and export to Excel or PDF for payroll / client sharing.

## Branding your PDF reports

Go to **Settings → Company Profile** (admin only) and fill in:

- Company name
- Address
- Logo (PNG, JPG, or GIF — a wide logo like 400×120px on a transparent or
  white background looks best)

Once saved, every PDF generated from **Reports → Export to PDF** will show
your logo, name, and address in the header, so it's ready to hand to a
client, payroll office, or head office without any manual editing.

## Backing up and restoring your data

Go to **Settings → Backup & Restore** (admin only):

- **Download backup (.json)** — saves everything (sites, employees,
  attendance history, users, and your company profile) into one JSON file.
  Do this regularly and keep the file somewhere safe (email it to yourself,
  save it to cloud storage, etc.) — it's your only copy outside the
  `attendance.db` file on this computer.
- **Import & replace all data** — upload a previously downloaded backup file
  to restore it. **This replaces every site, employee, attendance record,
  and login currently in the app** with what's in the file, so only do this
  when you're sure. You'll be logged out and asked to log back in afterwards
  since user accounts are also replaced.

This is also how you'd move the app to a new computer: install it fresh,
log in with the default admin account, then immediately import your latest
backup.

## Notes on this first version

This is intentionally the simple version — no GPS, no photos, no approval
workflows, no biometric login. A supervisor records attendance on behalf of
the crew, which is the fastest way to get reliable data flowing. Once this is
in daily use, natural next additions would be:

- A proper leave-type breakdown (annual/sick/unpaid) instead of one "Leave" status
- A holiday calendar so holiday status is applied automatically instead of manually
- Audit history showing who changed a past entry and when
- Locking attendance once payroll has been run for a period

## Deploying somewhere permanent

For real use beyond your own computer, run it on a small cloud server (or
your office server) behind a proper WSGI server such as `gunicorn`, and put
it behind HTTPS. The database can also be migrated from SQLite to PostgreSQL
later without changing the application logic — only the connection string in
`app.py` needs to change.

Also, before going live:
- Change `app.config["SECRET_KEY"]` in `app.py` to a long random string.
- Change the default admin password.
