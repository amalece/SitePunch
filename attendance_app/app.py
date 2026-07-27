import os
import json
from datetime import datetime, date, timedelta
from io import BytesIO
from werkzeug.utils import secure_filename

from flask import (
    Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
)
from flask_login import (
    LoginManager, login_user, logout_user, login_required, current_user
)
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable
)

from models import db, User, Site, Employee, Attendance, CompanySettings

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "static", "uploads")
ALLOWED_LOGO_EXTENSIONS = {"png", "jpg", "jpeg", "gif"}

app = Flask(__name__)
app.config["SECRET_KEY"] = "change-this-secret-key-in-production"
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{os.path.join(BASE_DIR, 'attendance.db')}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

db.init_app(app)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"
login_manager.login_message = "Please log in to continue."


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_time_to_minutes(t):
    """Convert 'HH:MM' string to minutes since midnight. Returns None if invalid."""
    if not t:
        return None
    try:
        h, m = t.split(":")
        return int(h) * 60 + int(m)
    except (ValueError, AttributeError):
        return None


def calculate_hours(start_time, end_time, break_minutes, standard_hours):
    """Returns (worked_hours, normal_hours, overtime_hours)."""
    start_m = parse_time_to_minutes(start_time)
    end_m = parse_time_to_minutes(end_time)
    if start_m is None or end_m is None:
        return 0.0, 0.0, 0.0

    total_minutes = end_m - start_m
    if total_minutes < 0:
        # overnight shift - add 24 hours
        total_minutes += 24 * 60

    break_minutes = break_minutes or 0
    worked_minutes = max(0, total_minutes - break_minutes)
    worked_hours = round(worked_minutes / 60.0, 2)

    standard_hours = standard_hours or 8.0
    normal_hours = round(min(worked_hours, standard_hours), 2)
    overtime_hours = round(max(0.0, worked_hours - standard_hours), 2)

    return worked_hours, normal_hours, overtime_hours


def require_admin():
    return current_user.is_authenticated and current_user.role == "admin"


def get_company_settings():
    settings = CompanySettings.query.first()
    if not settings:
        settings = CompanySettings(company_name="", address="", logo_filename=None)
        db.session.add(settings)
        db.session.commit()
    return settings


# ---------------------------------------------------------------------------
# Auth routes
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username).first()
        if user and user.active and user.check_password(password):
            login_user(user)
            next_page = request.args.get("next")
            return redirect(next_page or url_for("dashboard"))
        flash("Invalid username or password.", "error")
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

@app.route("/")
@login_required
def dashboard():
    today = date.today()
    total_active = Employee.query.filter_by(active=True).count()
    today_records = Attendance.query.filter_by(date=today).all()

    present = sum(1 for r in today_records if r.status == "present")
    absent = sum(1 for r in today_records if r.status == "absent")
    on_leave = sum(1 for r in today_records if r.status == "leave")
    not_entered = total_active - len(today_records)
    total_ot = round(sum(r.overtime_hours or 0 for r in today_records), 2)

    sites = Site.query.filter_by(active=True).all()
    site_summary = []
    for s in sites:
        site_emp_ids = [e.id for e in s.employees if e.active]
        site_present = sum(
            1 for r in today_records
            if r.employee_id in site_emp_ids and r.status == "present"
        )
        site_summary.append({"name": s.name, "present": site_present, "total": len(site_emp_ids)})

    return render_template(
        "dashboard.html",
        today=today,
        total_active=total_active,
        present=present,
        absent=absent,
        on_leave=on_leave,
        not_entered=max(0, not_entered),
        total_ot=total_ot,
        site_summary=site_summary,
    )


# ---------------------------------------------------------------------------
# Daily attendance
# ---------------------------------------------------------------------------

@app.route("/attendance", methods=["GET"])
@login_required
def attendance():
    selected_date_str = request.args.get("date") or date.today().isoformat()
    selected_site_id = request.args.get("site_id", type=int)

    sites = Site.query.filter_by(active=True).order_by(Site.name).all()
    if not selected_site_id and sites:
        selected_site_id = sites[0].id

    employees = []
    existing = {}
    if selected_site_id:
        employees = (
            Employee.query.filter_by(site_id=selected_site_id, active=True)
            .order_by(Employee.name)
            .all()
        )
        selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d").date()
        records = Attendance.query.filter(
            Attendance.date == selected_date,
            Attendance.employee_id.in_([e.id for e in employees]) if employees else False,
        ).all()
        existing = {r.employee_id: r for r in records}

    return render_template(
        "attendance.html",
        sites=sites,
        selected_site_id=selected_site_id,
        selected_date=selected_date_str,
        employees=employees,
        existing=existing,
    )


@app.route("/attendance/save", methods=["POST"])
@login_required
def attendance_save():
    selected_date_str = request.form.get("date")
    selected_site_id = request.form.get("site_id", type=int)
    selected_date = datetime.strptime(selected_date_str, "%Y-%m-%d").date()

    employees = Employee.query.filter_by(site_id=selected_site_id, active=True).all()

    for emp in employees:
        status = request.form.get(f"status_{emp.id}", "absent")
        start_time = request.form.get(f"start_{emp.id}", "").strip()
        end_time = request.form.get(f"end_{emp.id}", "").strip()
        break_minutes = request.form.get(f"break_{emp.id}", type=int) or 60
        remarks = request.form.get(f"remarks_{emp.id}", "").strip()

        worked_hours = normal_hours = overtime_hours = 0.0
        if status == "present":
            worked_hours, normal_hours, overtime_hours = calculate_hours(
                start_time, end_time, break_minutes, emp.standard_hours
            )
        else:
            start_time = None
            end_time = None

        record = Attendance.query.filter_by(employee_id=emp.id, date=selected_date).first()
        if not record:
            record = Attendance(employee_id=emp.id, date=selected_date)
            db.session.add(record)

        record.status = status
        record.start_time = start_time
        record.end_time = end_time
        record.break_minutes = break_minutes if status == "present" else None
        record.worked_hours = worked_hours
        record.normal_hours = normal_hours
        record.overtime_hours = overtime_hours
        record.remarks = remarks
        record.recorded_by = current_user.username
        record.recorded_at = datetime.utcnow()

    db.session.commit()
    flash("Attendance saved.", "success")
    return redirect(url_for("attendance", date=selected_date_str, site_id=selected_site_id))


# ---------------------------------------------------------------------------
# Employee master
# ---------------------------------------------------------------------------

@app.route("/employees")
@login_required
def employees():
    all_employees = Employee.query.order_by(Employee.active.desc(), Employee.name).all()
    sites = Site.query.filter_by(active=True).order_by(Site.name).all()
    return render_template("employees.html", employees=all_employees, sites=sites)


@app.route("/employees/save", methods=["POST"])
@login_required
def employee_save():
    emp_id = request.form.get("id", type=int)
    employee_code = request.form.get("employee_code", "").strip()
    name = request.form.get("name", "").strip()
    designation = request.form.get("designation", "").strip()
    site_id = request.form.get("site_id", type=int) or None
    standard_hours = request.form.get("standard_hours", type=float) or 8.0
    active = bool(request.form.get("active"))

    if not name:
        flash("Employee name is required.", "error")
        return redirect(url_for("employees"))

    if employee_code:
        existing_code = Employee.query.filter_by(employee_code=employee_code).first()
        if existing_code and existing_code.id != emp_id:
            flash(f"Employee ID '{employee_code}' is already used by {existing_code.name}.", "error")
            return redirect(url_for("employees"))

    if emp_id:
        emp = Employee.query.get(emp_id)
        if not emp:
            flash("Employee not found.", "error")
            return redirect(url_for("employees"))
    else:
        emp = Employee()
        db.session.add(emp)

    emp.employee_code = employee_code or None
    emp.name = name
    emp.designation = designation
    emp.site_id = site_id
    emp.standard_hours = standard_hours
    emp.active = active

    db.session.commit()
    flash(f"Employee '{name}' saved.", "success")
    return redirect(url_for("employees"))


@app.route("/employees/delete/<int:emp_id>", methods=["POST"])
@login_required
def employee_delete(emp_id):
    emp = Employee.query.get(emp_id)
    if emp:
        emp.active = False
        db.session.commit()
        flash(f"Employee '{emp.name}' deactivated.", "success")
    return redirect(url_for("employees"))


# ---------------------------------------------------------------------------
# Site master
# ---------------------------------------------------------------------------

@app.route("/sites")
@login_required
def sites():
    all_sites = Site.query.order_by(Site.active.desc(), Site.name).all()
    return render_template("sites.html", sites=all_sites)


@app.route("/sites/save", methods=["POST"])
@login_required
def site_save():
    site_id = request.form.get("id", type=int)
    name = request.form.get("name", "").strip()
    location = request.form.get("location", "").strip()
    active = bool(request.form.get("active"))

    if not name:
        flash("Site name is required.", "error")
        return redirect(url_for("sites"))

    if site_id:
        site = Site.query.get(site_id)
        if not site:
            flash("Site not found.", "error")
            return redirect(url_for("sites"))
    else:
        site = Site()
        db.session.add(site)

    site.name = name
    site.location = location
    site.active = active

    db.session.commit()
    flash(f"Site '{name}' saved.", "success")
    return redirect(url_for("sites"))


@app.route("/sites/delete/<int:site_id>", methods=["POST"])
@login_required
def site_delete(site_id):
    site = Site.query.get(site_id)
    if site:
        site.active = False
        db.session.commit()
        flash(f"Site '{site.name}' deactivated.", "success")
    return redirect(url_for("sites"))


# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------

STATUS_OPTIONS = [
    ("present", "Present"),
    ("absent", "Absent"),
    ("leave", "Leave"),
    ("holiday", "Holiday"),
    ("weekly_off", "Weekly Off"),
]


def _parse_report_filters():
    today = date.today()
    start_str = request.args.get("start_date") or (today - timedelta(days=7)).isoformat()
    end_str = request.args.get("end_date") or today.isoformat()
    site_id = request.args.get("site_id", type=int)
    employee_id = request.args.get("employee_id", type=int)
    designation = request.args.get("designation", "").strip()
    status = request.args.get("status", "").strip()
    return start_str, end_str, site_id, employee_id, designation, status


@app.route("/reports")
@login_required
def reports():
    start_str, end_str, site_id, employee_id, designation, status = _parse_report_filters()

    records, summary = _compute_report_data(start_str, end_str, site_id, employee_id, designation, status)

    sites = Site.query.filter_by(active=True).order_by(Site.name).all()
    employees = Employee.query.order_by(Employee.name).all()
    designations = sorted({
        e.designation for e in employees if e.designation
    })

    return render_template(
        "reports.html",
        records=records,
        summary=summary,
        sites=sites,
        employees=employees,
        designations=designations,
        status_options=STATUS_OPTIONS,
        start_date=start_str,
        end_date=end_str,
        selected_site_id=site_id,
        selected_employee_id=employee_id,
        selected_designation=designation,
        selected_status=status,
    )


@app.route("/reports/export")
@login_required
def reports_export():
    start_str, end_str, site_id, employee_id, designation, status = _parse_report_filters()
    records, _ = _compute_report_data(start_str, end_str, site_id, employee_id, designation, status)

    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    headers = [
        "Employee ID", "Employee", "Designation", "Site", "Date", "Status", "Start Time", "End Time",
        "Break (min)", "Worked Hours", "Normal Hours", "Overtime Hours", "Remarks",
    ]
    ws.append(headers)

    for r in records:
        ws.append([
            r.employee.employee_code or "",
            r.employee.name,
            r.employee.designation or "",
            r.employee.site.name if r.employee.site else "",
            r.date.isoformat(),
            r.status,
            r.start_time or "",
            r.end_time or "",
            r.break_minutes or "",
            r.worked_hours or 0,
            r.normal_hours or 0,
            r.overtime_hours or 0,
            r.remarks or "",
        ])

    for col_cells in ws.columns:
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = max(10, max_len + 2)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"attendance_{start_str}_to_{end_str}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# PDF report export (branded with company logo / name / address)
# ---------------------------------------------------------------------------

BRAND = colors.HexColor("#1d4ed8")
BRAND_SOFT = colors.HexColor("#eaf0fe")
ACCENT = colors.HexColor("#e08a1e")
INK = colors.HexColor("#12151f")
DIM = colors.HexColor("#667085")
LINE_C = colors.HexColor("#dde1e7")
SURFACE_RAISED = colors.HexColor("#f4f5f7")
OK_C = colors.HexColor("#17845a")
BAD_C = colors.HexColor("#c8323b")


def build_attendance_pdf(records, summary, start_str, end_str, site_name, company):
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
        title="Attendance Report",
    )
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle("title", parent=styles["Normal"], fontName="Helvetica-Bold",
                                  fontSize=18, textColor=INK, leading=22)
    style_sub = ParagraphStyle("sub", parent=styles["Normal"], fontName="Helvetica",
                                fontSize=9.5, textColor=DIM, leading=13)
    style_company = ParagraphStyle("company", parent=styles["Normal"], fontName="Helvetica-Bold",
                                    fontSize=13, textColor=INK, leading=16, alignment=TA_RIGHT)
    style_addr = ParagraphStyle("addr", parent=styles["Normal"], fontName="Helvetica",
                                 fontSize=8.5, textColor=DIM, leading=11, alignment=TA_RIGHT)
    style_h3 = ParagraphStyle("h3", parent=styles["Normal"], fontName="Helvetica-Bold",
                               fontSize=11, textColor=INK, spaceBefore=6, spaceAfter=6)
    style_cell = ParagraphStyle("cell", parent=styles["Normal"], fontName="Helvetica", fontSize=8)
    style_sum_cell = ParagraphStyle("sum_cell", parent=styles["Normal"], fontName="Helvetica",
                                     fontSize=8.5, textColor=INK, leading=10.5)
    style_entry_cell = ParagraphStyle("entry_cell", parent=styles["Normal"], fontName="Helvetica",
                                       fontSize=7.5, textColor=INK, leading=9.5)

    story = []

    # ---- Header row 1: logo (left) + company name/address (right) ----
    left_flow = []
    logo_path = None
    if company and company.logo_filename:
        candidate = os.path.join(UPLOAD_FOLDER, company.logo_filename)
        if os.path.exists(candidate):
            logo_path = candidate
    if logo_path:
        try:
            img = Image(logo_path)
            img._restrictSize(70 * mm, 30 * mm)
            left_flow.append(img)
        except Exception:
            pass

    right_flow = []
    if company and company.company_name:
        right_flow.append(Paragraph(company.company_name, style_company))
    if company and company.address:
        right_flow.append(Paragraph(company.address.replace("\n", "<br/>"), style_addr))

    if left_flow or right_flow:
        brand_header = Table(
            [[left_flow or "", right_flow or ""]],
            colWidths=[90 * mm, None],
        )
        brand_header.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (0, 0), (0, 0), "LEFT"),
            ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(brand_header)
        story.append(HRFlowable(width="100%", thickness=0.75, color=LINE_C, spaceAfter=12))

    # ---- Header row 2: report title + date range/site ----
    story.append(Paragraph("Attendance Report", style_title))
    story.append(Paragraph(
        f"{start_str} to {end_str}"
        + (f" &nbsp;\u00b7&nbsp; {site_name}" if site_name else " &nbsp;\u00b7&nbsp; All sites"),
        style_sub,
    ))
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=1.2, color=BRAND, spaceAfter=14))

    # ---- Summary by employee ----
    story.append(Paragraph("Summary by employee", style_h3))
    sum_header = ["Employee ID", "Employee", "Present", "Absent", "Leave", "Normal hrs", "Overtime hrs"]
    sum_rows = [sum_header]
    for s in summary.values():
        sum_rows.append([
            s.get("employee_code") or "\u2014",
            Paragraph(s["name"], style_sum_cell),
            str(s["present_days"]), str(s["absent_days"]), str(s["leave_days"]),
            f'{s["normal_hours"]:.1f}', f'{s["overtime_hours"]:.1f}',
        ])
    if len(sum_rows) == 1:
        sum_rows.append(["", Paragraph("No attendance recorded in this range.", style_sum_cell), "", "", "", "", ""])

    sum_table = Table(sum_rows, colWidths=[22 * mm, 42 * mm, 20 * mm, 20 * mm, 18 * mm, 24 * mm, 26 * mm], repeatRows=1)
    sum_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BRAND),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, SURFACE_RAISED]),
        ("TEXTCOLOR", (6, 1), (6, -1), ACCENT),
        ("BOX", (0, 0), (-1, -1), 0.6, LINE_C),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, LINE_C),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(sum_table)
    story.append(Spacer(1, 16))

    # ---- Daily entries ----
    story.append(Paragraph(f"Daily entries ({len(records)})", style_h3))
    entry_header = ["Emp ID", "Date", "Employee", "Site", "Status", "Start", "End", "Worked", "Normal", "OT"]
    entry_rows = [entry_header]
    for r in records:
        entry_rows.append([
            r.employee.employee_code or "\u2014",
            r.date.isoformat(),
            Paragraph(r.employee.name, style_entry_cell),
            Paragraph(r.employee.site.name if r.employee.site else "\u2014", style_entry_cell),
            r.status.replace("_", " ").title(),
            r.start_time or "\u2014",
            r.end_time or "\u2014",
            f"{r.worked_hours or 0:.1f}",
            f"{r.normal_hours or 0:.1f}",
            f"{r.overtime_hours or 0:.1f}",
        ])
    if len(entry_rows) == 1:
        entry_rows.append(["", "", Paragraph("No records", style_entry_cell), "", "", "", "", "", "", ""])

    entry_table = Table(
        entry_rows,
        colWidths=[14 * mm, 15 * mm, 33 * mm, 27 * mm, 16 * mm, 12 * mm, 12 * mm, 13 * mm, 13 * mm, 10 * mm],
        repeatRows=1,
    )
    entry_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), INK),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, SURFACE_RAISED]),
        ("TEXTCOLOR", (9, 1), (9, -1), ACCENT),
        ("BOX", (0, 0), (-1, -1), 0.6, LINE_C),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE_C),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(entry_table)

    def footer(c, d):
        c.saveState()
        c.setFont("Helvetica", 7.5)
        c.setFillColor(DIM)
        c.drawString(16 * mm, 10 * mm, "Generated by SitePunch")
        c.drawRightString(A4[0] - 16 * mm, 10 * mm, f"Page {d.page}")
        c.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    buf.seek(0)
    return buf


def _compute_report_data(start_str, end_str, site_id=None, employee_id=None, designation=None, status=None):
    start_dt = datetime.strptime(start_str, "%Y-%m-%d").date()
    end_dt = datetime.strptime(end_str, "%Y-%m-%d").date()

    query = Attendance.query.join(Employee).filter(
        Attendance.date >= start_dt, Attendance.date <= end_dt
    )
    if site_id:
        query = query.filter(Employee.site_id == site_id)
    if employee_id:
        query = query.filter(Employee.id == employee_id)
    if designation:
        query = query.filter(Employee.designation == designation)
    if status:
        query = query.filter(Attendance.status == status)
    records = query.order_by(Attendance.date, Employee.name).all()

    summary = {}
    for r in records:
        key = r.employee_id
        if key not in summary:
            summary[key] = {
                "name": r.employee.name,
                "employee_code": r.employee.employee_code or "",
                "present_days": 0, "absent_days": 0,
                "leave_days": 0, "normal_hours": 0.0, "overtime_hours": 0.0,
            }
        s = summary[key]
        if r.status == "present":
            s["present_days"] += 1
        elif r.status == "absent":
            s["absent_days"] += 1
        elif r.status == "leave":
            s["leave_days"] += 1
        s["normal_hours"] += r.normal_hours or 0
        s["overtime_hours"] += r.overtime_hours or 0

    for s in summary.values():
        s["normal_hours"] = round(s["normal_hours"], 2)
        s["overtime_hours"] = round(s["overtime_hours"], 2)

    return records, summary


@app.route("/reports/export_pdf")
@login_required
def reports_export_pdf():
    start_str, end_str, site_id, employee_id, designation, status = _parse_report_filters()
    records, summary = _compute_report_data(start_str, end_str, site_id, employee_id, designation, status)

    site_name = None
    if site_id:
        site = Site.query.get(site_id)
        site_name = site.name if site else None

    filter_labels = []
    if site_name:
        filter_labels.append(site_name)
    if employee_id:
        emp = Employee.query.get(employee_id)
        if emp:
            filter_labels.append(emp.name)
    if designation:
        filter_labels.append(designation)
    if status:
        filter_labels.append(dict(STATUS_OPTIONS).get(status, status))
    site_name = " \u00b7 ".join(filter_labels) if filter_labels else None

    company = get_company_settings()
    buf = build_attendance_pdf(records, summary, start_str, end_str, site_name, company)

    filename = f"attendance_report_{start_str}_to_{end_str}.pdf"
    return send_file(buf, as_attachment=True, download_name=filename, mimetype="application/pdf")


# ---------------------------------------------------------------------------
# Company profile (logo, name, address used to brand PDF reports)
# ---------------------------------------------------------------------------

@app.route("/settings/company")
@login_required
def company_settings_page():
    if not require_admin():
        flash("Only admins can manage company settings.", "error")
        return redirect(url_for("dashboard"))
    company = get_company_settings()
    return render_template("company_settings.html", company=company)


@app.route("/settings/company/save", methods=["POST"])
@login_required
def company_settings_save():
    if not require_admin():
        flash("Only admins can manage company settings.", "error")
        return redirect(url_for("dashboard"))

    company = get_company_settings()
    company.company_name = request.form.get("company_name", "").strip()
    company.address = request.form.get("address", "").strip()

    logo_file = request.files.get("logo")
    if logo_file and logo_file.filename:
        ext = logo_file.filename.rsplit(".", 1)[-1].lower() if "." in logo_file.filename else ""
        if ext in ALLOWED_LOGO_EXTENSIONS:
            filename = secure_filename(f"company_logo.{ext}")
            # remove any previously saved logo with a different extension
            for existing_ext in ALLOWED_LOGO_EXTENSIONS:
                old_path = os.path.join(UPLOAD_FOLDER, f"company_logo.{existing_ext}")
                if os.path.exists(old_path):
                    os.remove(old_path)
            logo_file.save(os.path.join(UPLOAD_FOLDER, filename))
            company.logo_filename = filename
        else:
            flash("Logo must be a PNG, JPG, or GIF file.", "error")
            return redirect(url_for("company_settings_page"))

    db.session.commit()
    flash("Company profile saved.", "success")
    return redirect(url_for("company_settings_page"))


@app.route("/settings/company/remove_logo", methods=["POST"])
@login_required
def company_logo_remove():
    if not require_admin():
        flash("Only admins can manage company settings.", "error")
        return redirect(url_for("dashboard"))
    company = get_company_settings()
    if company.logo_filename:
        old_path = os.path.join(UPLOAD_FOLDER, company.logo_filename)
        if os.path.exists(old_path):
            os.remove(old_path)
        company.logo_filename = None
        db.session.commit()
        flash("Logo removed.", "success")
    return redirect(url_for("company_settings_page"))


# ---------------------------------------------------------------------------
# Backup (JSON export) and restore (JSON import)
# ---------------------------------------------------------------------------

@app.route("/backup")
@login_required
def backup_page():
    if not require_admin():
        flash("Only admins can back up or restore data.", "error")
        return redirect(url_for("dashboard"))
    return render_template("backup.html")


@app.route("/backup/export")
@login_required
def backup_export():
    if not require_admin():
        flash("Only admins can back up or restore data.", "error")
        return redirect(url_for("dashboard"))

    company = get_company_settings()
    data = {
        "exported_at": datetime.utcnow().isoformat(),
        "app": "SitePunch",
        "version": 1,
        "company_settings": {
            "id": company.id,
            "company_name": company.company_name,
            "address": company.address,
            "logo_filename": company.logo_filename,
        },
        "users": [
            {
                "id": u.id, "username": u.username, "password_hash": u.password_hash,
                "role": u.role, "active": u.active,
            }
            for u in User.query.all()
        ],
        "sites": [
            {
                "id": s.id, "name": s.name, "location": s.location, "active": s.active,
            }
            for s in Site.query.all()
        ],
        "employees": [
            {
                "id": e.id, "employee_code": e.employee_code, "name": e.name,
                "designation": e.designation, "site_id": e.site_id,
                "standard_hours": e.standard_hours, "weekly_off_day": e.weekly_off_day,
                "active": e.active,
            }
            for e in Employee.query.all()
        ],
        "attendance": [
            {
                "id": a.id, "employee_id": a.employee_id, "date": a.date.isoformat(),
                "status": a.status, "start_time": a.start_time, "end_time": a.end_time,
                "break_minutes": a.break_minutes, "worked_hours": a.worked_hours,
                "normal_hours": a.normal_hours, "overtime_hours": a.overtime_hours,
                "remarks": a.remarks, "recorded_by": a.recorded_by,
                "recorded_at": a.recorded_at.isoformat() if a.recorded_at else None,
            }
            for a in Attendance.query.all()
        ],
    }

    buf = BytesIO(json.dumps(data, indent=2).encode("utf-8"))
    filename = f"sitepunch_backup_{date.today().isoformat()}.json"
    return send_file(buf, as_attachment=True, download_name=filename, mimetype="application/json")


@app.route("/backup/import", methods=["POST"])
@login_required
def backup_import():
    if not require_admin():
        flash("Only admins can back up or restore data.", "error")
        return redirect(url_for("dashboard"))

    if not request.form.get("confirm"):
        flash("Please confirm that you understand this will replace all existing data.", "error")
        return redirect(url_for("backup_page"))

    upload = request.files.get("backup_file")
    if not upload or not upload.filename:
        flash("Please choose a backup JSON file to import.", "error")
        return redirect(url_for("backup_page"))

    try:
        data = json.load(upload.stream)
    except (ValueError, UnicodeDecodeError):
        flash("That file could not be read as JSON. Please choose a valid SitePunch backup file.", "error")
        return redirect(url_for("backup_page"))

    required_keys = ["sites", "employees", "attendance", "users"]
    if not all(k in data for k in required_keys):
        flash("This doesn't look like a valid SitePunch backup file.", "error")
        return redirect(url_for("backup_page"))

    try:
        # Wipe existing data (children first to respect foreign keys)
        Attendance.query.delete()
        Employee.query.delete()
        Site.query.delete()
        User.query.delete()
        CompanySettings.query.delete()
        db.session.commit()

        cs = data.get("company_settings")
        if cs:
            db.session.add(CompanySettings(
                id=cs.get("id", 1),
                company_name=cs.get("company_name", ""),
                address=cs.get("address", ""),
                logo_filename=cs.get("logo_filename"),
            ))

        for s in data["sites"]:
            db.session.add(Site(
                id=s["id"], name=s["name"], location=s.get("location"),
                active=s.get("active", True),
            ))

        for e in data["employees"]:
            db.session.add(Employee(
                id=e["id"], employee_code=e.get("employee_code"), name=e["name"],
                designation=e.get("designation"),
                site_id=e.get("site_id"), standard_hours=e.get("standard_hours", 8.0),
                weekly_off_day=e.get("weekly_off_day", 6), active=e.get("active", True),
            ))

        for a in data["attendance"]:
            db.session.add(Attendance(
                id=a["id"], employee_id=a["employee_id"],
                date=datetime.strptime(a["date"], "%Y-%m-%d").date(),
                status=a["status"], start_time=a.get("start_time"), end_time=a.get("end_time"),
                break_minutes=a.get("break_minutes"), worked_hours=a.get("worked_hours", 0.0),
                normal_hours=a.get("normal_hours", 0.0), overtime_hours=a.get("overtime_hours", 0.0),
                remarks=a.get("remarks"), recorded_by=a.get("recorded_by"),
                recorded_at=(datetime.fromisoformat(a["recorded_at"]) if a.get("recorded_at") else datetime.utcnow()),
            ))

        for u in data["users"]:
            db.session.add(User(
                id=u["id"], username=u["username"], password_hash=u["password_hash"],
                role=u.get("role", "supervisor"), active=u.get("active", True),
            ))

        db.session.commit()
    except Exception as exc:
        db.session.rollback()
        flash(f"Import failed and no changes were made: {exc}", "error")
        return redirect(url_for("backup_page"))

    logout_user()
    flash("Backup imported successfully. Please log in again.", "success")
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Users / settings (admin only)
# ---------------------------------------------------------------------------

@app.route("/users")
@login_required
def users():
    if not require_admin():
        flash("Only admins can manage users.", "error")
        return redirect(url_for("dashboard"))
    all_users = User.query.order_by(User.username).all()
    return render_template("users.html", users=all_users)


@app.route("/users/save", methods=["POST"])
@login_required
def user_save():
    if not require_admin():
        flash("Only admins can manage users.", "error")
        return redirect(url_for("dashboard"))

    user_id = request.form.get("id", type=int)
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    role = request.form.get("role", "supervisor")
    active = bool(request.form.get("active"))

    if not username:
        flash("Username is required.", "error")
        return redirect(url_for("users"))

    if user_id:
        user = User.query.get(user_id)
        if not user:
            flash("User not found.", "error")
            return redirect(url_for("users"))
    else:
        if User.query.filter_by(username=username).first():
            flash("Username already exists.", "error")
            return redirect(url_for("users"))
        user = User(username=username)
        db.session.add(user)

    user.role = role
    user.active = active
    if password:
        user.set_password(password)
    elif not user_id:
        flash("Password is required for new users.", "error")
        db.session.rollback()
        return redirect(url_for("users"))

    db.session.commit()
    flash(f"User '{username}' saved.", "success")
    return redirect(url_for("users"))


# ---------------------------------------------------------------------------
# CLI: init db with sample data
# ---------------------------------------------------------------------------

def run_light_migrations():
    """Add any newly-introduced columns to an existing SQLite database.
    Safe to run every startup: only adds columns that don't already exist."""
    with db.engine.connect() as conn:
        cols = [row[1] for row in conn.exec_driver_sql("PRAGMA table_info(employees)")]
        if "employee_code" not in cols:
            conn.exec_driver_sql("ALTER TABLE employees ADD COLUMN employee_code VARCHAR(50)")
            conn.commit()


@app.cli.command("init-db")
def init_db():
    """Create tables and seed an initial admin user."""
    db.create_all()
    run_light_migrations()
    if not User.query.filter_by(username="admin").first():
        admin = User(username="admin", role="admin")
        admin.set_password("admin123")
        db.session.add(admin)
        db.session.commit()
        print("Created admin user: admin / admin123")
    else:
        print("Database already initialized.")


def ensure_db():
    with app.app_context():
        db.create_all()
        run_light_migrations()
        if not User.query.filter_by(username="admin").first():
            admin = User(username="admin", role="admin")
            admin.set_password("admin123")
            db.session.add(admin)
            db.session.commit()


if __name__ == "__main__":
    ensure_db()
    app.run(debug=True, host="0.0.0.0", port=5000)
