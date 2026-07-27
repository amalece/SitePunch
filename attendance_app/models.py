from datetime import datetime, date
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), default="supervisor")  # admin / supervisor
    active = db.Column(db.Boolean, default=True)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Site(db.Model):
    __tablename__ = "sites"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    location = db.Column(db.String(200))
    active = db.Column(db.Boolean, default=True)

    employees = db.relationship("Employee", backref="site", lazy=True)


class Employee(db.Model):
    __tablename__ = "employees"
    id = db.Column(db.Integer, primary_key=True)
    employee_code = db.Column(db.String(50))
    name = db.Column(db.String(120), nullable=False)
    designation = db.Column(db.String(120))
    site_id = db.Column(db.Integer, db.ForeignKey("sites.id"), nullable=True)
    standard_hours = db.Column(db.Float, default=8.0)
    weekly_off_day = db.Column(db.Integer, default=6)  # 0=Mon ... 6=Sun
    active = db.Column(db.Boolean, default=True)

    attendance_records = db.relationship("Attendance", backref="employee", lazy=True)


class CompanySettings(db.Model):
    __tablename__ = "company_settings"
    id = db.Column(db.Integer, primary_key=True)
    company_name = db.Column(db.String(200), default="")
    address = db.Column(db.String(500), default="")
    logo_filename = db.Column(db.String(200))


class Attendance(db.Model):
    __tablename__ = "attendance"
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employees.id"), nullable=False)
    date = db.Column(db.Date, nullable=False, default=date.today)
    status = db.Column(db.String(20), nullable=False, default="present")
    # present / absent / leave / holiday / weekly_off
    start_time = db.Column(db.String(5))  # "HH:MM"
    end_time = db.Column(db.String(5))
    break_minutes = db.Column(db.Integer, default=60)
    worked_hours = db.Column(db.Float, default=0.0)
    normal_hours = db.Column(db.Float, default=0.0)
    overtime_hours = db.Column(db.Float, default=0.0)
    remarks = db.Column(db.String(255))
    recorded_by = db.Column(db.String(64))
    recorded_at = db.Column(db.DateTime, default=datetime.utcnow)

    __table_args__ = (
        db.UniqueConstraint("employee_id", "date", name="uq_employee_date"),
    )
