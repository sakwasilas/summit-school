# ============================================================
# IMPORTS
# ============================================================
import os
import io
from collections import namedtuple
from datetime import datetime

import pandas as pd
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    flash,
    send_file,
)
from sqlalchemy import or_
from sqlalchemy.orm import joinedload
from werkzeug.utils import secure_filename

from connections import SessionLocal
from models import (
    User,
    Admin,
    HOD,
    Department,
    Course,
    Semester,
    Subject,
    Question,
    Quiz,
    StudentProfile,
    Result,
    Message,
    ActivityLog,
    KnecMark
)
from utils import parse_docx_questions


# ============================================================
# APP CONFIG
# ============================================================
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
app.secret_key = "132silas456sakwa789ayanga"


# ============================================================
# CONTEXT PROCESSOR (HOD)
# ============================================================
@app.context_processor
def inject_hod_context():
    if session.get("role") == "hod":
        db = SessionLocal()
        try:
            dept = db.query(Department).filter_by(id=session.get("department_id")).first()
            return {"department": dept, "hod_name": session.get("username")}
        finally:
            db.close()
    return {}


# ============================================================
# SEMESTER STATUS HELPER
# ============================================================
@app.context_processor
def inject_semester_helpers():
    from datetime import datetime as _dt
    import re

    MONTH_MAP = {
        'jan': 1, 'january': 1, 'feb': 2, 'february': 2,
        'mar': 3, 'march': 3, 'apr': 4, 'april': 4,
        'may': 5, 'jun': 6, 'june': 6, 'jul': 7, 'july': 7,
        'aug': 8, 'august': 8, 'sep': 9, 'sept': 9, 'september': 9,
        'oct': 10, 'october': 10, 'nov': 11, 'november': 11,
        'dec': 12, 'december': 12,
    }

    def parse_semester_months(name):
        if not name:
            return None, None
        parts = re.split(r'\s*(?:–|—|-|/|to|&)\s*', name.lower())
        months = []
        for part in parts:
            word = part.strip().split()[0] if part.strip() else ''
            if word in MONTH_MAP:
                months.append(MONTH_MAP[word])
        if len(months) >= 2:
            return months[0], months[-1]
        elif len(months) == 1:
            return months[0], months[0]
        return None, None

    def semester_time_status(semester):
        now_month = _dt.now().month
        start, end = parse_semester_months(semester.name if semester else '')

        if semester is None:
            return {"state": "pending", "label": "Pending", "short": ""}

        if start is None:
            if getattr(semester, 'is_current', False):
                return {"state": "current", "label": "Current", "short": semester.name}
            return {"state": "pending", "label": "Pending", "short": semester.name}

        if start <= end:
            in_range = start <= now_month <= end
            before = now_month < start
        else:
            in_range = now_month >= start or now_month <= end
            before = end < now_month < start

        if in_range or getattr(semester, 'is_current', False):
            return {"state": "current", "label": "Current", "short": semester.name}
        elif before:
            return {"state": "next", "label": "Next", "short": semester.name}
        else:
            return {"state": "done", "label": "Done", "short": semester.name}

    return {"semester_time_status": semester_time_status}


# ============================================================
# UPLOAD FOLDERS
# ============================================================
BASE_UPLOAD = os.path.join(os.getcwd(), "static", "uploads")
EXAMS_UPLOAD_FOLDER = os.path.join(BASE_UPLOAD, "exams")
QUESTION_IMAGES_FOLDER = os.path.join("static", "question_images")

for path in [BASE_UPLOAD, EXAMS_UPLOAD_FOLDER, QUESTION_IMAGES_FOLDER]:
    os.makedirs(path, exist_ok=True)

ALLOWED_EXAM_EXTENSIONS = {"docx"}


def allowed(filename, allowed_set):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed_set


# ============================================================
# SHARED HELPERS
# ============================================================
def _delete_quiz_tree(db, quiz):
    db.query(Result).filter(Result.quiz_id == quiz.id).delete(synchronize_session=False)
    db.query(Question).filter(Question.quiz_id == quiz.id).delete(synchronize_session=False)
    db.delete(quiz)


def _delete_quizzes_for(db, course_id=None, subject_id=None):
    q = db.query(Quiz)
    if course_id is not None:
        q = q.filter(Quiz.course_id == course_id)
    if subject_id is not None:
        q = q.filter(Quiz.subject_id == subject_id)
    for quiz in q.all():
        _delete_quiz_tree(db, quiz)


def _delete_student_tree(db, student):
    db.query(ActivityLog).filter(ActivityLog.student_id == student.id).delete(synchronize_session=False)
    db.query(Result).filter(Result.student_id == student.id).delete(synchronize_session=False)
    db.query(KnecMark).filter(KnecMark.student_id == student.id).delete(synchronize_session=False)
    user = student.user
    db.delete(student)
    if user:
        db.delete(user)


# ============================================================
# ROLE HELPERS
# ============================================================
def is_admin():
    return session.get("role") == "admin"


def is_hod():
    return session.get("role") == "hod"


def require_admin_or_hod():
    role = session.get("role")
    if role not in ("admin", "hod"):
        return redirect(url_for("login"))
    return None


# ============================================================
# SECTION 1: PUBLIC ROUTES
# ============================================================
@app.route("/")
def home():
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = False
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        db = SessionLocal()
        try:
            admin = db.query(Admin).filter_by(username=username).first()
            if admin and admin.password == password:
                session.update({"username": admin.username, "user_id": admin.id, "role": "admin"})
                return redirect(url_for("admin_dashboard"))

            hod = db.query(HOD).filter_by(username=username).first()
            if hod and hod.password == password:
                session.update({
                    "username": hod.username,
                    "user_id": hod.id,
                    "role": "hod",
                    "department_id": hod.department_id,
                })
                return redirect(url_for("hod_dashboard"))

            user = db.query(User).filter_by(username=username).first()
            if user and user.password == password:
                session.update({"username": user.username, "user_id": user.id, "role": "student"})
                return redirect(url_for("student_dashboard"))

            error = True
        finally:
            db.close()

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"]

        db = SessionLocal()
        try:
            if db.query(User).filter_by(username=username).first():
                flash("❌ Username already exists. Please choose another one.", "danger")
                return render_template("students/Register.html", username=username)

            db.add(User(username=username, password=password))
            db.commit()
            flash("✅ Registration successful! Please log in.", "success")
            return redirect(url_for("login"))
        except Exception:
            db.rollback()
            flash("❌ Something went wrong. Try again.", "danger")
            return render_template("students/Register.html", username=username)
        finally:
            db.close()

    return render_template("students/Register.html")


# ============================================================
# SECTION 2: ADMIN ROUTES
# ============================================================
@app.route("/admin_dashboard")
def admin_dashboard():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        courses = db.query(Course).all()
        students_count = db.query(StudentProfile).count()
        exams_count = db.query(Quiz).count()
        departments_count = db.query(Department).count()

        return render_template(
            "admin/admin_dashboard.html",
            courses=courses,
            username=session.get("username"),
            students_count=students_count,
            exams_count=exams_count,
            departments_count=departments_count,
            current_year=datetime.now().year,
        )
    finally:
        db.close()


# ADMIN | Departments
@app.route("/admin/departments", methods=["GET", "POST"])
def manage_departments():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        if request.method == "POST":
            name = request.form.get("name", "").strip()
            code = request.form.get("code", "").strip() or None

            if not name:
                flash("Department name is required.", "danger")
                return redirect(url_for("manage_departments"))

            if db.query(Department).filter_by(name=name).first():
                flash(f"Department '{name}' already exists.", "warning")
                return redirect(url_for("manage_departments"))

            db.add(Department(name=name, code=code))
            db.commit()
            flash(f"✅ Department '{name}' added.", "success")
            return redirect(url_for("manage_departments"))

        departments = db.query(Department).order_by(Department.name.asc()).all()

        dept_stats = []
        for d in departments:
            dept_stats.append({
                "department": d,
                "courses": db.query(Course).filter_by(department_id=d.id).count(),
                "hods": db.query(HOD).filter_by(department_id=d.id).count(),
            })

        return render_template("admin/manage_departments.html", dept_stats=dept_stats)
    finally:
        db.close()


@app.route("/admin/departments/delete/<int:dept_id>", methods=["POST"])
def delete_department(dept_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        dept = db.query(Department).filter_by(id=dept_id).first()
        if not dept:
            flash("Department not found.", "danger")
            return redirect(url_for("manage_departments"))

        courses_count = db.query(Course).filter_by(department_id=dept_id).count()
        if courses_count:
            flash(f"Cannot delete — {courses_count} course(s) still in this department.", "danger")
            return redirect(url_for("manage_departments"))

        hod_count = db.query(HOD).filter_by(department_id=dept_id).count()
        if hod_count:
            flash("Cannot delete — an HOD is still assigned to this department.", "danger")
            return redirect(url_for("manage_departments"))

        db.delete(dept)
        db.commit()
        flash("✅ Department deleted.", "success")
    finally:
        db.close()

    return redirect(url_for("manage_departments"))


# ADMIN | HODs
@app.route("/admin/hods", methods=["GET", "POST"])
def manage_hods():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            full_name = request.form.get("full_name", "").strip()
            department_id = request.form.get("department_id")

            if not all([username, password, full_name, department_id]):
                flash("All fields are required.", "danger")
                return redirect(url_for("manage_hods"))

            if db.query(HOD).filter_by(username=username).first():
                flash(f"Username '{username}' already exists.", "warning")
                return redirect(url_for("manage_hods"))

            if db.query(HOD).filter_by(department_id=int(department_id)).first():
                flash("That department already has an HOD.", "warning")
                return redirect(url_for("manage_hods"))

            db.add(HOD(
                username=username,
                password=password,
                full_name=full_name,
                department_id=int(department_id),
            ))
            db.commit()
            flash(f"✅ HOD '{username}' added.", "success")
            return redirect(url_for("manage_hods"))

        departments = db.query(Department).order_by(Department.name.asc()).all()
        hods = (
            db.query(HOD)
            .options(joinedload(HOD.department))
            .order_by(HOD.username.asc())
            .all()
        )
        taken = {h.department_id for h in hods}

        return render_template(
            "admin/manage_hods.html",
            departments=departments,
            hods=hods,
            taken_departments=taken,
        )
    finally:
        db.close()


@app.route("/admin/hods/delete/<int:hod_id>", methods=["POST"])
def delete_hod(hod_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        hod = db.query(HOD).filter_by(id=hod_id).first()
        if not hod:
            flash("HOD not found.", "danger")
        else:
            db.delete(hod)
            db.commit()
            flash("✅ HOD deleted.", "success")
    finally:
        db.close()

    return redirect(url_for("manage_hods"))


# ADMIN | Courses
@app.route("/manage_courses", methods=["GET"])
def manage_courses():
    blocked = require_admin_or_hod()
    if blocked:
        return blocked

    db = SessionLocal()
    try:
        if is_hod():
            courses = (
                db.query(Course)
                .filter_by(department_id=session["department_id"])
                .order_by(Course.name.asc())
                .all()
            )
        else:
            courses = db.query(Course).order_by(Course.name.asc()).all()

        return render_template("admin/manage_courses.html", courses=courses)
    finally:
        db.close()


@app.route("/add_course", methods=["GET", "POST"])
def add_course():
    blocked = require_admin_or_hod()
    if blocked:
        return blocked

    db = SessionLocal()
    try:
        if request.method == "POST":
            course_name = request.form["course_name"].strip()
            course_level = request.form["course_level"].strip()

            if is_hod():
                department_id = session["department_id"]
            else:
                department_id = request.form.get("department_id") or None
                department_id = int(department_id) if department_id else None

            if db.query(Course).filter_by(name=course_name).first():
                flash(f"Course '{course_name}' already exists!", "warning")
                return redirect(url_for("add_course"))

            db.add(Course(
                name=course_name,
                level=course_level,
                department_id=department_id,
            ))
            db.commit()
            flash("Course added successfully!", "success")
            return redirect(url_for("manage_courses"))

        if is_hod():
            departments = db.query(Department).filter_by(id=session["department_id"]).all()
        else:
            departments = db.query(Department).order_by(Department.name.asc()).all()

        return render_template("admin/add_course.html", departments=departments)

    except Exception as e:
        db.rollback()
        print("❌ Error in /add_course:", e)
        flash("An error occurred while adding the course.", "danger")
        return redirect(url_for("add_course"))
    finally:
        db.close()


@app.route("/edit_course/<int:course_id>", methods=["GET", "POST"])
def edit_course(course_id):
    blocked = require_admin_or_hod()
    if blocked:
        return blocked

    db = SessionLocal()
    try:
        course = db.query(Course).filter(Course.id == course_id).first()
        if not course:
            return "Course not found", 404

        if is_hod() and course.department_id != session["department_id"]:
            flash("That course is not in your department.", "danger")
            return redirect(url_for("manage_courses"))

        if request.method == "POST":
            course.name = request.form["course_name"].strip()
            course.level = request.form["course_level"].strip()

            if is_hod():
                course.department_id = session["department_id"]
            else:
                dept = request.form.get("department_id") or None
                course.department_id = int(dept) if dept else None

            db.commit()
            flash("Course updated.", "success")
            return redirect(url_for("manage_courses"))

        if is_hod():
            departments = db.query(Department).filter_by(id=session["department_id"]).all()
        else:
            departments = db.query(Department).order_by(Department.name.asc()).all()

        return render_template("admin/edit_course.html", course=course, departments=departments)
    finally:
        db.close()


@app.route("/delete_course/<int:course_id>", methods=["POST"])
def delete_course(course_id):
    blocked = require_admin_or_hod()
    if blocked:
        return blocked

    db = SessionLocal()
    try:
        course = db.query(Course).filter(Course.id == course_id).first()
        if not course:
            return "Course not found", 404

        if is_hod() and course.department_id != session["department_id"]:
            flash("That course is not in your department.", "danger")
            return redirect(url_for("manage_courses"))

        enrolled = db.query(StudentProfile).filter_by(course_id=course_id).count()
        if enrolled > 0:
            flash(f"Cannot delete — {enrolled} student(s) still enrolled.", "danger")
            return redirect(url_for("manage_courses"))

        _delete_quizzes_for(db, course_id=course_id)
        db.query(Message).filter(Message.course_id == course_id).delete(synchronize_session=False)

        for subject in db.query(Subject).filter(Subject.course_id == course_id).all():
            db.delete(subject)

        db.delete(course)
        db.commit()
        flash("Course deleted successfully.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error deleting course: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("manage_courses"))


# ADMIN | Subjects
@app.route("/add_subject", methods=["GET", "POST"])
def add_subject():
    blocked = require_admin_or_hod()
    if blocked:
        return blocked

    db = SessionLocal()
    try:
        if is_hod():
            hod_course_ids = [
                c.id for c in db.query(Course).filter_by(department_id=session["department_id"]).all()
            ]
        else:
            hod_course_ids = None

        if request.method == "POST":
            subject_name = request.form["subject_name"].strip()
            course_id = int(request.form["course_id"])
            semester_id = request.form.get("semester_id") or None
            module = request.form.get("module", "").strip() or None

            if semester_id:
                semester_id = int(semester_id)

            if is_hod() and course_id not in hod_course_ids:
                flash("That course is not in your department.", "danger")
                return redirect(url_for("add_subject"))

            q = db.query(Subject).filter_by(name=subject_name, course_id=course_id)
            if semester_id:
                q = q.filter_by(semester_id=semester_id)
            else:
                q = q.filter(Subject.semester_id.is_(None))
            if q.first():
                flash("Subject already exists for this course/semester!", "danger")
                return redirect(url_for("add_subject"))

            if semester_id:
                sem = db.query(Semester).filter_by(id=semester_id).first()
                if not sem or sem.course_id != course_id:
                    flash("Selected semester does not belong to that course.", "danger")
                    return redirect(url_for("add_subject"))

            db.add(Subject(
                name=subject_name,
                course_id=course_id,
                semester_id=semester_id,
                module=module,
            ))
            db.commit()
            flash("Subject added successfully!", "success")
            return redirect(url_for("manage_courses"))

        if is_hod():
            courses = (
                db.query(Course)
                .filter_by(department_id=session["department_id"])
                .order_by(Course.name.asc())
                .all()
            )
            semesters = (
                db.query(Semester).filter(Semester.course_id.in_(hod_course_ids)).all()
                if hod_course_ids else []
            )
        else:
            courses = db.query(Course).order_by(Course.name.asc()).all()
            semesters = db.query(Semester).all()

        return render_template("admin/add_subject.html", courses=courses, semesters=semesters)
    finally:
        db.close()


@app.route("/edit_subject/<int:subject_id>", methods=["GET", "POST"])
def edit_subject(subject_id):
    blocked = require_admin_or_hod()
    if blocked:
        return blocked

    db = SessionLocal()
    try:
        subject = db.query(Subject).filter(Subject.id == subject_id).first()
        if not subject:
            return "Subject not found", 404

        if is_hod():
            course = db.query(Course).filter_by(
                id=subject.course_id, department_id=session["department_id"]
            ).first()
            if not course:
                flash("That subject is not in your department.", "danger")
                return redirect(url_for("manage_courses"))

        if request.method == "POST":
            subject.name = request.form["subject_name"].strip()
            module = request.form.get("module", "").strip() or None
            subject.module = module

            semester_id = request.form.get("semester_id") or None
            if semester_id:
                semester_id = int(semester_id)
                sem = db.query(Semester).filter_by(id=semester_id).first()
                if not sem or sem.course_id != subject.course_id:
                    flash("Selected semester does not belong to that course.", "danger")
                    return redirect(url_for("edit_subject", subject_id=subject_id))
            subject.semester_id = semester_id
            db.commit()
            flash("Subject updated.", "success")
            return redirect(url_for("manage_courses"))

        semesters = (
            db.query(Semester)
            .filter_by(course_id=subject.course_id)
            .order_by(Semester.created_at.desc())
            .all()
        )
        return render_template("admin/edit_subject.html", subject=subject, semesters=semesters)
    finally:
        db.close()


@app.route("/delete_subject/<int:subject_id>", methods=["POST"])
def delete_subject(subject_id):
    blocked = require_admin_or_hod()
    if blocked:
        return blocked

    db = SessionLocal()
    try:
        subject = db.query(Subject).filter(Subject.id == subject_id).first()
        if not subject:
            return "Subject not found", 404

        if is_hod():
            course = db.query(Course).filter_by(
                id=subject.course_id, department_id=session["department_id"]
            ).first()
            if not course:
                flash("That subject is not in your department.", "danger")
                return redirect(url_for("manage_courses"))

        _delete_quizzes_for(db, subject_id=subject_id)
        db.query(Message).filter(Message.subject_id == subject_id).delete(synchronize_session=False)
        db.delete(subject)
        db.commit()
        flash("Subject deleted successfully.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error deleting subject: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("manage_courses"))


# ADMIN | Semesters
@app.route("/admin/semesters", methods=["GET"])
def manage_semesters():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        courses = db.query(Course).order_by(Course.name.asc()).all()
        semesters = (
            db.query(Semester)
            .options(joinedload(Semester.course))
            .order_by(Semester.created_at.desc())
            .all()
        )
        return render_template("admin/manage_semesters.html", courses=courses, semesters=semesters)
    finally:
        db.close()


@app.route("/admin/semesters/add", methods=["POST"])
def add_semester():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    name = request.form.get("name", "").strip()
    course_id = request.form.get("course_id")

    if not name or not course_id:
        flash("Semester name and course are required.", "danger")
        return redirect(url_for("manage_semesters"))

    db = SessionLocal()
    try:
        if db.query(Semester).filter_by(name=name, course_id=int(course_id)).first():
            flash(f"Semester '{name}' already exists for that course.", "warning")
            return redirect(url_for("manage_semesters"))

        db.add(Semester(name=name, course_id=int(course_id)))
        db.commit()
        flash("Semester added.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {e}", "danger")
    finally:
        db.close()

    return redirect(url_for("manage_semesters"))


@app.route("/admin/semesters/set_current/<int:semester_id>", methods=["POST"])
def set_current_semester(semester_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        sem = db.query(Semester).filter_by(id=semester_id).first()
        if not sem:
            flash("Semester not found.", "danger")
            return redirect(url_for("manage_semesters"))

        db.query(Semester).filter(
            Semester.course_id == sem.course_id,
            Semester.id != sem.id,
        ).update({"is_current": False}, synchronize_session=False)

        sem.is_current = True
        db.commit()
        flash(f"'{sem.name}' is now current for its course.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {e}", "danger")
    finally:
        db.close()

    return redirect(url_for("manage_semesters"))


@app.route("/admin/semesters/delete/<int:semester_id>", methods=["POST"])
def delete_semester(semester_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        sem = db.query(Semester).filter_by(id=semester_id).first()
        if not sem:
            flash("Semester not found.", "danger")
            return redirect(url_for("manage_semesters"))

        if db.query(Subject).filter_by(semester_id=semester_id).count():
            flash("Cannot delete — subjects are still in this semester.", "danger")
            return redirect(url_for("manage_semesters"))

        db.delete(sem)
        db.commit()
        flash("Semester deleted.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {e}", "danger")
    finally:
        db.close()

    return redirect(url_for("manage_semesters"))


# ADMIN | Upload Exam
@app.route("/upload_exam", methods=["GET", "POST"])
def upload_exam():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        courses, subjects = db.query(Course).all(), db.query(Subject).all()

        if request.method == "POST":
            title = request.form["title"].strip()
            course_id, subject_id = int(request.form["course"]), int(request.form["subject"])
            duration = int(request.form.get("duration", 30))
            file = request.files.get("quiz_file")

            if not file or not allowed(file.filename, ALLOWED_EXAM_EXTENSIONS):
                flash("❌ Upload a valid .docx file.", "danger")
                return redirect(request.url)

            filename = secure_filename(file.filename)
            file_path = os.path.join(EXAMS_UPLOAD_FOLDER, filename)
            file.save(file_path)

            questions = parse_docx_questions(file_path, image_output_dir=QUESTION_IMAGES_FOLDER)
            if not questions:
                flash("❌ No valid questions found.", "danger")
                return redirect(request.url)

            quiz = Quiz(title=title, course_id=course_id, subject_id=subject_id,
                        duration=duration, status="active")
            db.add(quiz)
            db.commit()
            db.refresh(quiz)

            for q in questions:
                db.add(Question(
                    quiz_id=quiz.id,
                    question_text=q.get("question", ""),
                    option_a=q.get("a", ""), option_b=q.get("b", ""),
                    option_c=q.get("c", ""), option_d=q.get("d", ""),
                    correct_option=q.get("answer", "").lower(),
                    marks=q.get("marks", 1),
                    extra_content=q.get("extra_content"),
                    image=q.get("image"),
                ))
            db.commit()
            flash(f"✅ Uploaded quiz with {len(questions)} question(s).", "success")

            return render_template(
                "admin/upload_exams.html",
                courses=courses, subjects=subjects,
                uploaded_quiz_id=quiz.id,
            )

        return render_template("admin/upload_exams.html", courses=courses, subjects=subjects)
    finally:
        db.close()


# ADMIN | Quizzes
@app.route("/admin/manage_quizzes", methods=["GET", "POST"])
def manage_quizzes():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        if request.method == "POST":
            quiz_id = int(request.form.get("quiz_id"))
            action = request.form.get("action")

            quiz = db.query(Quiz).filter_by(id=quiz_id).first()
            if not quiz:
                flash("Quiz not found.", "danger")
                return redirect(url_for("manage_quizzes"))

            if action == "activate":
                quiz.status = "active"
            elif action == "deactivate":
                quiz.status = "inactive"

            db.commit()
            flash(f"Quiz '{quiz.title}' has been {quiz.status}.", "success")
            return redirect(url_for("manage_quizzes"))

        quizzes = db.query(Quiz).all()
        return render_template("admin/manage_quizzes.html", quizzes=quizzes)
    finally:
        db.close()


@app.route("/admin/delete_quiz/<int:quiz_id>", methods=["POST"])
def delete_quiz(quiz_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        quiz = db.query(Quiz).filter_by(id=quiz_id).first()
        if quiz:
            _delete_quiz_tree(db, quiz)
            db.commit()
            flash("Quiz deleted successfully.", "success")
        else:
            flash("Quiz not found.", "danger")
    except Exception as e:
        db.rollback()
        flash(f"Error deleting quiz: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("manage_quizzes"))


@app.route("/admin/review_quiz/<int:quiz_id>", methods=["GET", "POST"])
def review_uploaded_quiz(quiz_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        quiz = db.query(Quiz).filter_by(id=quiz_id).first()
        if not quiz:
            flash("Quiz not found.", "danger")
            return redirect(url_for("upload_exam"))

        if request.method == "POST":
            action = request.form.get("action")
            if action == "delete":
                _delete_quiz_tree(db, quiz)
                db.commit()
                flash("❌ Quiz deleted.", "warning")
                return redirect(url_for("upload_exam"))
            elif action == "confirm":
                flash("✅ Quiz confirmed and saved.", "success")
                return redirect(url_for("upload_exam"))

        questions = db.query(Question).filter_by(quiz_id=quiz.id).all()
        return render_template("admin/review_uploaded_quiz.html", quiz=quiz, questions=questions)
    finally:
        db.close()


# ADMIN | Manage Students
@app.route("/manage_students", methods=["GET", "POST"])
def manage_students():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()

    if request.method == "POST":
        student_id = request.form.get("student_id")
        action = request.form.get("action")
        student = db.query(StudentProfile).get(student_id)

        if action == "toggle_block" and student:
            student.blocked = not student.blocked
            db.commit()
        elif action == "delete" and student:
            try:
                _delete_student_tree(db, student)
                db.commit()
                flash("✅ Student deleted.", "success")
            except Exception as e:
                db.rollback()
                flash(f"Error deleting student: {str(e)}", "danger")

        return redirect(url_for("manage_students"))

    search_query = request.args.get("search", "").strip()

    if search_query:
        students = (
            db.query(StudentProfile)
            .options(joinedload(StudentProfile.course))
            .filter(or_(
                StudentProfile.full_name.ilike(f"%{search_query}%"),
                StudentProfile.exam_type.ilike(f"%{search_query}%"),
                StudentProfile.admission_number.ilike(f"%{search_query}%"),
                StudentProfile.phone_number.ilike(f"%{search_query}%"),
            ))
            .all()
        )
    else:
        students = (
            db.query(StudentProfile)
            .options(joinedload(StudentProfile.course))
            .all()
        )

    return render_template("admin/manage_students.html", students=students, search_query=search_query)


@app.route("/show_credentials")
def show_credentials():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        users = db.query(User).options(joinedload(User.profile)).all()
    finally:
        db.close()

    return render_template("admin/show_credentials.html", users=users)


@app.route("/delete_user/<int:user_id>", methods=["POST"])
def delete_user(user_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        user = db.query(User).filter(User.id == user_id).first()
        if user:
            profile = user.profile
            if profile:
                _delete_student_tree(db, profile)
            else:
                db.delete(user)
            db.commit()
            flash("✅ User deleted.", "success")
        else:
            flash("User not found.", "danger")
    except Exception as e:
        db.rollback()
        flash(f"Error deleting user: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("show_credentials"))


# ADMIN | View Results
@app.route("/admin/view_results", methods=["GET", "POST"])
def view_results():
    if "user_id" not in session or session.get("role") != "admin":
        flash("Admin access required", "danger")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        courses = db.query(Course).all()
        subjects = db.query(Subject).all()

        selected_course = request.form.get("course")
        selected_subject = request.form.get("subject")
        export = request.form.get("export")

        query = (
            db.query(Result)
            .options(
                joinedload(Result.student),
                joinedload(Result.quiz).joinedload(Quiz.course),
                joinedload(Result.quiz).joinedload(Quiz.subject),
            )
            .join(Result.quiz)
            .join(Quiz.course)
            .join(Quiz.subject)
            .join(Result.student)
        )

        if selected_course:
            query = query.filter(Quiz.course_id == int(selected_course))
        if selected_subject:
            query = query.filter(Quiz.subject_id == int(selected_subject))

        results = query.all()

        if export == "true":
            data = []
            for r in results:
                data.append({
                    "Student Username": r.student.user.username if r.student and r.student.user else "N/A",
                    "Full Name": r.student.full_name if r.student else "N/A",
                    "Course": r.quiz.course.name if r.quiz.course else "N/A",
                    "Subject": r.quiz.subject.name if r.quiz.subject else "N/A",
                    "Quiz Title": r.quiz.title,
                    "Score": r.score, "Total Marks": r.total_marks,
                    "Percentage": r.percentage,
                    "Taken On": r.taken_on.strftime("%Y-%m-%d %H:%M:%S"),
                })

            df = pd.DataFrame(data)
            excel_path = os.path.join(EXAMS_UPLOAD_FOLDER, "quiz_results.xlsx")
            df.to_excel(excel_path, index=False)

            return send_file(
                excel_path,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True, download_name="quiz_results.xlsx",
            )

        return render_template(
            "admin/view_results.html",
            results=results, courses=courses, subjects=subjects,
            selected_course=selected_course, selected_subject=selected_subject,
        )
    finally:
        db.close()


# ADMIN | Messages
@app.route("/admin/messages", methods=["GET", "POST"])
def admin_messages():
    if session.get("role") != "admin":
        flash("Please log in as admin.", "error")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        if request.method == "POST":
            content = request.form["content"]
            target_type = request.form["target_type"]
            course_id = request.form.get("course_id") or None
            subject_id = request.form.get("subject_id") or None

            db.add(Message(
                content=content, target_type=target_type,
                course_id=course_id if target_type == "course" else None,
                subject_id=subject_id if target_type == "subject" else None,
            ))
            db.commit()
            flash("Message created successfully!", "success")
            return redirect(url_for("admin_messages"))

        courses = db.query(Course).all()
        subjects = db.query(Subject).all()
        messages = db.query(Message).order_by(Message.created_at.desc()).all()

        return render_template("admin/messages.html", courses=courses, subjects=subjects, messages=messages)
    finally:
        db.close()


@app.route("/admin/messages/delete/<int:message_id>", methods=["POST"])
def delete_admin_message(message_id):
    if session.get("role") != "admin":
        flash("Admin access required.", "danger")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        message = db.query(Message).filter_by(id=message_id).first()
        if not message:
            flash("Message not found.", "danger")
        else:
            db.delete(message)
            db.commit()
            flash("Message deleted successfully.", "success")
    finally:
        db.close()

    return redirect(url_for("admin_messages"))


# ADMIN | Student Activity
@app.route("/admin/student_activity")
def student_activity():
    if session.get("role") != "admin":
        flash("Please log in as admin.", "error")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        total_students = db.query(StudentProfile).count()

        total_doing_exam = (
            db.query(ActivityLog)
            .filter(ActivityLog.activity_type == "exam", ActivityLog.is_active == True)
            .count()
        )

        StudentActivity = namedtuple("StudentActivity", ["full_name", "course_name", "activity_type"])

        rows = (
            db.query(StudentProfile.full_name, Course.name.label("course_name"), ActivityLog.activity_type)
            .join(ActivityLog, ActivityLog.student_id == StudentProfile.id)
            .join(Course, StudentProfile.course_id == Course.id)
            .filter(ActivityLog.is_active == True, ActivityLog.activity_type == "exam")
            .all()
        )

        active_students = [StudentActivity(*row) for row in rows]

        return render_template(
            "admin/student_activity.html",
            total_students=total_students,
            total_doing_exam=total_doing_exam,
            active_students=active_students,
        )
    finally:
        db.close()


# ADMIN | Enter Marks
@app.route("/admin/marks", methods=["GET"])
def admin_marks_picker():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        departments = db.query(Department).order_by(Department.name.asc()).all()

        dept_cards = []
        for d in departments:
            course_ids = [c.id for c in db.query(Course).filter_by(department_id=d.id).all()]
            students_count = 0
            if course_ids:
                students_count = (
                    db.query(StudentProfile)
                    .filter(StudentProfile.course_id.in_(course_ids))
                    .count()
                )
            dept_cards.append({
                "department": d,
                "courses_count": len(course_ids),
                "students_count": students_count,
            })

        return render_template("admin/marks_step1_departments.html", dept_cards=dept_cards)
    finally:
        db.close()


@app.route("/admin/marks/department/<int:dept_id>", methods=["GET"])
def admin_marks_department(dept_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=dept_id).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("admin_marks_picker"))

        courses = (
            db.query(Course)
            .filter_by(department_id=dept_id)
            .order_by(Course.name.asc())
            .all()
        )

        course_id = request.args.get("course_id", type=int)
        semester_id = request.args.get("semester_id", type=int)

        course = None
        semester = None
        subjects = []
        semesters = []

        if course_id:
            course = db.query(Course).filter_by(id=course_id, department_id=dept_id).first()
        if course:
            semesters = (
                db.query(Semester)
                .filter_by(course_id=course.id)
                .order_by(Semester.created_at.desc())
                .all()
            )
            if semester_id:
                semester = db.query(Semester).filter_by(id=semester_id, course_id=course.id).first()
            elif semesters:
                semester = next((s for s in semesters if s.is_current), semesters[0])

            if semester:
                subjects = (
                    db.query(Subject)
                    .filter_by(course_id=course.id, semester_id=semester.id)
                    .order_by(Subject.name.asc())
                    .all()
                )

        return render_template(
            "admin/marks_step2_course.html",
            department=department,
            courses=courses,
            course=course,
            semesters=semesters,
            semester=semester,
            subjects=subjects,
        )
    finally:
        db.close()


@app.route("/admin/marks/enter", methods=["GET", "POST"])
def admin_enter_marks():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        src = request.form if request.method == "POST" else request.args
        subject_id = src.get("subject_id", type=int)
        semester_id = src.get("semester_id", type=int)

        if not subject_id or not semester_id:
            flash("Pick a subject and semester first.", "danger")
            return redirect(url_for("admin_marks_picker"))

        subject = db.query(Subject).filter_by(id=subject_id).first()
        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not subject or not semester:
            flash("Subject or semester not found.", "danger")
            return redirect(url_for("admin_marks_picker"))

        if subject.course_id != semester.course_id:
            flash("Subject and semester belong to different courses.", "danger")
            return redirect(url_for("admin_marks_picker"))

        course = db.query(Course).filter_by(id=subject.course_id).first()
        department = course.department if course else None

        students = (
            db.query(StudentProfile)
            .filter_by(course_id=subject.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )

        if request.method == "POST":
            def to_int_or_none(v):
                v = (v or "").strip()
                return int(v) if v.isdigit() else None

            saved = 0
            for s in students:
                c1 = to_int_or_none(request.form.get(f"cat1_{s.id}"))
                c2 = to_int_or_none(request.form.get(f"cat2_{s.id}"))
                fn = to_int_or_none(request.form.get(f"final_{s.id}"))

                existing = (
                    db.query(KnecMark)
                    .filter_by(student_id=s.id, subject_id=subject_id, semester_id=semester_id)
                    .first()
                )

                if c1 is None and c2 is None and fn is None:
                    if existing:
                        db.delete(existing)
                    continue

                if c1 is not None and not (0 <= c1 <= 15):
                    flash(f"CAT1 for {s.full_name} must be 0–15.", "warning")
                    continue
                if c2 is not None and not (0 <= c2 <= 15):
                    flash(f"CAT2 for {s.full_name} must be 0–15.", "warning")
                    continue
                if fn is not None and not (0 <= fn <= 70):
                    flash(f"Final for {s.full_name} must be 0–70.", "warning")
                    continue

                if existing:
                    existing.cat1 = c1
                    existing.cat2 = c2
                    existing.final = fn
                    existing.entered_at = datetime.utcnow()
                else:
                    db.add(KnecMark(
                        student_id=s.id, subject_id=subject_id,
                        semester_id=semester_id,
                        cat1=c1, cat2=c2, final=fn,
                    ))
                saved += 1

            db.commit()
            flash(f"✅ Saved marks for {saved} student(s).", "success")

            if department:
                return redirect(url_for(
                    "admin_marks_department",
                    dept_id=department.id,
                    course_id=subject.course_id,
                    semester_id=semester_id,
                ))
            return redirect(url_for("admin_marks_picker"))

        existing = {
            m.student_id: m
            for m in db.query(KnecMark).filter_by(subject_id=subject_id, semester_id=semester_id).all()
        }

        return render_template(
            "admin/marks_entry.html",
            subject=subject, semester=semester,
            course=course, department=department,
            students=students, existing=existing,
        )
    finally:
        db.close()


@app.route("/admin/marks/class", methods=["GET"])
def admin_class_marks():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    semester_id = request.args.get("semester_id", type=int)
    if not semester_id:
        flash("Pick a semester.", "warning")
        return redirect(url_for("admin_marks_picker"))

    db = SessionLocal()
    try:
        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not semester:
            return "Semester not found", 404

        students = (
            db.query(StudentProfile)
            .filter_by(course_id=semester.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )
        subjects = (
            db.query(Subject)
            .filter_by(course_id=semester.course_id, semester_id=semester_id)
            .order_by(Subject.name.asc())
            .all()
        )

        marks = {}
        for m in db.query(KnecMark).filter_by(semester_id=semester_id).all():
            marks.setdefault(m.student_id, {})[m.subject_id] = m

        return render_template(
            "admin/class_marks.html",
            semester=semester, students=students, subjects=subjects, marks=marks,
        )
    finally:
        db.close()


@app.route("/admin/marks/class/export")
def admin_class_marks_export():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    semester_id = request.args.get("semester_id", type=int)
    if not semester_id:
        flash("Pick a semester first.", "warning")
        return redirect(url_for("admin_marks_picker"))

    db = SessionLocal()
    try:
        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not semester:
            return "Semester not found", 404

        course = db.query(Course).filter_by(id=semester.course_id).first()
        students = (
            db.query(StudentProfile)
            .filter_by(course_id=semester.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )
        subjects = (
            db.query(Subject)
            .filter_by(course_id=semester.course_id, semester_id=semester_id)
            .order_by(Subject.name.asc())
            .all()
        )

        marks = {}
        for m in db.query(KnecMark).filter_by(semester_id=semester_id).all():
            marks.setdefault(m.student_id, {})[m.subject_id] = m

        data = []
        for s in students:
            row = {
                "Admission No": s.admission_number,
                "Student Name": s.full_name,
                "Exam Type": (s.exam_type or "").upper(),
            }
            grand_total = 0
            counted = 0
            for subj in subjects:
                m = marks.get(s.id, {}).get(subj.id)
                c1 = m.cat1 if m and m.cat1 is not None else ""
                c2 = m.cat2 if m and m.cat2 is not None else ""
                fn = m.final if m and m.final is not None else ""
                total = (m.cat1 or 0) + (m.cat2 or 0) + (m.final or 0) if m else ""

                prefix = subj.name.split(" ", 1)[0]
                row[f"{prefix} CAT1"] = c1
                row[f"{prefix} CAT2"] = c2
                row[f"{prefix} Final"] = fn
                row[f"{prefix} Total"] = total

                if isinstance(total, int):
                    grand_total += total
                    counted += 1

            row["Grand Total"] = grand_total
            row["Average"] = round(grand_total / counted, 2) if counted else ""
            data.append(row)

        df = pd.DataFrame(data)
        identity_cols = ["Admission No", "Student Name", "Exam Type"]
        total_cols = ["Grand Total", "Average"]
        subject_cols = [c for c in df.columns if c not in identity_cols + total_cols]
        df = df[identity_cols + subject_cols + total_cols]

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Marks")
        buf.seek(0)

        safe_course = (course.name if course else "Class").replace(" ", "_")[:30]
        safe_sem = semester.name.replace(" ", "_")

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"Marks_{safe_course}_{safe_sem}.xlsx",
        )
    finally:
        db.close()


# ADMIN | Reports
@app.route("/admin/reports", methods=["GET"])
def admin_reports_picker():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        courses = db.query(Course).order_by(Course.name.asc()).all()
        semesters = db.query(Semester).order_by(Semester.course_id.asc(), Semester.created_at.desc()).all()
        return render_template("admin/reports_picker.html", courses=courses, semesters=semesters)
    finally:
        db.close()


@app.route("/admin/reports/class", methods=["GET"])
def admin_reports_class():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    semester_id = request.args.get("semester_id", type=int)
    if not semester_id:
        flash("Pick a semester first.", "warning")
        return redirect(url_for("admin_reports_picker"))

    db = SessionLocal()
    try:
        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not semester:
            return "Semester not found", 404

        students = (
            db.query(StudentProfile)
            .filter_by(course_id=semester.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )

        student_summary = []
        for s in students:
            is_knec = (s.exam_type or "").strip().lower() == "knec"
            total_sum = 0
            subjects_count = 0

            if is_knec:
                for m in db.query(KnecMark).filter_by(student_id=s.id, semester_id=semester_id).all():
                    total_sum += (m.cat1 or 0) + (m.cat2 or 0) + (m.final or 0)
                    subjects_count += 1

            avg = round(total_sum / subjects_count, 2) if subjects_count else None
            grade = None
            if avg is not None:
                grade = 1 if avg >= 84 else 2 if avg >= 75 else 3 if avg >= 66 else 4 if avg >= 60 else 5 if avg >= 50 else 6 if avg >= 40 else 7

            result = None
            if avg is not None:
                result = "DISTINCTION" if avg >= 75 else "CREDIT" if avg >= 60 else "PASS" if avg >= 40 else "REFER"

            student_summary.append({
                "student": s, "is_knec": is_knec, "avg": avg,
                "overall_grade": grade, "overall_result": result,
                "subjects_count": subjects_count,
            })

        return render_template(
            "admin/reports_class.html",
            semester=semester, student_summary=student_summary,
            total_students=len(students),
        )
    finally:
        db.close()


@app.route("/admin/reports/class/export")
def admin_reports_class_export():
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    semester_id = request.args.get("semester_id", type=int)
    if not semester_id:
        flash("Pick a semester first.", "warning")
        return redirect(url_for("admin_reports_picker"))

    db = SessionLocal()
    try:
        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not semester:
            return "Semester not found", 404

        course = db.query(Course).filter_by(id=semester.course_id).first()
        students = (
            db.query(StudentProfile)
            .filter_by(course_id=semester.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )

        data = []
        for s in students:
            is_knec = (s.exam_type or "").strip().lower() == "knec"
            total_sum = 0
            subjects_count = 0
            if is_knec:
                for m in db.query(KnecMark).filter_by(student_id=s.id, semester_id=semester_id).all():
                    total_sum += (m.cat1 or 0) + (m.cat2 or 0) + (m.final or 0)
                    subjects_count += 1

            avg = round(total_sum / subjects_count, 2) if subjects_count else ""
            grade = ""
            if isinstance(avg, float):
                grade = 1 if avg >= 84 else 2 if avg >= 75 else 3 if avg >= 66 else 4 if avg >= 60 else 5 if avg >= 50 else 6 if avg >= 40 else 7

            result = ""
            if isinstance(avg, float):
                result = "DISTINCTION" if avg >= 75 else "CREDIT" if avg >= 60 else "PASS" if avg >= 40 else "REFER"

            data.append({
                "Admission No": s.admission_number,
                "Student Name": s.full_name,
                "Course": course.name if course else "",
                "Semester": semester.name,
                "Exam Type": (s.exam_type or "").upper(),
                "Subjects Graded": subjects_count,
                "Total Marks": total_sum,
                "Average": avg,
                "Grade": grade,
                "Overall Result": result,
            })

        df = pd.DataFrame(data)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Class Report")
        buf.seek(0)

        safe_course = (course.name if course else "Class").replace(" ", "_")[:30]
        safe_sem = semester.name.replace(" ", "_")

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"ClassReport_{safe_course}_{safe_sem}.xlsx",
        )
    finally:
        db.close()


@app.route("/admin/reports/student/<int:student_id>/<int:semester_id>")
def admin_report_student(student_id, semester_id):
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        student = db.query(StudentProfile).filter_by(id=student_id).first()
        if not student:
            return "Student not found", 404

        sem = db.query(Semester).filter_by(id=semester_id, course_id=student.course_id).first()
        if not sem:
            return "Semester not found for this student", 404

        subjects = (
            db.query(Subject)
            .filter_by(course_id=student.course_id, semester_id=sem.id)
            .order_by(Subject.name.asc())
            .all()
        )

        rows = []
        for subj in subjects:
            m = db.query(KnecMark).filter_by(
                student_id=student.id, subject_id=subj.id, semester_id=sem.id
            ).first()
            c1 = m.cat1 if m and m.cat1 is not None else None
            c2 = m.cat2 if m and m.cat2 is not None else None
            fn = m.final if m and m.final is not None else None
            rows.append({
                "subject": subj, "cat1": c1, "cat2": c2, "final": fn,
                "total": (c1 or 0) + (c2 or 0) + (fn or 0),
            })

        return render_template(
            "students/print_report.html",
            student=student, semester=sem, rows=rows,
            current_year=datetime.now().year, admin_mode=True,
        )
    finally:
        db.close()


# ============================================================
# SECTION 3: HOD ROUTES
# ============================================================

@app.route("/hod/dashboard")
def hod_dashboard():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=session["department_id"]).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("logout"))

        courses = (
            db.query(Course)
            .filter_by(department_id=department.id)
            .order_by(Course.name.asc())
            .all()
        )
        course_ids = [c.id for c in courses]

        students_count = 0
        subjects_count = 0
        semesters_count = 0
        if course_ids:
            students_count = db.query(StudentProfile).filter(StudentProfile.course_id.in_(course_ids)).count()
            subjects_count = db.query(Subject).filter(Subject.course_id.in_(course_ids)).count()
            semesters_count = db.query(Semester).filter(Semester.course_id.in_(course_ids)).count()

        course_cards = []
        for c in courses:
            current_sem = db.query(Semester).filter_by(course_id=c.id, is_current=True).first()
            c_students = db.query(StudentProfile).filter_by(course_id=c.id).count()
            c_subjects = db.query(Subject).filter_by(course_id=c.id).count()
            course_cards.append({
                "course": c,
                "current_semester": current_sem,
                "students_count": c_students,
                "subjects_count": c_subjects,
            })

        return render_template(
            "hod/dashboard.html",
            hod_name=session.get("username"),
            department=department,
            course_cards=course_cards,
            students_count=students_count,
            subjects_count=subjects_count,
            semesters_count=semesters_count,
            current_year=datetime.now().year,
        )
    finally:
        db.close()


# HOD | Manage Courses
@app.route("/hod/courses", methods=["GET"])
def hod_manage_courses():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=department_id).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("logout"))

        courses = (
            db.query(Course)
            .filter_by(department_id=department_id)
            .order_by(Course.name.asc())
            .all()
        )

        for course in courses:
            course.students_count = db.query(StudentProfile).filter_by(course_id=course.id).count()
            course.subjects_count = db.query(Subject).filter_by(course_id=course.id).count()

        return render_template(
            "hod/manage_courses.html",
            department=department,
            courses=courses,
            hod_name=session.get("username"),
        )
    finally:
        db.close()


@app.route("/hod/courses/add", methods=["POST"])
def hod_add_course():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    course_name = request.form.get("course_name", "").strip()
    course_level = request.form.get("course_level", "").strip()

    if not course_name or not course_level:
        flash("Course name and level are required.", "danger")
        return redirect(url_for("hod_manage_courses"))

    db = SessionLocal()
    try:
        existing = db.query(Course).filter_by(name=course_name).first()
        if existing:
            flash(f"Course '{course_name}' already exists.", "warning")
            return redirect(url_for("hod_manage_courses"))

        db.add(Course(name=course_name, level=course_level, department_id=department_id))
        db.commit()
        flash(f"✅ Course '{course_name}' added.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_manage_courses"))


@app.route("/hod/courses/edit/<int:course_id>", methods=["GET", "POST"])
def hod_edit_course(course_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        course = db.query(Course).filter_by(id=course_id, department_id=department_id).first()
        if not course:
            flash("Course not found in your department.", "danger")
            return redirect(url_for("hod_manage_courses"))

        if request.method == "POST":
            name = request.form.get("course_name", "").strip()
            level = request.form.get("course_level", "").strip()

            if not name or not level:
                flash("Course name and level are required.", "danger")
                return redirect(url_for("hod_edit_course", course_id=course_id))

            dup = db.query(Course).filter(Course.name == name, Course.id != course_id).first()
            if dup:
                flash(f"Course '{name}' already exists.", "warning")
                return redirect(url_for("hod_edit_course", course_id=course_id))

            course.name = name
            course.level = level
            db.commit()
            flash("✅ Course updated.", "success")
            return redirect(url_for("hod_manage_courses"))

        department = db.query(Department).filter_by(id=department_id).first()
        return render_template(
            "hod/edit_course.html",
            course=course,
            department=department,
            hod_name=session.get("username"),
        )
    finally:
        db.close()


@app.route("/hod/courses/delete/<int:course_id>", methods=["POST"])
def hod_delete_course(course_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        course = db.query(Course).filter_by(id=course_id, department_id=department_id).first()
        if not course:
            flash("Course not found in your department.", "danger")
            return redirect(url_for("hod_manage_courses"))

        enrolled = db.query(StudentProfile).filter_by(course_id=course_id).count()
        if enrolled > 0:
            flash(f"Cannot delete — {enrolled} student(s) still enrolled.", "danger")
            return redirect(url_for("hod_manage_courses"))

        _delete_quizzes_for(db, course_id=course_id)
        db.query(Message).filter(Message.course_id == course_id).delete(synchronize_session=False)

        for subject in db.query(Subject).filter(Subject.course_id == course_id).all():
            db.delete(subject)

        db.delete(course)
        db.commit()
        flash("✅ Course deleted.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_manage_courses"))


# HOD | Manage Subjects
@app.route("/hod/subjects", methods=["GET"])
def hod_manage_subjects():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=department_id).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("logout"))

        courses = (
            db.query(Course)
            .filter_by(department_id=department_id)
            .order_by(Course.name.asc())
            .all()
        )
        course_ids = [c.id for c in courses]

        selected_course_id = request.args.get("course_id", type=int)

        subjects_query = (
            db.query(Subject)
            .options(joinedload(Subject.course), joinedload(Subject.semester))
            .filter(Subject.course_id.in_(course_ids))
        )
        if selected_course_id and selected_course_id in course_ids:
            subjects_query = subjects_query.filter(Subject.course_id == selected_course_id)

        subjects = subjects_query.order_by(Subject.name.asc()).all()

        semesters = []
        if course_ids:
            semesters = (
                db.query(Semester)
                .options(joinedload(Semester.course))
                .filter(Semester.course_id.in_(course_ids))
                .order_by(Semester.name.asc())
                .all()
            )

        return render_template(
            "hod/manage_subjects.html",
            department=department,
            courses=courses,
            subjects=subjects,
            semesters=semesters,
            selected_course_id=selected_course_id,
            hod_name=session.get("username"),
        )
    finally:
        db.close()


@app.route("/hod/subjects/add", methods=["POST"])
def hod_add_subject():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    subject_name = request.form.get("subject_name", "").strip()
    course_id = request.form.get("course_id", type=int)
    semester_id = request.form.get("semester_id", type=int)
    module = request.form.get("module", "").strip() or None

    if not subject_name or not course_id:
        flash("Subject name and course are required.", "danger")
        return redirect(url_for("hod_manage_subjects"))

    db = SessionLocal()
    try:
        course = db.query(Course).filter_by(id=course_id, department_id=department_id).first()
        if not course:
            flash("Course not found in your department.", "danger")
            return redirect(url_for("hod_manage_subjects"))

        if semester_id:
            sem = db.query(Semester).filter_by(id=semester_id).first()
            if not sem or sem.course_id != course_id:
                flash("Selected semester doesn't belong to that course.", "danger")
                return redirect(url_for("hod_manage_subjects"))

        q = db.query(Subject).filter_by(name=subject_name, course_id=course_id)
        if semester_id:
            q = q.filter_by(semester_id=semester_id)
        else:
            q = q.filter(Subject.semester_id.is_(None))

        if q.first():
            flash("Subject already exists for this course/semester.", "warning")
            return redirect(url_for("hod_manage_subjects"))

        db.add(Subject(name=subject_name, course_id=course_id,
                       semester_id=semester_id, module=module))
        db.commit()
        flash(f"✅ Subject '{subject_name}' added.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_manage_subjects"))


@app.route("/hod/subjects/edit/<int:subject_id>", methods=["GET", "POST"])
def hod_edit_subject(subject_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        subject = db.query(Subject).filter_by(id=subject_id).first()
        if not subject:
            flash("Subject not found.", "danger")
            return redirect(url_for("hod_manage_subjects"))

        course = db.query(Course).filter_by(id=subject.course_id).first()
        if not course or course.department_id != department_id:
            flash("Subject not in your department.", "danger")
            return redirect(url_for("hod_manage_subjects"))

        if request.method == "POST":
            name = request.form.get("subject_name", "").strip()
            module = request.form.get("module", "").strip() or None
            semester_id = request.form.get("semester_id", type=int)

            if not name:
                flash("Subject name is required.", "danger")
                return redirect(url_for("hod_edit_subject", subject_id=subject_id))

            if semester_id:
                sem = db.query(Semester).filter_by(id=semester_id).first()
                if not sem or sem.course_id != subject.course_id:
                    flash("Selected semester doesn't belong to that course.", "danger")
                    return redirect(url_for("hod_edit_subject", subject_id=subject_id))

            subject.name = name
            subject.module = module
            subject.semester_id = semester_id
            db.commit()
            flash("✅ Subject updated.", "success")
            return redirect(url_for("hod_manage_subjects"))

        semesters = (
            db.query(Semester)
            .filter_by(course_id=subject.course_id)
            .order_by(Semester.name.asc())
            .all()
        )
        return render_template(
            "hod/edit_subject.html",
            subject=subject,
            semesters=semesters,
            department=db.query(Department).filter_by(id=department_id).first(),
            hod_name=session.get("username"),
        )
    finally:
        db.close()


@app.route("/hod/subjects/delete/<int:subject_id>", methods=["POST"])
def hod_delete_subject(subject_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        subject = db.query(Subject).filter_by(id=subject_id).first()
        if not subject:
            flash("Subject not found.", "danger")
            return redirect(url_for("hod_manage_subjects"))

        course = db.query(Course).filter_by(id=subject.course_id).first()
        if not course or course.department_id != department_id:
            flash("Subject not in your department.", "danger")
            return redirect(url_for("hod_manage_subjects"))

        _delete_quizzes_for(db, subject_id=subject_id)
        db.query(Message).filter(Message.subject_id == subject_id).delete(synchronize_session=False)
        db.query(KnecMark).filter(KnecMark.subject_id == subject_id).delete(synchronize_session=False)

        db.delete(subject)
        db.commit()
        flash("✅ Subject deleted.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_manage_subjects"))


# HOD | Manage Semesters
@app.route("/hod/semesters", methods=["GET"])
def hod_manage_semesters():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=department_id).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("logout"))

        courses = (
            db.query(Course)
            .filter_by(department_id=department_id)
            .order_by(Course.name.asc())
            .all()
        )
        course_ids = [c.id for c in courses]

        semesters = []
        if course_ids:
            semesters = (
                db.query(Semester)
                .options(joinedload(Semester.course))
                .filter(Semester.course_id.in_(course_ids))
                .order_by(Semester.created_at.desc())
                .all()
            )

        # ✅ Count subjects per semester WHILE the session is open
        for sem in semesters:
            sem.subject_count = (
                db.query(Subject)
                .filter_by(semester_id=sem.id)
                .count()
            )

        return render_template(
            "hod/manage_semesters.html",
            department=department,
            courses=courses,
            semesters=semesters,
            hod_name=session.get("username"),
        )
    finally:
        db.close()


@app.route("/hod/semesters/add", methods=["POST"])
def hod_add_semester():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    name = request.form.get("name", "").strip()
    course_id = request.form.get("course_id", type=int)

    if not name or not course_id:
        flash("Semester name and course are required.", "danger")
        return redirect(url_for("hod_manage_semesters"))

    db = SessionLocal()
    try:
        course = db.query(Course).filter_by(id=course_id, department_id=department_id).first()
        if not course:
            flash("Course not in your department.", "danger")
            return redirect(url_for("hod_manage_semesters"))

        if db.query(Semester).filter_by(name=name, course_id=course_id).first():
            flash(f"Semester '{name}' already exists.", "warning")
            return redirect(url_for("hod_manage_semesters"))

        db.add(Semester(name=name, course_id=course_id))
        db.commit()
        flash(f"✅ Semester '{name}' added.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_manage_semesters"))


@app.route("/hod/semesters/set_current/<int:semester_id>", methods=["POST"])
def hod_set_current_semester(semester_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        sem = db.query(Semester).filter_by(id=semester_id).first()
        if not sem:
            flash("Semester not found.", "danger")
            return redirect(url_for("hod_manage_semesters"))

        course = db.query(Course).filter_by(id=sem.course_id).first()
        if not course or course.department_id != department_id:
            flash("Semester not in your department.", "danger")
            return redirect(url_for("hod_manage_semesters"))

        db.query(Semester).filter(
            Semester.course_id == sem.course_id,
            Semester.id != sem.id,
        ).update({"is_current": False}, synchronize_session=False)

        sem.is_current = True
        db.commit()
        flash(f"'{sem.name}' is now current.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {e}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_manage_semesters"))


@app.route("/hod/semesters/delete/<int:semester_id>", methods=["POST"])
def hod_delete_semester(semester_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        sem = db.query(Semester).filter_by(id=semester_id).first()
        if not sem:
            flash("Semester not found.", "danger")
            return redirect(url_for("hod_manage_semesters"))

        course = db.query(Course).filter_by(id=sem.course_id).first()
        if not course or course.department_id != department_id:
            flash("Semester not in your department.", "danger")
            return redirect(url_for("hod_manage_semesters"))

        if db.query(Subject).filter_by(semester_id=semester_id).count():
            flash("Cannot delete — subjects are still in this semester.", "danger")
            return redirect(url_for("hod_manage_semesters"))

        db.delete(sem)
        db.commit()
        flash("✅ Semester deleted.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error: {e}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_manage_semesters"))


# ============================================================
# HOD | Manage Students — WITH semester_map AND set_student_module
# ============================================================
@app.route("/hod/students", methods=["GET", "POST"])
def hod_students_list():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=department_id).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("logout"))

        courses = (
            db.query(Course)
            .filter_by(department_id=department_id)
            .order_by(Course.name.asc())
            .all()
        )
        course_ids = [c.id for c in courses]

        if request.method == "POST":
            student_id = request.form.get("student_id", type=int)
            action = request.form.get("action")

            student = db.query(StudentProfile).filter_by(id=student_id).first()
            if not student or student.course_id not in course_ids:
                flash("Student not found in your department.", "danger")
                return redirect(url_for("hod_students_list"))

            if action == "toggle_block":
                student.blocked = not student.blocked
                db.commit()
                state = "blocked" if student.blocked else "activated"
                flash(f"✅ Student {state}.", "success")
            elif action == "delete":
                try:
                    _delete_student_tree(db, student)
                    db.commit()
                    flash("✅ Student deleted.", "success")
                except Exception as e:
                    db.rollback()
                    flash(f"Error deleting student: {e}", "danger")

            return redirect(url_for("hod_students_list"))

        search_query = request.args.get("search", "").strip()

        if course_ids:
            q = (
                db.query(StudentProfile)
                .options(
                    joinedload(StudentProfile.course),
                    joinedload(StudentProfile.current_semester),
                )
                .filter(StudentProfile.course_id.in_(course_ids))
            )
            if search_query:
                q = q.filter(or_(
                    StudentProfile.full_name.ilike(f"%{search_query}%"),
                    StudentProfile.admission_number.ilike(f"%{search_query}%"),
                    StudentProfile.phone_number.ilike(f"%{search_query}%"),
                ))
            students = q.order_by(StudentProfile.full_name.asc()).all()
        else:
            students = []

        # ✅ Build semester_map for the Current Module dropdown
        semester_map = {}
        for cid in course_ids:
            sems = (
                db.query(Semester)
                .filter_by(course_id=cid)
                .order_by(Semester.created_at.asc())
                .all()
            )
            semester_map[cid] = [{"id": s.id, "name": s.name} for s in sems]

        return render_template(
            "hod/students_list.html",
            department=department,
            hod_name=session.get("username"),
            students=students,
            search_query=search_query,
            semester_map=semester_map,
        )
    finally:
        db.close()


# ✅ MISSING ROUTE — add this after hod_students_list
@app.route("/hod/student/<int:student_id>/set-module", methods=["POST"])
def hod_set_student_module(student_id):
    """Directly set a student's current module."""
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    semester_id = request.form.get("semester_id", type=int)

    db = SessionLocal()
    try:
        student = db.query(StudentProfile).filter_by(id=student_id).first()
        if not student:
            flash("Student not found.", "danger")
            return redirect(url_for("hod_students_list"))

        course = db.query(Course).filter_by(id=student.course_id).first()
        if not course or course.department_id != department_id:
            flash("Student not in your department.", "danger")
            return redirect(url_for("hod_students_list"))

        if semester_id:
            sem = db.query(Semester).filter_by(id=semester_id, course_id=student.course_id).first()
            if not sem:
                flash("That module doesn't belong to this student's course.", "danger")
                return redirect(url_for("hod_students_list"))

            student.current_semester_id = sem.id
            db.commit()
            flash(f"✅ {student.full_name} assigned to '{sem.name}'.", "success")
        else:
            student.current_semester_id = None
            db.commit()
            flash(f"✅ {student.full_name} module cleared.", "success")

    except Exception as e:
        db.rollback()
        flash(f"Error: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_students_list"))


@app.route("/hod/student/<int:student_id>/promote", methods=["POST"])
def hod_promote_student(student_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        student = db.query(StudentProfile).filter_by(id=student_id).first()
        if not student:
            flash("Student not found.", "danger")
            return redirect(url_for("hod_students_list"))

        course = db.query(Course).filter_by(id=student.course_id).first()
        if not course or course.department_id != department_id:
            flash("Student not in your department.", "danger")
            return redirect(url_for("hod_students_list"))

        semesters = (
            db.query(Semester)
            .filter_by(course_id=student.course_id)
            .order_by(Semester.created_at.asc())
            .all()
        )

        if not semesters:
            flash("No semesters defined for this course yet.", "warning")
            return redirect(url_for("hod_students_list"))

        current_sem_id = None
        for sem in semesters:
            has_marks = db.query(KnecMark).filter_by(
                student_id=student.id,
                semester_id=sem.id,
            ).count() > 0

            if not has_marks:
                current_sem_id = sem.id
                break

        if current_sem_id is None:
            flash(f"{student.full_name} is already in the final semester.", "warning")
            return redirect(url_for("hod_students_list"))

        current_idx = next(
            (i for i, s in enumerate(semesters) if s.id == current_sem_id),
            None
        )
        if current_idx is None or current_idx + 1 >= len(semesters):
            flash("Already in the last semester — cannot promote further.", "warning")
            return redirect(url_for("hod_students_list"))

        next_sem = semesters[current_idx + 1]

        db.query(Semester).filter(
            Semester.course_id == student.course_id,
        ).update({"is_current": False}, synchronize_session=False)

        next_sem.is_current = True
        db.commit()

        flash(f"✅ {student.full_name} promoted to '{next_sem.name}'.", "success")

    except Exception as e:
        db.rollback()
        flash(f"Error promoting student: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for("hod_students_list"))


# HOD | Course Detail
@app.route("/hod/course/<int:course_id>")
def hod_course_detail(course_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]

    db = SessionLocal()
    try:
        course = db.query(Course).filter_by(id=course_id, department_id=department_id).first()
        if not course:
            flash("Course not found in your department.", "danger")
            return redirect(url_for("hod_dashboard"))

        students = (
            db.query(StudentProfile)
            .options(joinedload(StudentProfile.course))
            .filter_by(course_id=course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )

        semesters = (
            db.query(Semester)
            .filter_by(course_id=course_id)
            .order_by(Semester.created_at.asc())
            .all()
        )
        semester_groups = []
        for sem in semesters:
            subs = (
                db.query(Subject)
                .filter_by(course_id=course_id, semester_id=sem.id)
                .order_by(Subject.name.asc())
                .all()
            )
            semester_groups.append({
                "semester": sem,
                "subjects": subs,
                "is_current": bool(sem.is_current),
            })

        unassigned = (
            db.query(Subject)
            .filter(Subject.course_id == course_id)
            .filter(Subject.semester_id.is_(None))
            .order_by(Subject.name.asc())
            .all()
        )

        active_count = sum(1 for s in students if not s.blocked)
        blocked_count = len(students) - active_count

        return render_template(
            "hod/course_detail.html",
            department=db.query(Department).filter_by(id=department_id).first(),
            course=course,
            students=students,
            semester_groups=semester_groups,
            unassigned_subjects=unassigned,
            active_count=active_count,
            blocked_count=blocked_count,
        )
    finally:
        db.close()


# HOD | Student Detail
@app.route("/hod/student/<int:student_id>")
def hod_student_detail(student_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]

    db = SessionLocal()
    try:
        student = db.query(StudentProfile).filter_by(id=student_id).first()
        if not student:
            flash("Student not found.", "danger")
            return redirect(url_for("hod_students_list"))

        course = db.query(Course).filter_by(id=student.course_id).first()
        if not course or course.department_id != department_id:
            flash("That student is not in your department.", "danger")
            return redirect(url_for("hod_dashboard"))

        semesters = (
            db.query(Semester)
            .filter_by(course_id=student.course_id)
            .order_by(Semester.created_at.asc())
            .all()
        )

        semester_reports = []
        grand_total = 0
        grand_count = 0

        for sem in semesters:
            subjects = (
                db.query(Subject)
                .filter_by(course_id=student.course_id, semester_id=sem.id)
                .order_by(Subject.name.asc())
                .all()
            )
            rows = []
            sem_total = 0
            sem_count = 0

            for subj in subjects:
                m = db.query(KnecMark).filter_by(
                    student_id=student.id,
                    subject_id=subj.id,
                    semester_id=sem.id,
                ).first()
                c1 = m.cat1 if m and m.cat1 is not None else None
                c2 = m.cat2 if m and m.cat2 is not None else None
                fn = m.final if m and m.final is not None else None
                total = (c1 or 0) + (c2 or 0) + (fn or 0) if m else 0
                has_marks = m is not None

                rows.append({
                    "subject": subj,
                    "cat1": c1,
                    "cat2": c2,
                    "final": fn,
                    "total": total,
                    "has_marks": has_marks,
                })
                if has_marks:
                    sem_total += total
                    sem_count += 1

            avg = round(sem_total / sem_count, 2) if sem_count else None
            semester_reports.append({
                "semester": sem,
                "rows": rows,
                "is_current": bool(sem.is_current),
                "sem_total": sem_total,
                "sem_count": sem_count,
                "avg": avg,
            })
            grand_total += sem_total
            grand_count += sem_count

        overall_avg = round(grand_total / grand_count, 2) if grand_count else None

        return render_template(
            "hod/student_detail.html",
            department=db.query(Department).filter_by(id=department_id).first(),
            student=student,
            course=course,
            semester_reports=semester_reports,
            grand_total=grand_total,
            grand_count=grand_count,
            overall_avg=overall_avg,
        )
    finally:
        db.close()


# HOD | All Students Directory
@app.route("/hod/all-students", methods=["GET"])
def hod_all_students():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=department_id).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("logout"))

        courses = (
            db.query(Course)
            .filter_by(department_id=department_id)
            .order_by(Course.name.asc())
            .all()
        )
        course_ids = [c.id for c in courses]

        search = request.args.get("search", "").strip()
        course_id = request.args.get("course_id", type=int)
        exam_type = request.args.get("exam_type", "").strip()

        query = db.query(StudentProfile).options(joinedload(StudentProfile.course))
        if course_ids:
            query = query.filter(StudentProfile.course_id.in_(course_ids))
        else:
            query = query.filter(False)

        if search:
            like = f"%{search}%"
            query = query.filter(or_(
                StudentProfile.full_name.ilike(like),
                StudentProfile.admission_number.ilike(like),
                StudentProfile.phone_number.ilike(like),
            ))
        if course_id and course_id in course_ids:
            query = query.filter(StudentProfile.course_id == course_id)
        if exam_type:
            query = query.filter(StudentProfile.exam_type == exam_type)

        students = query.order_by(StudentProfile.full_name.asc()).all()

        return render_template(
            "hod/all_students.html",
            department=department,
            students=students,
            courses=courses,
            hod_name=session.get("username"),
            search=search,
            course_id=course_id,
            exam_type=exam_type,
        )
    finally:
        db.close()


# HOD | All Courses & Subjects
@app.route("/hod/all-courses", methods=["GET"])
def hod_all_courses():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=department_id).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("logout"))

        courses = (
            db.query(Course)
            .filter_by(department_id=department_id)
            .order_by(Course.name.asc())
            .all()
        )

        courses_data = []
        for course in courses:
            students_count = db.query(StudentProfile).filter_by(course_id=course.id).count()
            semesters = (
                db.query(Semester)
                .filter_by(course_id=course.id)
                .order_by(Semester.created_at.asc())
                .all()
            )

            semesters_data = []
            total_subjects = 0
            for sem in semesters:
                subs = (
                    db.query(Subject)
                    .filter_by(course_id=course.id, semester_id=sem.id)
                    .order_by(Subject.name.asc())
                    .all()
                )
                total_subjects += len(subs)
                semesters_data.append({"semester": sem, "subjects": subs})

            unassigned = (
                db.query(Subject)
                .filter(Subject.course_id == course.id)
                .filter(Subject.semester_id.is_(None))
                .order_by(Subject.name.asc())
                .all()
            )
            total_subjects += len(unassigned)

            courses_data.append({
                "course": course,
                "students_count": students_count,
                "semesters": semesters_data,
                "unassigned": unassigned,
                "total_subjects": total_subjects,
            })

        return render_template(
            "hod/all_courses.html",
            department=department,
            courses=courses_data,
            hod_name=session.get("username"),
        )
    finally:
        db.close()


# HOD | Enter Marks
@app.route("/hod/marks")
def hod_marks_picker():
    if session.get("role") != "hod":
        return redirect(url_for("login"))
    return redirect(url_for("hod_marks_department", dept_id=session["department_id"]))


@app.route("/hod/marks/department/<int:dept_id>", methods=["GET"])
def hod_marks_department(dept_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    if dept_id != session["department_id"]:
        flash("You can only manage your own department.", "danger")
        return redirect(url_for("hod_dashboard"))

    db = SessionLocal()
    try:
        department = db.query(Department).filter_by(id=dept_id).first()
        if not department:
            flash("Department not found.", "danger")
            return redirect(url_for("logout"))

        courses = (
            db.query(Course)
            .filter_by(department_id=dept_id)
            .order_by(Course.name.asc())
            .all()
        )

        course_id = request.args.get("course_id", type=int)
        semester_id = request.args.get("semester_id", type=int)

        course = None
        semester = None
        subjects = []
        semesters = []

        if course_id:
            course = db.query(Course).filter_by(id=course_id, department_id=dept_id).first()
        if course:
            semesters = (
                db.query(Semester)
                .filter_by(course_id=course.id)
                .order_by(Semester.created_at.desc())
                .all()
            )
            if semester_id:
                semester = db.query(Semester).filter_by(id=semester_id, course_id=course.id).first()
            elif semesters:
                semester = next((s for s in semesters if s.is_current), semesters[0])

            if semester:
                subjects = (
                    db.query(Subject)
                    .filter_by(course_id=course.id, semester_id=semester.id)
                    .order_by(Subject.name.asc())
                    .all()
                )

        return render_template(
            "admin/marks_step2_course.html",
            department=department,
            courses=courses,
            course=course,
            semesters=semesters,
            semester=semester,
            subjects=subjects,
            hod_mode=True,
        )
    finally:
        db.close()


@app.route("/hod/marks/enter", methods=["GET", "POST"])
def hod_enter_marks():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        department_id = session["department_id"]
        src = request.form if request.method == "POST" else request.args
        subject_id = src.get("subject_id", type=int)
        semester_id = src.get("semester_id", type=int)

        if not subject_id or not semester_id:
            flash("Pick a subject and semester.", "danger")
            return redirect(url_for("hod_marks_picker"))

        subject = db.query(Subject).filter_by(id=subject_id).first()
        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not subject or not semester:
            flash("Subject or semester not found.", "danger")
            return redirect(url_for("hod_marks_picker"))

        course = db.query(Course).filter_by(id=subject.course_id).first()
        if not course or course.department_id != department_id:
            flash("You can only manage marks for your department.", "danger")
            return redirect(url_for("hod_dashboard"))

        if subject.course_id != semester.course_id:
            flash("Subject and semester belong to different courses.", "danger")
            return redirect(url_for("hod_marks_picker"))

        students = (
            db.query(StudentProfile)
            .filter_by(course_id=subject.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )

        if request.method == "POST":
            def to_int_or_none(v):
                v = (v or "").strip()
                return int(v) if v.isdigit() else None

            saved = 0
            for s in students:
                c1 = to_int_or_none(request.form.get(f"cat1_{s.id}"))
                c2 = to_int_or_none(request.form.get(f"cat2_{s.id}"))
                fn = to_int_or_none(request.form.get(f"final_{s.id}"))

                existing = (
                    db.query(KnecMark)
                    .filter_by(student_id=s.id, subject_id=subject_id, semester_id=semester_id)
                    .first()
                )

                if c1 is None and c2 is None and fn is None:
                    if existing:
                        db.delete(existing)
                    continue

                if c1 is not None and not (0 <= c1 <= 15):
                    flash(f"CAT1 for {s.full_name} must be 0–15.", "warning")
                    continue
                if c2 is not None and not (0 <= c2 <= 15):
                    flash(f"CAT2 for {s.full_name} must be 0–15.", "warning")
                    continue
                if fn is not None and not (0 <= fn <= 70):
                    flash(f"Final for {s.full_name} must be 0–70.", "warning")
                    continue

                if existing:
                    existing.cat1 = c1
                    existing.cat2 = c2
                    existing.final = fn
                    existing.entered_at = datetime.utcnow()
                else:
                    db.add(KnecMark(
                        student_id=s.id, subject_id=subject_id,
                        semester_id=semester_id,
                        cat1=c1, cat2=c2, final=fn,
                    ))
                saved += 1

            db.commit()
            flash(f"✅ Saved marks for {saved} student(s).", "success")
            return redirect(url_for(
                "hod_marks_department",
                dept_id=department_id,
                course_id=subject.course_id,
                semester_id=semester_id,
            ))

        existing = {
            m.student_id: m
            for m in db.query(KnecMark).filter_by(subject_id=subject_id, semester_id=semester_id).all()
        }

        return render_template(
            "admin/marks_entry.html",
            subject=subject, semester=semester,
            course=course, department=course.department,
            students=students, existing=existing,
            hod_mode=True,
        )
    finally:
        db.close()


@app.route("/hod/marks/class")
def hod_class_marks():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    semester_id = request.args.get("semester_id", type=int)

    db = SessionLocal()
    try:
        if not semester_id:
            flash("Pick a semester.", "warning")
            return redirect(url_for("hod_marks_picker"))

        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not semester:
            flash("Semester not found.", "danger")
            return redirect(url_for("hod_marks_picker"))

        course = db.query(Course).filter_by(id=semester.course_id).first()
        if not course or course.department_id != department_id:
            flash("That semester is not in your department.", "danger")
            return redirect(url_for("hod_dashboard"))

        students = (
            db.query(StudentProfile)
            .filter_by(course_id=semester.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )
        subjects = (
            db.query(Subject)
            .filter_by(course_id=semester.course_id, semester_id=semester_id)
            .order_by(Subject.name.asc())
            .all()
        )

        marks = {}
        for m in db.query(KnecMark).filter_by(semester_id=semester_id).all():
            marks.setdefault(m.student_id, {})[m.subject_id] = m

        return render_template(
            "admin/class_marks.html",
            semester=semester, students=students, subjects=subjects, marks=marks,
            hod_mode=True,
        )
    finally:
        db.close()


@app.route("/hod/student/<int:student_id>/marks", methods=["GET", "POST"])
def hod_student_marks(student_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]

    db = SessionLocal()
    try:
        student = db.query(StudentProfile).filter_by(id=student_id).first()
        if not student:
            flash("Student not found.", "danger")
            return redirect(url_for("hod_students_list"))

        course = db.query(Course).filter_by(id=student.course_id).first()
        if not course or course.department_id != department_id:
            flash("That student is not in your department.", "danger")
            return redirect(url_for("hod_dashboard"))

        semester_id = (
            request.args.get("semester_id", type=int)
            or request.form.get("semester_id", type=int)
        )

        semesters = (
            db.query(Semester)
            .filter_by(course_id=student.course_id)
            .order_by(Semester.created_at.desc())
            .all()
        )

        semester = None
        if semester_id:
            semester = next((s for s in semesters if s.id == semester_id), None)
        elif semesters:
            semester = next((s for s in semesters if s.is_current), semesters[0])

        subjects = []
        if semester:
            subjects = (
                db.query(Subject)
                .filter_by(course_id=student.course_id, semester_id=semester.id)
                .order_by(Subject.name.asc())
                .all()
            )

        if request.method == "POST" and semester:
            def to_int_or_none(v):
                v = (v or "").strip()
                return int(v) if v.isdigit() else None

            saved = 0
            for subj in subjects:
                c1 = to_int_or_none(request.form.get(f"cat1_{subj.id}"))
                c2 = to_int_or_none(request.form.get(f"cat2_{subj.id}"))
                fn = to_int_or_none(request.form.get(f"final_{subj.id}"))

                existing = (
                    db.query(KnecMark)
                    .filter_by(
                        student_id=student.id,
                        subject_id=subj.id,
                        semester_id=semester.id,
                    )
                    .first()
                )

                if c1 is None and c2 is None and fn is None:
                    if existing:
                        db.delete(existing)
                    continue

                if c1 is not None and not (0 <= c1 <= 15):
                    flash(f"CAT1 for {subj.name} must be 0–15.", "warning")
                    continue
                if c2 is not None and not (0 <= c2 <= 15):
                    flash(f"CAT2 for {subj.name} must be 0–15.", "warning")
                    continue
                if fn is not None and not (0 <= fn <= 70):
                    flash(f"Final for {subj.name} must be 0–70.", "warning")
                    continue

                if existing:
                    existing.cat1 = c1
                    existing.cat2 = c2
                    existing.final = fn
                    existing.entered_at = datetime.utcnow()
                else:
                    db.add(KnecMark(
                        student_id=student.id,
                        subject_id=subj.id,
                        semester_id=semester.id,
                        cat1=c1, cat2=c2, final=fn,
                    ))
                saved += 1

            db.commit()
            flash(f"✅ Saved marks for {saved} subject(s).", "success")
            return redirect(url_for(
                "hod_student_marks",
                student_id=student.id,
                semester_id=semester.id,
            ))

        rows = []
        for subj in subjects:
            m = (
                db.query(KnecMark)
                .filter_by(
                    student_id=student.id,
                    subject_id=subj.id,
                    semester_id=semester.id if semester else None,
                )
                .first()
            )
            rows.append({
                "subject": subj,
                "cat1": m.cat1 if m and m.cat1 is not None else None,
                "cat2": m.cat2 if m and m.cat2 is not None else None,
                "final": m.final if m and m.final is not None else None,
            })

        return render_template(
            "hod/student_marks.html",
            student=student,
            course=course,
            semester=semester,
            semesters=semesters,
            rows=rows,
        )
    finally:
        db.close()


@app.route("/hod/marks/class/export")
def hod_class_marks_export():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    semester_id = request.args.get("semester_id", type=int)
    if not semester_id:
        flash("Pick a semester.", "warning")
        return redirect(url_for("hod_marks_picker"))

    db = SessionLocal()
    try:
        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not semester:
            flash("Semester not found.", "danger")
            return redirect(url_for("hod_dashboard"))

        course = db.query(Course).filter_by(id=semester.course_id).first()
        if not course or course.department_id != department_id:
            flash("That semester is not in your department.", "danger")
            return redirect(url_for("hod_dashboard"))

        students = (
            db.query(StudentProfile)
            .filter_by(course_id=semester.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )
        subjects = (
            db.query(Subject)
            .filter_by(course_id=semester.course_id, semester_id=semester_id)
            .order_by(Subject.name.asc())
            .all()
        )

        marks = {}
        for m in db.query(KnecMark).filter_by(semester_id=semester_id).all():
            marks.setdefault(m.student_id, {})[m.subject_id] = m

        data = []
        for s in students:
            row = {
                "Admission No": s.admission_number,
                "Student Name": s.full_name,
                "Exam Type": (s.exam_type or "").upper(),
            }
            grand_total = 0
            counted = 0
            for subj in subjects:
                m = marks.get(s.id, {}).get(subj.id)
                c1 = m.cat1 if m and m.cat1 is not None else ""
                c2 = m.cat2 if m and m.cat2 is not None else ""
                fn = m.final if m and m.final is not None else ""
                total = (m.cat1 or 0) + (m.cat2 or 0) + (m.final or 0) if m else ""

                prefix = subj.name.split(" ", 1)[0]
                row[f"{prefix} CAT1"] = c1
                row[f"{prefix} CAT2"] = c2
                row[f"{prefix} Final"] = fn
                row[f"{prefix} Total"] = total

                if isinstance(total, int):
                    grand_total += total
                    counted += 1

            row["Grand Total"] = grand_total
            row["Average"] = round(grand_total / counted, 2) if counted else ""
            data.append(row)

        df = pd.DataFrame(data)
        identity_cols = ["Admission No", "Student Name", "Exam Type"]
        total_cols = ["Grand Total", "Average"]
        subject_cols = [c for c in df.columns if c not in identity_cols + total_cols]
        df = df[identity_cols + subject_cols + total_cols]

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Marks")
        buf.seek(0)

        safe_course = (course.name if course else "Class").replace(" ", "_")[:30]
        safe_sem = semester.name.replace(" ", "_")

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"Marks_{safe_course}_{safe_sem}.xlsx",
        )
    finally:
        db.close()


@app.route("/hod/reports/class/export")
def hod_reports_class_export():
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]
    semester_id = request.args.get("semester_id", type=int)
    if not semester_id:
        flash("Pick a semester.", "warning")
        return redirect(url_for("hod_marks_picker"))

    db = SessionLocal()
    try:
        semester = db.query(Semester).filter_by(id=semester_id).first()
        if not semester:
            flash("Semester not found.", "danger")
            return redirect(url_for("hod_dashboard"))

        course = db.query(Course).filter_by(id=semester.course_id).first()
        if not course or course.department_id != department_id:
            flash("That semester is not in your department.", "danger")
            return redirect(url_for("hod_dashboard"))

        students = (
            db.query(StudentProfile)
            .filter_by(course_id=semester.course_id)
            .order_by(StudentProfile.full_name.asc())
            .all()
        )

        data = []
        for s in students:
            is_knec = (s.exam_type or "").strip().lower() == "knec"
            total_sum = 0
            subjects_count = 0
            if is_knec:
                for m in db.query(KnecMark).filter_by(student_id=s.id, semester_id=semester_id).all():
                    total_sum += (m.cat1 or 0) + (m.cat2 or 0) + (m.final or 0)
                    subjects_count += 1

            avg = round(total_sum / subjects_count, 2) if subjects_count else ""
            grade = ""
            if isinstance(avg, float):
                grade = 1 if avg >= 84 else 2 if avg >= 75 else 3 if avg >= 66 else 4 if avg >= 60 else 5 if avg >= 50 else 6 if avg >= 40 else 7

            result = ""
            if isinstance(avg, float):
                result = "DISTINCTION" if avg >= 75 else "CREDIT" if avg >= 60 else "PASS" if avg >= 40 else "REFER"

            data.append({
                "Admission No": s.admission_number,
                "Student Name": s.full_name,
                "Course": course.name if course else "",
                "Semester": semester.name,
                "Exam Type": (s.exam_type or "").upper(),
                "Subjects Graded": subjects_count,
                "Total Marks": total_sum,
                "Average": avg,
                "Grade": grade,
                "Overall Result": result,
            })

        df = pd.DataFrame(data)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Class Report")
        buf.seek(0)

        safe_course = (course.name if course else "Class").replace(" ", "_")[:30]
        safe_sem = semester.name.replace(" ", "_")

        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"ClassReport_{safe_course}_{safe_sem}.xlsx",
        )
    finally:
        db.close()


@app.route("/hod/reports/student/<int:student_id>/<int:semester_id>")
def hod_report_student(student_id, semester_id):
    if session.get("role") != "hod":
        return redirect(url_for("login"))

    department_id = session["department_id"]

    db = SessionLocal()
    try:
        student = db.query(StudentProfile).filter_by(id=student_id).first()
        if not student:
            return "Student not found", 404

        course = db.query(Course).filter_by(id=student.course_id).first()
        if not course or course.department_id != department_id:
            flash("That student is not in your department.", "danger")
            return redirect(url_for("hod_dashboard"))

        sem = db.query(Semester).filter_by(id=semester_id, course_id=student.course_id).first()
        if not sem:
            return "Semester not found for this student", 404

        subjects = (
            db.query(Subject)
            .filter_by(course_id=student.course_id, semester_id=sem.id)
            .order_by(Subject.name.asc())
            .all()
        )

        rows = []
        for subj in subjects:
            m = db.query(KnecMark).filter_by(
                student_id=student.id, subject_id=subj.id, semester_id=sem.id
            ).first()
            c1 = m.cat1 if m and m.cat1 is not None else None
            c2 = m.cat2 if m and m.cat2 is not None else None
            fn = m.final if m and m.final is not None else None
            rows.append({
                "subject": subj, "cat1": c1, "cat2": c2, "final": fn,
                "total": (c1 or 0) + (c2 or 0) + (fn or 0),
            })

        return render_template(
            "students/print_report.html",
            student=student, semester=sem, rows=rows,
            current_year=datetime.now().year, admin_mode=True,
        )
    finally:
        db.close()


# ============================================================
# SECTION 4: STUDENT ROUTES
# ============================================================
@app.route("/complete_profile", methods=["GET", "POST"])
def complete_profile():
    if "user_id" not in session:
        return redirect(url_for("login"))

    db = SessionLocal()
    user_id = session["user_id"]

    if request.method == "POST":
        full_name = request.form.get("full_name").strip()
        exam_type = request.form.get("exam_type")
        course_id = request.form.get("course_id")
        admission_number = request.form.get("admission_number").strip()
        phone_number = request.form.get("phone_number").strip()

        try:
            if db.query(StudentProfile).filter_by(admission_number=admission_number).first():
                flash("❌ Admission number already exists.", "danger")
                courses = db.query(Course).all()
                return render_template("students/complete_profile.html", courses=courses,
                                       full_name=full_name, exam_type=exam_type,
                                       admission_number=admission_number, phone_number=phone_number)

            if db.query(StudentProfile).filter_by(user_id=user_id).first():
                flash("❌ You have already completed your profile.", "warning")
                return redirect(url_for("student_dashboard"))

            db.add(StudentProfile(
                full_name=full_name, exam_type=exam_type,
                course_id=course_id, admission_number=admission_number,
                phone_number=phone_number, user_id=user_id,
            ))
            db.commit()
            flash("✅ Profile completed successfully!", "success")
            return redirect(url_for("student_dashboard"))
        except Exception:
            db.rollback()
            flash("❌ Something went wrong. Try again.", "danger")
            courses = db.query(Course).all()
            return render_template("students/complete_profile.html", courses=courses)
        finally:
            db.close()

    courses = db.query(Course).all()
    db.close()
    return render_template("students/complete_profile.html", courses=courses)


@app.route("/student/dashboard")
def student_dashboard():
    if "username" not in session or session.get("role") != "student":
        flash("Please log in as a student first.", "error")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        user = db.query(User).filter_by(username=session["username"]).first()
        if not user:
            flash("User not found.", "error")
            return redirect(url_for("logout"))

        student_profile = db.query(StudentProfile).filter_by(user_id=user.id).first()
        if not student_profile:
            flash("Complete your profile before proceeding.", "warning")
            return redirect(url_for("complete_profile"))

        is_knec = (student_profile.exam_type or "").strip().lower() == "knec"

        semesters = (
            db.query(Semester)
            .filter_by(course_id=student_profile.course_id)
            .order_by(Semester.created_at.asc())
            .all()
        )

        all_quizzes = (
            db.query(Quiz)
            .filter_by(course_id=student_profile.course_id, status="active")
            .order_by(Quiz.id.asc())
            .all()
        )
        quizzes_by_subject = {}
        for q in all_quizzes:
            quizzes_by_subject.setdefault(q.subject_id, []).append(q)

        results = db.query(Result).filter_by(student_id=student_profile.id).all()
        results_by_quiz = {r.quiz_id: r for r in results}
        taken_quiz_ids = set(results_by_quiz.keys())

        def build(subject):
            qs = quizzes_by_subject.get(subject.id, [])
            taken = sum(1 for q in qs if q.id in taken_quiz_ids)
            row = {
                "subject": subject, "quizzes": qs,
                "total": len(qs), "taken": taken,
                "percent": round(taken / len(qs) * 100) if qs else 0,
                "cat1": None, "cat2": None, "final": None, "total_marks": None,
            }
            if is_knec and subject.semester_id:
                m = db.query(KnecMark).filter_by(
                    student_id=student_profile.id,
                    subject_id=subject.id,
                    semester_id=subject.semester_id,
                ).first()
                if m:
                    row["cat1"] = m.cat1
                    row["cat2"] = m.cat2
                    row["final"] = m.final
                    row["total_marks"] = (m.cat1 or 0) + (m.cat2 or 0) + (m.final or 0)
            return row

        all_subjects = (
            db.query(Subject)
            .filter_by(course_id=student_profile.course_id)
            .order_by(Subject.name.asc())
            .all()
        )
        all_subjects_data = [build(s) for s in all_subjects]

        current_semester = None
        semester_groups = []
        for sem in semesters:
            subs = (
                db.query(Subject)
                .filter_by(course_id=student_profile.course_id, semester_id=sem.id)
                .order_by(Subject.name.asc())
                .all()
            )
            if sem.is_current:
                current_semester = sem
            semester_groups.append({
                "semester": sem,
                "subjects": [build(s) for s in subs],
                "count": len(subs),
                "is_current": bool(sem.is_current),
            })

        unassigned = (
            db.query(Subject)
            .filter(Subject.course_id == student_profile.course_id)
            .filter(Subject.semester_id.is_(None))
            .order_by(Subject.name.asc())
            .all()
        )
        unassigned_subjects = [build(s) for s in unassigned]

        messages = (
            db.query(Message)
            .filter(
                (Message.target_type == "all") |
                ((Message.target_type == "course") & (Message.course_id == student_profile.course_id))
            )
            .order_by(Message.created_at.desc())
            .all()
        )

        return render_template(
            "students/student_dashboard.html",
            username=user.username,
            student_profile=student_profile,
            is_knec=is_knec,
            semester_groups=semester_groups,
            current_semester=current_semester,
            unassigned_subjects=unassigned_subjects,
            all_subjects_data=all_subjects_data,
            total_subjects=len(all_subjects),
            taken_quiz_ids=list(taken_quiz_ids),
            results_by_quiz=results_by_quiz,
            messages_from_admin=messages,
            current_year=datetime.now().year,
        )
    finally:
        db.close()


@app.route("/take_exam/<int:quiz_id>", methods=["GET", "POST"])
def take_exam(quiz_id):
    if "user_id" not in session:
        flash("Please log in first.", "warning")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        student = db.query(StudentProfile).filter_by(user_id=session["user_id"]).first()
        if not student:
            flash("Student profile not found.", "danger")
            return redirect(url_for("student_dashboard"))

        if student.blocked:
            flash("You can't access this exam. Please clear the fee to regain access.", "danger")
            return redirect(url_for("student_dashboard"))

        quiz = db.query(Quiz).filter_by(id=quiz_id).first()
        if not quiz:
            flash("Quiz not found.", "danger")
            return redirect(url_for("student_dashboard"))

        db.query(ActivityLog).filter_by(student_id=student.id, is_active=True).update({"is_active": False})
        db.add(ActivityLog(student_id=student.id, activity_type="exam", is_active=True))
        db.commit()

        if request.method == "POST":
            score = 0
            total_marks = 0

            questions = db.query(Question).filter_by(quiz_id=quiz.id).order_by(Question.id.asc()).all()

            for q in questions:
                selected = request.form.get(f"question_{q.id}")
                total_marks += q.marks
                if selected and selected.strip().lower() == (q.correct_option or "").strip().lower():
                    score += q.marks

            percentage = (score / total_marks) * 100 if total_marks > 0 else 0

            db.add(Result(student_id=student.id, quiz_id=quiz.id,
                          score=score, total_marks=total_marks, percentage=percentage))

            db.query(ActivityLog).filter_by(
                student_id=student.id, activity_type="exam", is_active=True,
            ).update({"is_active": False})
            db.commit()

            return render_template(
                "students/take_exam.html",
                quiz=quiz, questions=[], student=student,
                show_result=True, score=score, total_marks=total_marks, percentage=percentage,
            )

        questions = db.query(Question).filter_by(quiz_id=quiz.id).order_by(Question.id.asc()).all()
        return render_template("students/take_exam.html", quiz=quiz, questions=questions,
                               student=student, show_result=False)
    finally:
        db.close()


@app.route("/student/results")
def student_results():
    if "username" not in session or session.get("role") != "student":
        flash("Please log in as a student first.", "error")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        user = db.query(User).filter_by(username=session["username"]).first()
        if not user:
            flash("User not found.", "error")
            return redirect(url_for("logout"))

        student = db.query(StudentProfile).filter_by(user_id=user.id).first()
        if not student:
            flash("Student profile not found.", "error")
            return redirect(url_for("logout"))

        is_knec = (student.exam_type or "").strip().lower() == "knec"

        if not is_knec:
            results = (
                db.query(Result)
                .options(joinedload(Result.quiz).joinedload(Quiz.subject))
                .join(Quiz, Result.quiz_id == Quiz.id)
                .filter(Result.student_id == student.id)
                .order_by(Result.taken_on.desc())
                .all()
            )
            return render_template("students/results.html", is_knec=False, results=results,
                                   student=student, current_year=datetime.now().year)

        semesters = (
            db.query(Semester)
            .filter_by(course_id=student.course_id)
            .order_by(Semester.created_at.asc())
            .all()
        )

        semester_reports = []
        for sem in semesters:
            subjects = (
                db.query(Subject)
                .filter_by(course_id=student.course_id, semester_id=sem.id)
                .order_by(Subject.name.asc())
                .all()
            )
            rows = []
            for subj in subjects:
                m = db.query(KnecMark).filter_by(
                    student_id=student.id, subject_id=subj.id, semester_id=sem.id,
                ).first()
                c1 = m.cat1 if m and m.cat1 is not None else None
                c2 = m.cat2 if m and m.cat2 is not None else None
                fn = m.final if m and m.final is not None else None
                rows.append({
                    "subject": subj, "cat1": c1, "cat2": c2, "final": fn,
                    "total": (c1 or 0) + (c2 or 0) + (fn or 0),
                })
            semester_reports.append({"semester": sem, "rows": rows, "is_current": bool(sem.is_current)})

        return render_template("students/results.html", is_knec=True, student=student,
                               semester_reports=semester_reports, current_year=datetime.now().year)
    finally:
        db.close()


@app.route("/student/print_report")
def student_print_report():
    if "username" not in session or session.get("role") != "student":
        flash("Please log in as a student first.", "error")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        user = db.query(User).filter_by(username=session["username"]).first()
        if not user:
            return redirect(url_for("logout"))

        student = db.query(StudentProfile).filter_by(user_id=user.id).first()
        if not student:
            return redirect(url_for("logout"))

        if (student.exam_type or "").strip().lower() != "knec":
            flash("Printable reports are only available for KNEC students.", "warning")
            return redirect(url_for("student_results"))

        sem_id = request.args.get("semester_id", type=int)
        if sem_id:
            sem = db.query(Semester).filter_by(id=sem_id, course_id=student.course_id).first()
        else:
            sem = db.query(Semester).filter_by(course_id=student.course_id, is_current=True).first()

        if not sem:
            flash("No semester selected.", "warning")
            return redirect(url_for("student_results"))

        subjects = (
            db.query(Subject)
            .filter_by(course_id=student.course_id, semester_id=sem.id)
            .order_by(Subject.name.asc())
            .all()
        )

        rows = []
        for subj in subjects:
            m = db.query(KnecMark).filter_by(
                student_id=student.id, subject_id=subj.id, semester_id=sem.id,
            ).first()
            c1 = m.cat1 if m and m.cat1 is not None else None
            c2 = m.cat2 if m and m.cat2 is not None else None
            fn = m.final if m and m.final is not None else None
            rows.append({
                "subject": subj, "cat1": c1, "cat2": c2, "final": fn,
                "total": (c1 or 0) + (c2 or 0) + (fn or 0),
            })

        return render_template("students/print_report.html", student=student,
                               semester=sem, rows=rows, current_year=datetime.now().year)
    finally:
        db.close()


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)