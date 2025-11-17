from flask import Flask, render_template, request, redirect, url_for, session, flash,Response
from werkzeug.utils import secure_filename
import pandas as pd
from flask import send_file
from datetime import datetime

import os
import re
import pandas as pd

from connections import SessionLocal
from models import User, Admin, Course, Subject, Question, Quiz, Video, Document,StudentProfile,Result,Message,ActivityLog
from utils import parse_docx_questions, get_quiz_status
from utils import extract_drive_id, get_drive_embed_url

from sqlalchemy.exc import IntegrityError


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB limit
app.secret_key = '132silas456sakwa789ayanga'

# -----------------------------
# Upload folders & allowed types
# -----------------------------
BASE_UPLOAD = os.path.join(os.getcwd(), 'static', 'uploads')
VIDEOS_UPLOAD_FOLDER = os.path.join(BASE_UPLOAD, 'videos')
DOCUMENTS_UPLOAD_FOLDER = os.path.join(BASE_UPLOAD, 'documents')
EXAMS_UPLOAD_FOLDER = os.path.join(BASE_UPLOAD, 'exams')
QUESTION_IMAGES_FOLDER = os.path.join('static', 'question_images')

for path in [BASE_UPLOAD, VIDEOS_UPLOAD_FOLDER, DOCUMENTS_UPLOAD_FOLDER, EXAMS_UPLOAD_FOLDER, QUESTION_IMAGES_FOLDER]:
    os.makedirs(path, exist_ok=True)

ALLOWED_DOCUMENT_EXTENSIONS = {'pdf', 'doc', 'docx', 'ppt', 'pptx'}
ALLOWED_VIDEO_EXTENSIONS = {'mp4', 'mov', 'avi', 'mkv'}
ALLOWED_EXAM_EXTENSIONS = {'docx'}

def allowed(filename, allowed_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set


# -----------------------------
# Routes: Auth + Dashboards
# -----------------------------
@app.route('/')
def home():
    return redirect(url_for('login'))
'''
login route
'''
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = False
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        db = SessionLocal()
        try:
            # Admin
            admin = db.query(Admin).filter_by(username=username).first()
            if admin and admin.password == password:
                session.update({'username': admin.username, 'user_id': admin.id, 'role': 'admin'})
                return redirect(url_for('admin_dashboard'))

            # Student
            user = db.query(User).filter_by(username=username).first()
            if user and user.password == password:
                session.update({'username': user.username, 'user_id': user.id, 'role': 'student'})
                return redirect(url_for('student_dashboard'))

            error = True
        finally:
            db.close()
    return render_template('login.html', error=error)
'''
logout route'''
@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))
'''
admin dashboard
'''
@app.route('/admin_dashboard')
def admin_dashboard():
    if session.get('role') != 'admin':
        return redirect(url_for('login'))
    db = SessionLocal()
    try:
        courses = db.query(Course).all()
        return render_template('admin/admin_dashboard.html', courses=courses, username=session.get('username'))
    finally:
        db.close()

# -----------------------------
# admin Courses & Subjects
# -----------------------------
@app.route('/add_course', methods=['GET', 'POST'])
def add_course():
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))  

    db = SessionLocal()
    try:
        if request.method == 'POST':
            course_name = request.form['course_name']
            course_level = request.form['course_level']

            existing_course = db.query(Course).filter_by(name=course_name).first()
            if existing_course:
                flash(f"Course '{course_name}' already exists!", "warning")
                return redirect(url_for('add_course'))

            new_course = Course(name=course_name, level=course_level)
            db.add(new_course)
            db.commit()
            flash("Course added successfully!", "success")
            return redirect(url_for('admin_dashboard'))

        return render_template('admin/add_course.html')

    except Exception as e:
        db.rollback()
        print("❌ Error in /add_course:", e)
        flash("An error occurred while adding the course.", "danger")
        return redirect(url_for('add_course'))

    finally:
        db.close()


'''
admin edit a course
'''
@app.route('/edit_course/<int:course_id>', methods=['GET', 'POST'])
def edit_course(course_id):
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))  
    
    db = SessionLocal()
    
   
    course = db.query(Course).filter(Course.id == course_id).first()
    
    if not course:
        return "Course not found", 404  

    if request.method == 'POST':
       
        course.name = request.form['course_name']
        course.level = request.form['course_level']  
        
        db.commit() 
        
        return redirect(url_for('admin_dashboard'))  
    
    
    return render_template('admin/edit_course.html', course=course)

@app.route('/delete_course/<int:course_id>', methods=['POST'])
def delete_course(course_id):
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))  
    
    db = SessionLocal()
    try:
        course = db.query(Course).filter(Course.id == course_id).first()
        if not course:
            return "Course not found", 404  

        # 🔍 Check if any students are enrolled in this course
        enrolled_students = db.query(StudentProfile).filter(StudentProfile.course_id == course_id).count()
        if enrolled_students > 0:
            flash(f"Cannot delete course because {enrolled_students} student(s) are still enrolled.", "danger")
            return redirect(url_for('manage_courses'))

        # ✅ Delete related documents first
        db.query(Document).filter(Document.course_id == course_id).delete()

        # ✅ Then delete the course
        db.delete(course)
        db.commit()
        flash("Course deleted successfully.", "success")
    except Exception as e:
        db.rollback()
        flash(f"Error deleting course: {str(e)}", "danger")
    finally:
        db.close()

    return redirect(url_for('manage_courses'))


'''
admin manage course 
'''
@app.route('/manage_courses', methods=['GET'])
def manage_courses():
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))  
    
    db = SessionLocal()
    courses = db.query(Course).all()  
    return render_template('admin/manage_courses.html', courses=courses)

'''
admin add subject
'''
@app.route('/add_subject', methods=['GET', 'POST'])
def add_subject():
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))  
    
    db = SessionLocal()

    if request.method == 'POST':
        subject_name = request.form['subject_name']
        course_id = request.form['course_id']

        existing_subject = db.query(Subject).filter_by(name=subject_name, course_id=course_id).first()
        if existing_subject:
            flash('Subject already exists for this course!', 'danger')
            return redirect(url_for('add_subject'))

        new_subject = Subject(name=subject_name, course_id=course_id)
        db.add(new_subject)
        db.commit()
        flash('Subject added successfully!', 'success')

        return redirect(url_for('manage_courses'))  
    
    
    courses = db.query(Course).all()
    return render_template('admin/add_subject.html', courses=courses)


'''admin edit subject'''

@app.route('/edit_subject/<int:subject_id>', methods=['GET', 'POST'])
def edit_subject(subject_id):
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))  
    db = SessionLocal()
    subject = db.query(Subject).filter(Subject.id == subject_id).first()

    if not subject:
        return "Subject not found", 404  

    if request.method == 'POST':
        
        subject.name = request.form['subject_name']
        db.commit()  
        return redirect(url_for('manage_courses'))  

    return render_template('admin/edit_subject.html', subject=subject)

'''
admin delete subject 
'''

@app.route('/delete_subject/<int:subject_id>', methods=['POST'])
def delete_subject(subject_id):
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))  
    
    db = SessionLocal()
    subject = db.query(Subject).filter(Subject.id == subject_id).first()

    if not subject:
        return "Subject not found", 404  
    
    db.delete(subject)  
    db.commit() 

    return redirect(url_for('manage_courses'))  


# -----------------------------
# admin Exams (.docx parsing)
# -----------------------------
@app.route('/upload_exam', methods=['GET', 'POST'])
def upload_exam():
    if session.get('role') != 'admin':
        return redirect(url_for('login'))
    db = SessionLocal()
    try:
        courses, subjects = db.query(Course).all(), db.query(Subject).all()
        if request.method == 'POST':
            title = request.form['title'].strip()
            course_id, subject_id = int(request.form['course']), int(request.form['subject'])
            duration = int(request.form.get('duration', 30))
            file = request.files.get('quiz_file')

            if not file or not allowed(file.filename, ALLOWED_EXAM_EXTENSIONS):
                flash('❌ Upload a valid .docx file.', 'danger')
                return redirect(request.url)

            filename = secure_filename(file.filename)
            file_path = os.path.join(EXAMS_UPLOAD_FOLDER, filename)
            file.save(file_path)

            questions = parse_docx_questions(file_path, image_output_dir=QUESTION_IMAGES_FOLDER)
            if not questions:
                flash("❌ No valid questions found.", "danger")
                return redirect(request.url)

            quiz = Quiz(title=title, course_id=course_id, subject_id=subject_id,
                        duration=duration, status='active')
            db.add(quiz)
            db.commit()
            db.refresh(quiz)

            for q in questions:
                db.add(Question(quiz_id=quiz.id, question_text=q.get("question", ""),
                                option_a=q.get("a", ""), option_b=q.get("b", ""),
                                option_c=q.get("c", ""), option_d=q.get("d", ""),
                                correct_option=q.get("answer", "").lower(),
                                marks=q.get("marks", 1), extra_content=q.get("extra_content"),
                                image=q.get("image")))
            db.commit()
            flash(f"✅ Uploaded quiz with {len(questions)} question(s).", "success")

            return render_template('admin/upload_exams.html', courses=courses, subjects=subjects, uploaded_quiz_id=quiz.id)

        return render_template('admin/upload_exams.html', courses=courses, subjects=subjects)
    finally:
        db.close()


# -----------------------------
# admin Videos
# -----------------------------
@app.route('/admin/upload_video', methods=['GET', 'POST'])
def upload_video():
    if session.get('role') != 'admin':
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        courses, subjects = db.query(Course).all(), db.query(Subject).all()
        videos = db.query(Video).all()

        if request.method == 'POST':
            title = request.form['title'].strip()
            course_id = int(request.form['course'])
            subject_id = int(request.form['subject'])
            drive_url = request.form['drive_url'].strip()  # get full URL

            # Optional: basic validation for Google Drive URL
            if "drive.google.com" not in drive_url:
                flash("❌ Invalid Google Drive link.", "danger")
                return redirect(request.url)

            db.add(Video(title=title, course_id=course_id, subject_id=subject_id, drive_url=drive_url))
            db.commit()
            flash(f"✅ Video '{title}' uploaded!", "success")
            return redirect(url_for('upload_video'))

        return render_template('admin/upload_video.html', courses=courses, subjects=subjects, videos=videos)
    finally:
        db.close()


'''
admin delete video
'''
@app.route('/admin/delete_video/<int:video_id>', methods=['POST'])
def delete_video(video_id):
    if session.get('role') != 'admin':
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        video = db.get(Video, video_id)  
        if not video:
            flash('Video not found.', 'danger')
            return redirect(url_for('upload_video'))

        db.delete(video)  
        db.commit()
        flash("✅ Video deleted.", "success")
        return redirect(url_for('upload_video'))
    finally:
        db.close()



# -----------------------------
#admin  Documents
# -----------------------------
@app.route("/admin/upload_document", methods=["GET", "POST"])
def upload_document():
    db = SessionLocal()
    try:
        courses = db.query(Course).all()
        subjects = db.query(Subject).all()

        if request.method == "POST":
            title = request.form.get("title", "").strip()
            course_id = request.form.get("course")
            subject_id = request.form.get("subject")
            file = request.files.get("document")  # must match <input name="document">

            if not title or not course_id or not subject_id:
                flash("⚠️ Please fill in all fields.", "danger")
                return redirect(request.url)

            if not file or not allowed(file.filename, ALLOWED_DOCUMENT_EXTENSIONS):
                flash("❌ Please upload a valid document (PDF, DOC, DOCX, PPT, PPTX).", "danger")
                return redirect(request.url)

            filename = secure_filename(file.filename)
            save_path = os.path.join(DOCUMENTS_UPLOAD_FOLDER, filename)
            file.save(save_path)

            new_doc = Document(
                title=title,
                filename=filename,
                course_id=int(course_id),
                subject_id=int(subject_id),
            )
            db.add(new_doc)
            db.commit()

            flash(f"✅ Document '{title}' uploaded successfully!", "success")
            return redirect(url_for("upload_document"))

        return render_template("admin/upload_document.html", courses=courses, subjects=subjects)
    finally:
        db.close()

# -----------------------------
# Admin: Manage Documents
# -----------------------------
from sqlalchemy.orm import joinedload

@app.route("/admin/documents")
def manage_documents():
    db = SessionLocal()
    try:
        documents = (
            db.query(Document)
            .options(joinedload(Document.subject), joinedload(Document.course))
            .all()
        )
        return render_template("admin/manage_documents.html", documents=documents)
    finally:
        db.close()


@app.route("/admin/documents/delete/<int:doc_id>", methods=["POST"])
def delete_document(doc_id):
    db = SessionLocal()
    try:
        doc = db.query(Document).get(doc_id)
        if not doc:
            flash("❌ Document not found.", "danger")
            return redirect(url_for("manage_documents"))

        file_path = os.path.join(DOCUMENTS_UPLOAD_FOLDER, doc.filename)
        if os.path.exists(file_path):
            os.remove(file_path)

        db.delete(doc)
        db.commit()

        flash(f"🗑️ Document '{doc.title}' deleted successfully!", "success")
    finally:
        db.close()

    return redirect(url_for("manage_documents"))


'''
admin manage students'''
from sqlalchemy.orm import joinedload
from sqlalchemy import or_

@app.route('/manage_students', methods=['GET', 'POST'])
def manage_students():
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))

    db = SessionLocal()

    if request.method == 'POST':
        student_id = request.form.get('student_id')
        action = request.form.get('action')
        student = db.query(StudentProfile).get(student_id)

        if action == 'toggle_block' and student:
            student.blocked = not student.blocked
            db.commit()

        elif action == 'delete' and student:
            user = student.user
            if user:
                db.delete(user)
            db.delete(student)
            db.commit()

        return redirect(url_for('manage_students'))

    search_query = request.args.get('search', '').strip()

    if search_query:
        students = db.query(StudentProfile).filter(
            or_(
                StudentProfile.full_name.ilike(f'%{search_query}%'),
                StudentProfile.exam_type.ilike(f'%{search_query}%'),
                StudentProfile.admission_number.ilike(f'%{search_query}%'),
                StudentProfile.phone_number.ilike(f'%{search_query}%'),
                StudentProfile.course.has(name=search_query)
            )
        ).all()
    else:
        students = db.query(StudentProfile).all()

    return render_template('admin/manage_students.html', students=students, search_query=search_query)


from sqlalchemy.orm import joinedload

@app.route('/show_credentials')
def show_credentials():
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        # ✅ Use the correct relationship name
        users = db.query(User).options(joinedload(User.profile)).all()
    finally:
        db.close()

    return render_template('admin/show_credentials.html', users=users)
'''
admin delete user
'''
@app.route('/delete_user/<int:user_id>', methods=['POST'])
def delete_user(user_id):
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        user = db.query(User).filter(User.id == user_id).first()
        if user:
            db.delete(user)
            db.commit()
    finally:
        db.close()

    return redirect(url_for('show_credentials'))

from sqlalchemy.orm import joinedload
@app.route('/admin/view_results', methods=['GET', 'POST'])
def view_results():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash("Admin access required", "danger")
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        courses = db.query(Course).all()
        subjects = db.query(Subject).all()

        selected_course = request.form.get('course')
        selected_subject = request.form.get('subject')
        export = request.form.get('export')

        query = db.query(Result).join(Result.quiz).join(Quiz.course).join(Quiz.subject).join(Result.student)

        if selected_course:
            query = query.filter(Quiz.course_id == int(selected_course))
        if selected_subject:
            query = query.filter(Quiz.subject_id == int(selected_subject))

        # Use joinedload for eager loading StudentProfile along with results
        query = query.options(joinedload(Result.student))

        results = query.all()

        if export == 'true':
            data = []
            for r in results:
                data.append({
                    'Student Username': r.student.user.username if r.student and r.student.user else 'N/A',
                    'Full Name': r.student.full_name if r.student else 'N/A',
                    'Course': r.quiz.course.name if r.quiz.course else 'N/A',
                    'Subject': r.quiz.subject.name if r.quiz.subject else 'N/A',
                    'Quiz Title': r.quiz.title,
                    'Score': r.score,
                    'Total Marks': r.total_marks,
                    'Percentage': r.percentage,
                    'Taken On': r.taken_on.strftime("%Y-%m-%d %H:%M:%S")
                })

            df = pd.DataFrame(data)
            excel_path = os.path.join(EXAMS_UPLOAD_FOLDER, 'quiz_results.xlsx')
            df.to_excel(excel_path, index=False)

            return send_file(
                excel_path,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='quiz_results.xlsx'
            )

        return render_template('admin/view_results.html',
                               results=results,
                               courses=courses,
                               subjects=subjects,
                               selected_course=selected_course,
                               selected_subject=selected_subject)

    finally:
        db.close()



'''
admin manage quizzes
'''
@app.route('/admin/manage_quizzes', methods=['GET', 'POST'])
def manage_quizzes():
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        if request.method == 'POST':
            quiz_id = int(request.form.get('quiz_id'))
            action = request.form.get('action')

            quiz = db.query(Quiz).filter_by(id=quiz_id).first()
            if not quiz:
                flash("Quiz not found.", "danger")
                return redirect(url_for('manage_quizzes'))

            if action == 'activate':
                quiz.status = 'active'
            elif action == 'deactivate':
                quiz.status = 'inactive'

            db.commit()
            flash(f"Quiz '{quiz.title}' has been {quiz.status}.", "success")
            return redirect(url_for('manage_quizzes'))

        quizzes = db.query(Quiz).all()
        return render_template('admin/manage_quizzes.html', quizzes=quizzes)
    finally:
        db.close()

'''
admin delete quiz
'''

@app.route('/admin/delete_quiz/<int:quiz_id>', methods=['POST'])
def delete_quiz(quiz_id):
    if 'role' not in session or session['role'] != 'admin':
        return redirect(url_for('login'))
    db = SessionLocal()
    quiz = db.query(Quiz).filter_by(id=quiz_id).first()
    if quiz:
        db.delete(quiz)
        db.commit()
        flash('Quiz deleted successfully.', 'success')
    else:
        flash('Quiz not found.', 'danger')
    return redirect(url_for('manage_quizzes')) 

'''
admin confirm and delete exams
'''
@app.route('/admin/review_quiz/<int:quiz_id>', methods=['GET', 'POST'])
def review_uploaded_quiz(quiz_id):
    if session.get('role') != 'admin':
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        quiz = db.query(Quiz).filter_by(id=quiz_id).first()
        if not quiz:
            flash("Quiz not found.", "danger")
            return redirect(url_for('upload_exam'))

        if request.method == 'POST':
            action = request.form.get('action')
            if action == 'delete':
                db.query(Question).filter_by(quiz_id=quiz.id).delete()
                db.delete(quiz)
                db.commit()
                flash("❌ Quiz deleted.", "warning")
                return redirect(url_for('upload_exam'))

            elif action == 'confirm':
                flash("✅ Quiz confirmed and saved.", "success")
                return redirect(url_for('upload_exam'))

        questions = db.query(Question).filter_by(quiz_id=quiz.id).all()
        return render_template('admin/review_uploaded_quiz.html', quiz=quiz, questions=questions)
    finally:
        db.close()


'''
admin toggle document active/inactive
'''
@app.route("/admin/document/<int:doc_id>/toggle")
def toggle_document_status(doc_id):
    db = SessionLocal()
    try:
        document = db.query(Document).filter_by(id=doc_id).first()
        if document:
            document.is_active = not document.is_active
            db.commit()
            flash("✅ Document status updated.", "success")
        else:
            flash("❌ Document not found.", "danger")
    finally:
        db.close()
    return redirect(url_for("list_documents"))

'''
admin view uploaded video
'''
@app.route('/admin/videos')
def list_videos():
    if session.get('role') != 'admin':
        return redirect(url_for('login'))
    db = SessionLocal()
    try:
        videos = db.query(Video).all()
        return render_template('admin/list_videos.html', videos=videos)
    finally:
        db.close()


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

            msg = Message(
                content=content,
                target_type=target_type,
                course_id=course_id if target_type == "course" else None,
                subject_id=subject_id if target_type == "subject" else None,
            )
            db.add(msg)
            db.commit()
            flash("Message created successfully!", "success")
            return redirect(url_for("admin_messages"))

        courses = db.query(Course).all()
        subjects = db.query(Subject).all()
        messages = db.query(Message).order_by(Message.created_at.desc()).all()

        return render_template("admin/messages.html", courses=courses, subjects=subjects, messages=messages)
    finally:
        db.close()

'''
delete message route
'''
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

'''
manage activity
'''
from collections import namedtuple

@app.route('/admin/student_activity')
def student_activity():
    if session.get("role") != "admin":
        flash("Please log in as admin.", "error")
        return redirect(url_for("login"))

    db = SessionLocal()
    try:
        total_students = db.query(StudentProfile).count()

        total_watching_video = db.query(ActivityLog).filter(
            ActivityLog.activity_type == "video", ActivityLog.is_active == True
        ).count()

        total_reading_document = db.query(ActivityLog).filter(
            ActivityLog.activity_type == "document", ActivityLog.is_active == True
        ).count()

        total_doing_exam = db.query(ActivityLog).filter(
            ActivityLog.activity_type == "exam", ActivityLog.is_active == True
        ).count()

        StudentActivity = namedtuple("StudentActivity", ["full_name", "course_name", "activity_type"])

        rows = (
            db.query(StudentProfile.full_name, Course.name.label("course_name"), ActivityLog.activity_type)
            .join(ActivityLog, ActivityLog.student_id == StudentProfile.id)
            .join(Course, StudentProfile.course_id == Course.id)
            .filter(ActivityLog.is_active == True)
            .all()
        )

        active_students = [StudentActivity(*row) for row in rows]

        return render_template(
            "admin/student_activity.html",
            total_students=total_students,
            total_watching_video=total_watching_video,
            total_reading_document=total_reading_document,
            total_doing_exam=total_doing_exam,
            active_students=active_students,
        )
    finally:
        db.close()

# -----------------------------
# Student functionality
# -----------------------------

from sqlalchemy.exc import IntegrityError

# -----------------------------
# Student Registration
# -----------------------------
@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username'].strip()
        password = request.form['password']

        db = SessionLocal()
        try:
            # ✅ Pre-check username
            if db.query(User).filter_by(username=username).first():
                flash('❌ Username already exists. Please choose another one.', 'danger')
                return render_template('students/Register.html', username=username)

            # Add new user
            user = User(username=username, password=password)
            db.add(user)
            db.commit()

            flash('✅ Registration successful! Please log in.', 'success')
            return redirect(url_for('login'))

        except Exception:
            db.rollback()
            flash('❌ Something went wrong. Try again.', 'danger')
            return render_template('students/Register.html', username=username)

        finally:
            db.close()

    return render_template('students/Register.html')


# -----------------------------
# Complete Student Profile
# -----------------------------
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
            # ✅ Pre-check admission_number
            if db.query(StudentProfile).filter_by(admission_number=admission_number).first():
                flash("❌ Admission number already exists. Please use another one.", "danger")
                courses = db.query(Course).all()
                return render_template("students/complete_profile.html",
                                       courses=courses,
                                       full_name=full_name,
                                       exam_type=exam_type,
                                       admission_number=admission_number,
                                       phone_number=phone_number)

            # ✅ Ensure user doesn’t already have a profile
            if db.query(StudentProfile).filter_by(user_id=user_id).first():
                flash("❌ You have already completed your profile.", "warning")
                return redirect(url_for("student_dashboard"))

            # Create profile
            profile = StudentProfile(
                full_name=full_name,
                exam_type=exam_type,
                course_id=course_id,
                admission_number=admission_number,
                phone_number=phone_number,
                user_id=user_id,
            )
            db.add(profile)
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




#student dashboard
''''student dashboard'''

@app.route('/student/dashboard')
def student_dashboard():
    if 'username' not in session or session.get('role') != 'student':
        flash('Please log in as a student first.', 'error')
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        user = db.query(User).filter_by(username=session['username']).first()
        if not user:
            flash("User not found.", "error")
            return redirect(url_for('logout'))

        student_profile = db.query(StudentProfile).filter_by(user_id=user.id).first()
        if not student_profile:
            flash("Complete your profile before proceeding.", "warning")
            return redirect(url_for('complete_profile'))

       
        available_quizzes = db.query(Quiz).filter_by(
            course_id=student_profile.course_id,
            status='active'
        ).all()

        
        taken_quiz_ids = set(
            r.quiz_id for r in db.query(Result).filter_by(student_id=user.id).all()
        )

      
        available_videos = db.query(Video).filter_by(course_id=student_profile.course_id).all()
        available_documents = db.query(Document).filter_by(course_id=student_profile.course_id).all()

        
        messages = db.query(Message).filter(
            (Message.target_type == "all") |
            ((Message.target_type == "course") & (Message.course_id == student_profile.course_id))
        ).order_by(Message.created_at.desc()).all()

        return render_template(
            'students/student_dashboard.html',
            username=user.username,
            quizzes=available_quizzes,
            videos=available_videos,
            documents=available_documents,
            taken_quiz_ids=taken_quiz_ids,
            student_profile=student_profile,
            messages_from_admin=messages
        )
    finally:
        db.close()

'''
student documents and videos
'''
@app.route('/view_document/<int:document_id>')
def view_document(document_id):
    if 'user_id' not in session:
        flash("Please log in first.", "error")
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        student = db.query(StudentProfile).filter_by(user_id=session['user_id']).first()
        if student and student.blocked:
            flash("You can't access this material. Please clear the fee to regain access.", "danger")
            return redirect(url_for('student_dashboard'))

        document = db.query(Document).filter_by(id=document_id).first()
        if not document:
            flash("Document not found.", "error")
            return redirect(url_for('complete_profile'))

        if student:
            db.query(ActivityLog).filter_by(student_id=student.id, is_active=True).update({"is_active": False})
            log = ActivityLog(student_id=student.id, activity_type="document", is_active=True)
            db.add(log)
            db.commit()

        document_path = os.path.join(DOCUMENTS_UPLOAD_FOLDER, document.filename)

        file_extension = document.filename.split('.')[-1].lower()
        if file_extension == 'pdf':
            return render_template('students/view_document.html', document=document, is_pdf=True)
        elif file_extension in ['docx', 'doc', 'pptx', 'ppt']:
            return render_template('students/view_document.html', document=document, is_pdf=False)
        else:
            return "Unsupported file type", 404

    finally:
        db.close()


'''student watch video'''
@app.route("/watch_video/<int:video_id>")
def watch_video(video_id):
    db = SessionLocal()
    
    
    user_id = session.get("user_id")
    if not user_id:
        flash("Please log in first.")
        return redirect(url_for("login"))
    
    student = db.query(StudentProfile).filter_by(user_id=user_id).first()
    if not student:
        flash("Profile not found.")
        return redirect(url_for("login"))
    
    if student.blocked:
        flash("You are blocked from accessing this content.")
        return redirect(url_for("dashboard"))
    
    
    video = db.query(Video).filter_by(id=video_id).first()
    if not video:
        flash("Video not found.")
        return redirect(url_for("dashboard"))
    
    
    file_id = extract_drive_id(video.drive_url)
    if not file_id:
        flash("Invalid video link.")
        return redirect(url_for("dashboard"))
    
    embed_url = get_drive_embed_url(file_id)
    
   
    db.query(ActivityLog).filter_by(student_id=student.id, activity_type="video", is_active=True).update({"is_active": False})
    
    new_activity = ActivityLog(
        student_id=student.id,
        activity_type="video",
        is_active=True,
        started_at=datetime.utcnow()
    )
    db.add(new_activity)
    db.commit()
    
    
    return render_template("students/watch_video.html", video=video, embed_url=embed_url)



'''
Video stream app
'''
@app.route('/video_stream/<filename>')
def video_stream(filename):
    path = os.path.join(VIDEOS_UPLOAD_FOLDER, filename)
    file_size = os.path.getsize(path)
    range_header = request.headers.get("Range", None)

    if range_header:
        byte1, byte2 = 0, None
        m = re.search(r"(\d+)-(\d*)", range_header)
        if m:
            g = m.groups()
            if g[0]:
                byte1 = int(g[0])
            if g[1]:
                byte2 = int(g[1])
        length = (byte2 or file_size - 1) - byte1 + 1
        with open(path, "rb") as f:
            f.seek(byte1)
            data = f.read(length)
        rv = Response(
            data,
            206,
            mimetype="video/mp4",
            content_type="video/mp4",
            direct_passthrough=True
        )
        rv.headers.add("Content-Range", f"bytes {byte1}-{byte1+length-1}/{file_size}")
        return rv

    return send_file(path, mimetype="video/mp4")

'''
student take exam
'''
@app.route('/take_exam/<int:quiz_id>', methods=["GET", "POST"])
def take_exam(quiz_id):
    if 'user_id' not in session:
        flash("Please log in first.", "warning")
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        
        student = db.query(StudentProfile).filter_by(user_id=session['user_id']).first()
        if not student:
            flash("Student profile not found.", "danger")
            return redirect(url_for('student_dashboard'))

        
        if student.blocked:
            flash("You can't access this exam. Please clear the fee to regain access.", "danger")
            return redirect(url_for('student_dashboard'))

       
        quiz = db.query(Quiz).filter_by(id=quiz_id).first()
        if not quiz:
            flash("Quiz not found.", "danger")
            return redirect(url_for('student_dashboard'))

        
        db.query(ActivityLog).filter_by(student_id=student.id, is_active=True).update({"is_active": False})
        log = ActivityLog(student_id=student.id, activity_type="exam", is_active=True)
        db.add(log)
        db.commit()

       
        if request.method == "POST":
            score = 0
            total_marks = 0
            questions = db.query(Question).filter_by(quiz_id=quiz.id).all()

            for q in questions:
                selected = request.form.get(f"question_{q.id}") 
                total_marks += q.marks  
                if selected and selected == q.correct_option: 
                    score += q.marks  

            percentage = (score / total_marks) * 100 if total_marks > 0 else 0

            # ✅ save result
            result = Result(
                student_id=student.id,
                quiz_id=quiz.id,
                score=score,
                total_marks=total_marks,
                percentage=percentage
            )
            db.add(result)

            
            db.query(ActivityLog).filter_by(student_id=student.id, activity_type="exam", is_active=True).update({"is_active": False})
            db.commit()

            flash(f"You scored {score} out of {total_marks} ({percentage:.2f}%)", "success")
            return redirect(url_for('student_dashboard'))

        # ✅ render exam page
        questions = db.query(Question).filter_by(quiz_id=quiz.id).all()
        return render_template("students/take_exam.html", quiz=quiz, questions=questions)

    finally:
        db.close()

@app.route('/student/results')
def student_results():
    if 'username' not in session or session.get('role') != 'student':
        flash('Please log in as a student first.', 'error')
        return redirect(url_for('login'))

    db = SessionLocal()
    try:
        # ✅ get user
        user = db.query(User).filter_by(username=session['username']).first()
        if not user:
            flash("User not found.", "error")
            return redirect(url_for('logout'))

        # ✅ get student profile
        student = db.query(StudentProfile).filter_by(user_id=user.id).first()
        if not student:
            flash("Student profile not found.", "error")
            return redirect(url_for('logout'))

        # ✅ get results linked to that student
        results = (
            db.query(Result)
              .join(Quiz, Result.quiz_id == Quiz.id)
              .filter(Result.student_id == student.id)   # ✅ correct now
              .order_by(Result.taken_on.desc())
              .all()
        )

        return render_template('students/results.html', results=results)

    finally:
        db.close()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
