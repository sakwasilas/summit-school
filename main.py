from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
import os
import random
import string
import io
import csv
import docx

app = Flask(__name__)
app.secret_key = 'secretkey'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['UPLOAD_FOLDER'] = 'uploads'

ALLOWED_EXTENSIONS = {'mp4', 'avi', 'mov', 'mkv'}
db = SQLAlchemy(app)

# ----------------------------------
# Shared Routes
# ----------------------------------
@app.route('/')
def index():
    return redirect(url_for('login'))

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

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))
