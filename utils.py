import docx
import re
import os
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# ---------------------------------------------------------
# LOAD DOCX
# ---------------------------------------------------------
def load_docx(path):
    return docx.Document(path)

# ---------------------------------------------------------
# HELPER MATCHERS
# ---------------------------------------------------------
def is_question_line(text):
    return bool(re.match(r"^\d+[\.\)]\s*", text))

def is_option_line(text):
    return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

def is_answer_line(text):
    return text.lower().startswith(("answer", "ans", "correct"))

# ---------------------------------------------------------
# IMAGE EXTRACTION
# ---------------------------------------------------------
def extract_images(document, output_dir, q_index):
    os.makedirs(output_dir, exist_ok=True)
    images = []
    count = 0

    for rel in document.part.rels.values():
        if rel.reltype == RT.IMAGE:
            count += 1
            ext = rel.target_ref.split('.')[-1]
            filename = f"q{q_index}_img{count}.{ext}"
            filepath = os.path.join(output_dir, filename)

            with open(filepath, "wb") as f:
                f.write(rel.target_part.blob)

            images.append(filename)

    return images

# ---------------------------------------------------------
# DOCX PARSER (MAIN FIXES HERE)
# ---------------------------------------------------------
def parse_docx_questions(path, image_output_dir=None):
    doc = load_docx(path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    questions = []
    current = None
    q_index = 0

    for line in paragraphs:

        # ---------- NEW QUESTION ----------
        if is_question_line(line):
            if current:
                questions.append(current)

            q_index += 1
            question_text = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

            # extract marks like (2 mks)
            mk = re.search(r"\((\d+)\s*mks?\)", line, re.IGNORECASE)
            marks = int(mk.group(1)) if mk else 1

            current = {
                "question": question_text,
                "a": "",
                "b": "",
                "c": "",
                "d": "",
                "answer": "",
                "marks": marks,
                "image": None
            }

            # extract image for only this question
            if image_output_dir:
                imgs = extract_images(doc, image_output_dir, q_index)
                if imgs:
                    current["image"] = imgs[0]

        # ---------- OPTION A-D ----------
        elif is_option_line(line) and current:
            letter = line[0].lower()
            text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
            current[letter] = text

        # ---------- ANSWER ----------
        elif is_answer_line(line) and current:
            raw = line.split(":")[-1].strip().lower()
            clean = re.sub(r"[^a-d]", "", raw)
            current["answer"] = clean

    if current:
        questions.append(current)

    return questions

# ---------------------------------------------------------
# SCORING ENGINE (BIG FIX HERE)
# ---------------------------------------------------------
def compute_score(questions, student_answers):
    score = 0
    total_marks = 0
    details = []

    for index, q in enumerate(questions, start=1):
        correct = q["answer"].strip().lower()
        total_marks += q["marks"]

        # get student's answer safely
        student_answer = ""
        for key in (index, f"q{index}", q["question"]):
            if key in student_answers:
                student_answer = student_answers[key].strip().lower()
                break

        got_it = student_answer == correct
        if got_it:
            score += q["marks"]

        details.append({
            "question": q["question"],
            "correct": correct,
            "student_answer": student_answer,
            "marks": q["marks"],
            "earned": q["marks"] if got_it else 0
        })

    percentage = round((score / total_marks) * 100, 2) if total_marks else 0

    return {
        "score": score,
        "total": total_marks,
        "percentage": percentage,
        "details": details
    }

# ---------------------------------------------------------
# QUIZ STATUS
# ---------------------------------------------------------
def get_quiz_status(questions, student_answers):
    status_list = []

    for index, q in enumerate(questions, start=1):
        correct = q["answer"].strip().lower()

        student_answer = ""
        for key in (index, f"q{index}", q["question"]):
            if key in student_answers:
                student_answer = student_answers[key].strip().lower()
                break

        if not student_answer:
            status = "unanswered"
        elif student_answer == correct:
            status = "correct"
        else:
            status = "incorrect"

        status_list.append({
            "question_index": index,
            "status": status,
            "student_answer": student_answer,
            "correct_answer": correct
        })

    return status_list

# ---------------------------------------------------------
# GOOGLE DRIVE HELPERS
# ---------------------------------------------------------
def extract_drive_id(url):
    patterns = [
        r"https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)",
        r"https://drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)"
    ]
    for pattern in patterns:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return url

def get_drive_embed_url(drive_url_or_id):
    file_id = extract_drive_id(drive_url_or_id)
    return f"https://drive.google.com/file/d/{file_id}/preview"
