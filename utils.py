import docx
import re
import os
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# ---------------------------------------------------------
# LOAD DOCX
# ---------------------------------------------------------
def load_docx(path):
    """Load a DOCX file."""
    return docx.Document(path)

# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------
def is_question_line(text):
    return re.match(r"^\d+[\.\)]\s*", text) is not None

def is_option_line(text):
    return re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE) is not None

def is_answer_line(text):
    return text.lower().startswith(("answer", "ans", "correct"))

# ---------------------------------------------------------
# IMAGE EXTRACTION
# ---------------------------------------------------------
def extract_images(document, output_dir, q_index):
    """Extract images from DOCX and save them in the specified folder."""
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
# PARSER
# ---------------------------------------------------------
def parse_docx_questions(path, image_output_dir=None):
    """Parse DOCX file into a list of questions with options, answer, marks, and optional image."""
    doc = load_docx(path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    questions = []
    current = None
    q_index = 0

    for line in paragraphs:
        # -------------------------- NEW QUESTION --------------------------
        if is_question_line(line):
            if current:
                questions.append(current)
            q_index += 1
            q_text = re.sub(r"^\d+[\.\)]\s*", "", line).strip()
            mk = re.search(r"\((\d+)\s*mks?\)", line, re.IGNORECASE)
            marks = int(mk.group(1)) if mk else 1
            current = {
                "question": q_text,
                "a": "",
                "b": "",
                "c": "",
                "d": "",
                "answer": "",
                "marks": marks,
                "image": None
            }
            if image_output_dir:
                imgs = extract_images(doc, image_output_dir, q_index)
                if imgs:
                    current["image"] = imgs[0]

        # -------------------------- OPTIONS A–D --------------------------
        elif is_option_line(line) and current:
            letter = line[0].lower()
            text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
            current[letter] = text

        # -------------------------- ANSWER LINE --------------------------
        elif is_answer_line(line) and current:
            raw = line.split(":")[-1].strip().lower()
            clean = re.sub(r"[^a-d]", "", raw)
            current["answer"] = clean

    if current:
        questions.append(current)

    return questions

# ---------------------------------------------------------
# SCORING ENGINE
# ---------------------------------------------------------
def compute_score(questions, student_answers):
    """
    Compute the student's score.
    `student_answers` can be keyed by:
      - question index (int)
      - question text (str)
      - 'q1', 'q2', ... style
    """
    total_marks = 0
    score = 0
    details = []

    for index, q in enumerate(questions, start=1):
        total_marks += q["marks"]
        correct = q["answer"].lower().strip()

        # Try multiple ways to find the student's answer
        student_answer = ""
        for key in (index, f"q{index}", q["question"]):
            if key in student_answers:
                student_answer = student_answers[key].lower().strip()
                break

        is_correct = student_answer == correct
        if is_correct:
            score += q["marks"]

        details.append({
            "question": q["question"],
            "correct": correct,
            "student_answer": student_answer,
            "marks": q["marks"],
            "earned": q["marks"] if is_correct else 0
        })

    return {
        "score": score,
        "total": total_marks,
        "percentage": round((score / total_marks) * 100, 2) if total_marks else 0,
        "details": details
    }

# ---------------------------------------------------------
# QUIZ STATUS
# ---------------------------------------------------------
def get_quiz_status(questions, student_answers):
    """
    Returns the status of each question: correct, incorrect, or unanswered.
    Works with flexible keys (index, 'qX', or full question text).
    """
    status_list = []
    for index, q in enumerate(questions, start=1):
        student_answer = ""
        for key in (index, f"q{index}", q["question"]):
            if key in student_answers:
                student_answer = student_answers[key].strip().lower()
                break
        correct_answer = q["answer"].strip().lower()
        if not student_answer:
            status = "unanswered"
        elif student_answer == correct_answer:
            status = "correct"
        else:
            status = "incorrect"
        status_list.append({
            "question_index": index,
            "status": status,
            "student_answer": student_answer,
            "correct_answer": correct_answer
        })
    return status_list

# ---------------------------------------------------------
# GOOGLE DRIVE HELPERS
# ---------------------------------------------------------
def extract_drive_id(url):
    """
    Extract the file ID from a Google Drive URL.
    Supports:
      - https://drive.google.com/file/d/FILE_ID/view?usp=sharing
      - https://drive.google.com/open?id=FILE_ID
      - FILE_ID directly
    """
    patterns = [
        r"https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)",
        r"https://drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return url  # assume input is already a file ID

def get_drive_embed_url(drive_url_or_id):
    """Returns an embeddable Google Drive URL."""
    file_id = extract_drive_id(drive_url_or_id)
    return f"https://drive.google.com/file/d/{file_id}/preview"
