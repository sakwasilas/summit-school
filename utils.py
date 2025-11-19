import docx
import re
import os
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# ---------------------------------------------------------
# LOAD DOCX
# ---------------------------------------------------------

def load_docx(path):
    """Load a DOCX file and return a Document object."""
    return docx.Document(path)

# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------

def is_question_line(text):
    """Check if a line is a question line (starts with a number)."""
    return re.match(r"^\d+[\.\)]\s*", text) is not None

def is_option_line(text):
    """Check if a line is an option line (A-D)."""
    return re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE) is not None

def is_answer_line(text):
    """Check if a line contains the answer."""
    return text.lower().startswith(("answer", "ans", "correct"))

# ---------------------------------------------------------
# IMAGE EXTRACTION
# ---------------------------------------------------------

def extract_images(document, output_dir, q_index):
    """Extract images from a DOCX document into a specified directory."""
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
    """
    Parse DOCX questions into structured format.
    Returns a list of dictionaries with question, options, answer, marks, and image.
    """
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
            clean = re.sub(r"[^a-d]", "", raw)  # Keep only a-d
            current["answer"] = clean

    if current:
        questions.append(current)

    return questions

# ---------------------------------------------------------
# SCORING ENGINE
# ---------------------------------------------------------

def compute_score(questions, student_answers):
    """
    Compute total score, percentage, and detailed results for a student's answers.
    student_answers: dict with {question_index: answer_letter}
    """
    total_marks = 0
    score = 0
    details = []

    for index, q in enumerate(questions, start=1):
        total_marks += q["marks"]
        correct = q["answer"].lower().strip()
        student_answer = student_answers.get(index, "").lower().strip()
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
# QUIZ STATUS FUNCTION
# ---------------------------------------------------------

def get_quiz_status(questions, student_answers):
    """
    Wrapper function to get quiz status for a student.
    Returns same as compute_score.
    """
    return compute_score(questions, student_answers)
