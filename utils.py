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
    """Check if a line is a question line (e.g., '1. Question?')."""
    return re.match(r"^\d+[\.\)]\s*", text) is not None

def is_option_line(text):
    """Check if a line is an option (A–D)."""
    return re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE) is not None

def is_answer_line(text):
    """Check if a line contains the answer."""
    return text.lower().startswith(("answer", "ans", "correct"))

# ---------------------------------------------------------
# IMAGE EXTRACTION
# ---------------------------------------------------------
def extract_images(document, output_dir, q_index):
    """Extract all images from a DOCX and save them to output_dir."""
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

    return images if images else None

# ---------------------------------------------------------
# PARSER
# ---------------------------------------------------------
def parse_docx_questions(path, image_output_dir=None):
    """Parse questions, options, answers, marks, and images from DOCX."""
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

            # Extract marks if provided
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
                "images": None
            }

            # Extract images if directory is provided
            if image_output_dir:
                imgs = extract_images(doc, image_output_dir, q_index)
                if imgs:
                    current["images"] = imgs

        # -------------------------- OPTIONS A–D --------------------------
        elif is_option_line(line) and current:
            letter = line[0].lower()
            text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
            current[letter] = text

        # -------------------------- ANSWER LINE --------------------------
        elif is_answer_line(line) and current:
            raw = line.split(":")[-1].strip().lower()
            # Keep only letters A–D, allow multiple answers separated by comma
            clean = ",".join(re.findall(r"[a-d]", raw))
            current["answer"] = clean

    if current:
        questions.append(current)

    return questions

# ---------------------------------------------------------
# SCORING ENGINE
# ---------------------------------------------------------
def compute_score(questions, student_answers):
    """
    Compute student score.

    `questions`: list from parse_docx_questions
    `student_answers`: dict {question_index: "A" or "A,C"}
    """
    total_marks = 0
    score = 0
    details = []

    for index, q in enumerate(questions, start=1):
        total_marks += q["marks"]
        correct_answers = set(q["answer"].split(","))  # handle multiple answers
        student_answer = set(student_answers.get(index, "").lower().replace(" ", "").split(","))

        is_correct = student_answer == correct_answers

        if is_correct:
            score += q["marks"]

        details.append({
            "question": q["question"],
            "correct": ",".join(correct_answers),
            "student_answer": ",".join(student_answer),
            "marks": q["marks"],
            "earned": q["marks"] if is_correct else 0,
            "images": q.get("images")
        })

    percentage = round((score / total_marks) * 100, 2) if total_marks else 0

    return {
        "score": score,
        "total": total_marks,
        "percentage": percentage,
        "details": details
    }
