import re
import os
from docx import Document

# --------------------------------------------
# GOOGLE DRIVE HELPERS
# --------------------------------------------

def extract_drive_id(url: str):
    """
    Extracts the file ID from a Google Drive share URL.
    """
    patterns = [
        r"/d/([a-zA-Z0-9_-]+)",              # /d/FILE_ID/
        r"id=([a-zA-Z0-9_-]+)",             # ?id=FILE_ID
        r"file/d/([a-zA-Z0-9_-]+)"          # file/d/FILE_ID
    ]

    for p in patterns:
        match = re.search(p, url)
        if match:
            return match.group(1)

    return None


def get_drive_embed_url(file_id: str):
    """
    Returns a Google Drive embeddable video link.
    """
    return f"https://drive.google.com/file/d/{file_id}/preview"



# --------------------------------------------
# QUIZ STATUS HELPER
# --------------------------------------------

def get_quiz_status(quiz):
    if quiz.status == "active":
        return "Active"
    else:
        return "Inactive"



# --------------------------------------------
# .DOCX QUESTION PARSING ENGINE
# --------------------------------------------

def parse_docx_questions(filepath, image_output_dir=None):
    """
    A UNIVERSAL parser that supports:
    -------------------------------------
    1️⃣ Multiple-choice questions  
    2️⃣ Case study questions  
    3️⃣ Accounting structured questions  

    Returns a list of dictionaries:
    {
        "question": "...",
        "a": "...",
        "b": "...",
        "c": "...",
        "d": "...",
        "answer": "...",
        "marks": 2,
        "extra_content": "...",   # for long questions
        "image": None
    }
    """

    doc = Document(filepath)
    text_lines = []

    for para in doc.paragraphs:
        line = para.text.strip()
        if line:
            text_lines.append(line)

    questions = []
    current_question = None
    options_collected = 0

    mcq_pattern = r"^\s*(\d+)\.\s*(.+?)\((\d+)\s*mks?\)$"   # e.g. 1. What is...? (2mks)

    option_pattern = r"^[A-D]\.?[\)]?\s*(.+)$"             # A.answer, A) answer, A. answer

    for line in text_lines:

        # -------------------------------
        # 1️⃣ Detect MULTIPLE CHOICE QUESTION
        # -------------------------------
        q_match = re.match(mcq_pattern, line, re.IGNORECASE)
        if q_match:
            # Save previous question
            if current_question:
                questions.append(current_question)

            q_number, q_text, marks = q_match.groups()

            current_question = {
                "question": q_text.strip(),
                "marks": int(marks),
                "a": "",
                "b": "",
                "c": "",
                "d": "",
                "answer": "",
                "extra_content": None,
                "image": None
            }
            options_collected = 0
            continue

        # -------------------------------
        # 2️⃣ Detect OPTIONS A/B/C/D
        # -------------------------------
        if current_question:
            if re.match(r"^A", line):
                current_question["a"] = clean_option(line)
                options_collected += 1
                continue
            if re.match(r"^B", line):
                current_question["b"] = clean_option(line)
                options_collected += 1
                continue
            if re.match(r"^C", line):
                current_question["c"] = clean_option(line)
                options_collected += 1
                continue
            if re.match(r"^D", line):
                current_question["d"] = clean_option(line)
                options_collected += 1
                continue

        # -------------------------------
        # 3️⃣ Detect ANSWER: X
        # -------------------------------
        if current_question and line.lower().startswith("answer"):
            ans = line.split(":")[-1].strip()
            current_question["answer"] = ans.lower()
            continue

        # -------------------------------
        # 4️⃣ Case study OR Accounting questions
        # (Long paragraph before any “Answer:”)
        # -------------------------------
        if current_question and current_question["answer"] == "":
            # Append long text into extra_content
            if current_question.get("extra_content") is None:
                current_question["extra_content"] = line
            else:
                current_question["extra_content"] += "\n" + line

    # Add last question
    if current_question:
        questions.append(current_question)

    return questions



# --------------------------------------------
# CLEAN OPTIONS LIKE:
# A.answer:A  -> A
# A) answer   -> answer
# --------------------------------------------

def clean_option(text):
    """
    Extracts clean option text by removing A., A), A-answer etc.
    """
    text = text.strip()
    text = re.sub(r"^[A-D][\.\:\)\-]*\s*", "", text)
    return text
