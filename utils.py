import docx
import re

# ------------------------------
# Internal helpers
# ------------------------------

def load_docx(path: str):
    """Load all paragraphs from a DOCX file."""
    doc = docx.Document(path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def is_question_line(text: str):
    """Detect question numbering like: 1.  or 1)"""
    return re.match(r"^\d+[\.\)]\s*", text) is not None


def is_option_line(text: str):
    """Detect options A-D"""
    return re.match(r"^[A-D][\.\):]\s*", text, re.IGNORECASE) is not None


def is_answer_line(text: str):
    """Detect 'Answer: X' lines"""
    return text.lower().startswith("answer")


# ------------------------------
# MCQ Parsing
# ------------------------------

def parse_mcqs(paragraphs):
    mcqs = []
    current = None

    for line in paragraphs:
        if is_question_line(line):
            if current:
                mcqs.append(current)
            current = {"question": line, "options": {}, "answer": None}

        elif is_option_line(line) and current:
            letter = line[0].upper()
            text = re.sub(r"^[A-D][\.\):]\s*", "", line)
            current["options"][letter] = text.strip()

        elif is_answer_line(line) and current:
            current["answer"] = line.split(":")[-1].strip().upper()

    if current:
        mcqs.append(current)

    return mcqs


# ------------------------------
# Case Study Parsing
# ------------------------------

def parse_case_studies(paragraphs):
    case_studies = []
    current = []
    is_case = False

    for line in paragraphs:
        if "case study" in line.lower():
            if current:
                case_studies.append("\n".join(current))
            current = [line]
            is_case = True

        elif is_case:
            if is_question_line(line):  # end of case study
                case_studies.append("\n".join(current))
                is_case = False
                current = []
                continue
            current.append(line)

    if current and is_case:
        case_studies.append("\n".join(current))

    return case_studies


# ------------------------------
# Main function your app.py expects
# ------------------------------

def parse_docx_questions(path: str):
    """
    Required by app.py
    Returns:
        {
            "mcqs": [...],
            "case_studies": [...],
        }
    """
    paragraphs = load_docx(path)
    return {
        "mcqs": parse_mcqs(paragraphs),
        "case_studies": parse_case_studies(paragraphs)
    }


# ------------------------------
# Dummy function for app.py
# ------------------------------

def get_quiz_status():
    """
    Some projects check quiz progress or availability.
    If you don't need anything complicated,
    return a simple static status dictionary.
    """
    return {
        "status": "ok",
        "message": "Quiz system operational"
    }
