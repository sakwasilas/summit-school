import docx
import re

def load_docx(path: str):
    """Load all paragraphs from a docx file."""
    doc = docx.Document(path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def is_question_line(text: str):
    """Detect lines like: 1. What is...? OR 12) What is...?"""
    return re.match(r"^\d+[\.\)]\s*", text) is not None


def is_option_line(text: str):
    """Detect options such as:
       A.answer:...
       A. ...
       A) ...
    """
    return re.match(r"^[A-D][\.\):]\s*", text, re.IGNORECASE) is not None


def is_answer_line(text: str):
    """Detect the answer line: Answer: X"""
    return text.lower().startswith("answer")


def parse_mcqs(paragraphs):
    """Extract MCQs from the document."""
    mcqs = []
    current = {}

    for line in paragraphs:
        if is_question_line(line):
            if current:
                mcqs.append(current)
            current = {
                "question": line,
                "options": {},
                "answer": None
            }

        elif is_option_line(line) and current:
            key = line[0].upper()
            option_text = re.sub(r"^[A-D][\.\):]\s*", "", line)
            current["options"][key] = option_text.strip()

        elif is_answer_line(line) and current:
            ans = line.split(":")[-1].strip().upper()
            current["answer"] = ans

    if current:
        mcqs.append(current)

    return mcqs


def parse_case_studies(paragraphs):
    """Extract case study blocks (multi-paragraph). 
       A new case study begins when a paragraph contains 'Case Study' or 'Study'.
    """
    case_studies = []
    current = []
    recording = False

    for line in paragraphs:
        if "case study" in line.lower():
            if current:
                case_studies.append("\n".join(current))
            current = [line]
            recording = True

        elif recording:
            if is_question_line(line):  # stop at MCQ set
                recording = False
                continue
            current.append(line)

    if current:
        case_studies.append("\n".join(current))

    return case_studies


def parse_document(path: str):
    """Full extraction for:
       - MCQs
       - Case studies
    """
    paragraphs = load_docx(path)

    return {
        "mcqs": parse_mcqs(paragraphs),
        "case_studies": parse_case_studies(paragraphs)
    }


# ---------- Example run ----------
if __name__ == "__main__":
    data = parse_document("fof cat1 2025.docx")

    print("\n=== MCQs Found ===")
    for q in data["mcqs"]:
        print(q)

    print("\n=== CASE STUDIES ===")
    for cs in data["case_studies"]:
        print("\n", cs, "\n")
