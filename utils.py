import docx
import re
import os
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# ---------------------------------------------------------
# LOAD DOCX
# ---------------------------------------------------------

def load_docx(path):
    doc = docx.Document(path)
    return doc


# ---------------------------------------------------------
# HELPERS TO DETECT LINES
# ---------------------------------------------------------

def is_question_line(text):
    # Example: "1. What is..." or "1) What is..."
    return re.match(r"^\d+[\.\)]\s*", text) is not None


def is_option_line(text):
    # Example: "A. Option" or "A) Option"
    return re.match(r"^[A-D][\.\):]\s*", text, re.IGNORECASE) is not None


def is_answer_line(text):
    # Example: "Answer: C"
    return text.lower().startswith("answer")


# ---------------------------------------------------------
# IMAGE EXTRACTION
# ---------------------------------------------------------

def extract_images(document, output_dir, question_index):
    """
    Extract images from a DOCX document.
    Saves them as PNG/JPG files inside output_dir.
    """
    os.makedirs(output_dir, exist_ok=True)

    images = {}
    count = 0

    for rel in document.part.rels.values():
        if rel.reltype == RT.IMAGE:
            count += 1
            ext = rel.target_ref.split('.')[-1]
            filename = f"q{question_index}_img{count}.{ext}"
            filepath = os.path.join(output_dir, filename)

            with open(filepath, "wb") as f:
                f.write(rel.target_part.blob)

            images[count] = filename

    return images


# ---------------------------------------------------------
# MAIN MCQ PARSER
# ---------------------------------------------------------

def parse_docx_questions(path, image_output_dir=None):
    """
    RETURNS a list of questions in format:
    {
        'question': 'What is...',
        'a': 'Option text',
        'b': 'Option text',
        'c': 'Option text',
        'd': 'Option text',
        'answer': 'c',
        'marks': 2,
        'image': 'filename.png' (if exists)
    }
    """

    doc = load_docx(path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    questions = []
    current = None
    q_index = 0

    for line in paragraphs:

        # ---------------------------
        # Question line
        # ---------------------------
        if is_question_line(line):
            if current:
                questions.append(current)

            q_index += 1
            question_text = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

            # Extract marks (e.g. "(2mks)")
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

            # Extract images
            if image_output_dir:
                imgs = extract_images(doc, image_output_dir, q_index)
                if imgs:
                    current["image"] = list(imgs.values())[0]

        # ---------------------------
        # Option line
        # ---------------------------
        elif is_option_line(line) and current:
            letter = line[0].lower()
            text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
            current[letter] = text

        # ---------------------------
        # Answer line
        # ---------------------------
        elif is_answer_line(line) and current:
            ans = line.split(":")[-1].strip().lower()
            current["answer"] = ans

    if current:
        questions.append(current)

    return questions


# ---------------------------------------------------------
# CASE STUDY PARSER
# ---------------------------------------------------------

def parse_case_studies(paragraphs):
    case_studies = []
    block = []
    capturing = False

    for line in paragraphs:
        if "case study" in line.lower():
            if block:
                case_studies.append("\n".join(block))
            block = [line]
            capturing = True

        elif capturing:
            if is_question_line(line):
                case_studies.append("\n".join(block))
                block = []
                capturing = False
            else:
                block.append(line)

    if capturing and block:
        case_studies.append("\n".join(block))

    return case_studies


# ---------------------------------------------------------
# QUIZ STATUS
# ---------------------------------------------------------

def get_quiz_status():
    return {"status": "ok", "message": "Quiz system operational"}


# ---------------------------------------------------------
# GOOGLE DRIVE HELPERS
# ---------------------------------------------------------

def extract_drive_id(url):
    patterns = [
        r"/d/([a-zA-Z0-9_-]+)",
        r"id=([a-zA-Z0-9_-]+)"
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None


def get_drive_embed_url(file_id):
    if not file_id:
        return None
    return f"https://drive.google.com/file/d/{file_id}/preview"
