import os
import re
from urllib.parse import urlparse, parse_qs
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

# -----------------------
# CONFIG
# -----------------------
DEFAULT_IMAGE_DIR = "static/question_images"


# -----------------------
# HELPERS
# -----------------------
def iter_block_items(parent):
    """Yield paragraphs and tables in document order."""
    for child in parent.element.body:
        if child.tag.endswith("p"):
            yield Paragraph(child, parent)
        elif child.tag.endswith("tbl"):
            yield Table(child, parent)


def save_image_from_run(run, output_dir, image_index):
    """Extract images embedded in DOCX runs."""
    if not hasattr(run, "element"):
        return None

    # Look for images inside the run
    drawing_elements = run.element.findall(".//a:blip", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
    if not drawing_elements:
        return None

    image_name = f"img_{image_index}.png"
    image_path = os.path.join(output_dir, image_name)

    for blip in drawing_elements:
        embed = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        if not embed:
            continue

        img_part = run.part.related_parts[embed]
        with open(image_path, "wb") as f:
            f.write(img_part.blob)

        return image_name

    return None


def extract_table_html(table):
    """Convert DOCX table to basic HTML."""
    html = "<table border='1' style='border-collapse: collapse;'>"
    for row in table.rows:
        html += "<tr>"
        for cell in row.cells:
            html += f"<td>{cell.text.strip()}</td>"
        html += "</tr>"
    html += "</table><br>"
    return html


# ================================================================
# ======================  PARSER FUNCTION  ========================
# ================================================================
def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
    """
    Parse case-study + MCQ questions from a .docx file.
    Correct answers are stored internally as 'correct_answer'
    but NEVER returned to students.
    """

    document = Document(file_stream)
    os.makedirs(image_output_dir, exist_ok=True)

    questions = []
    current_question = None
    extra_html_parts = []
    image_counter = 0

    for block in iter_block_items(document):

        # --------------------------------------------------
        # PARAGRAPH BLOCK
        # --------------------------------------------------
        if isinstance(block, Paragraph):
            para = block
            text = para.text.strip()

            # Extract inline images
            for run in para.runs:
                image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
                if image_name:
                    image_counter += 1
                    if current_question:
                        current_question["image"] = image_name
                    continue

            if not text:
                continue

            # ------------------------------
            # NEW QUESTION START
            # ------------------------------
            if re.match(r"^\d+[\.\)]", text):

                # Save previous question
                if current_question:
                    if extra_html_parts:
                        current_question["extra_content"] += ''.join(extra_html_parts)
                        extra_html_parts = []
                    questions.append(current_question)

                # Extract marks
                marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
                marks = int(marks_match.group(1)) if marks_match else 1

                # Clean
                clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
                question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)

                current_question = {
                    "question": question_text,
                    "a": "", "b": "", "c": "", "d": "",
                    "correct_answer": "",    # internal only
                    "extra_content": "",
                    "image": None,
                    "marks": marks
                }
                continue

            # ------------------------------
            # OPTIONS (A–D)
            # ------------------------------
            if re.match(r"^\(?[a-dA-D][\.\)]", text):
                match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
                if match and current_question:
                    label = match.group(1).lower()
                    current_question[label] = match.group(2).strip()
                continue

            # ------------------------------
            # CORRECT ANSWER
            # (stored but NOT returned)
            # ------------------------------
            if re.match(r"^(answer|correct answer):", text, re.IGNORECASE):
                m = re.search(r":\s*([a-dA-D])", text)
                if m and current_question:
                    current_question["correct_answer"] = m.group(1).lower()
                continue

            # ------------------------------
            # EXTRA CONTENT
            # ------------------------------
            if current_question:
                extra_html_parts.append(f"<p>{text}</p>")
                continue

        # --------------------------------------------------
        # TABLE BLOCK
        # --------------------------------------------------
        elif isinstance(block, Table):
            table_html = extract_table_html(block)
            if current_question:
                current_question["extra_content"] += table_html
            continue

    # --------------------------------------------------
    # SAVE LAST QUESTION
    # --------------------------------------------------
    if current_question:
        if extra_html_parts:
            current_question["extra_content"] += ''.join(extra_html_parts)
        questions.append(current_question)

    # --------------------------------------------------
    # REMOVE CORRECT ANSWERS BEFORE SENDING TO STUDENTS
    # --------------------------------------------------
    safe_questions = []
    for q in questions:
        safe = {k: v for k, v in q.items() if k != "correct_answer"}
        safe_questions.append(safe)

    return safe_questions


# ================================================================
# ================  PLACEHOLDER FOR RENDER  ======================
# ================================================================
def get_quiz_status(user_id):
    return "active"


# ================================================================
# ================  GOOGLE DRIVE HELPERS  =========================
# ================================================================
def extract_drive_id(url: str):
    if not url:
        return None

    # /file/d/<ID>/view
    m = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
    if m:
        return m.group(1)

    # ?id=<ID>
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)
    if "id" in qs:
        return qs["id"][0]

    return None


def get_drive_embed_url(file_id: str):
    return f"https://drive.google.com/file/d/{file_id}/preview"
