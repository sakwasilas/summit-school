import docx
import re
import os
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# ------------------- IMAGE EXTRACTION (FIXED - NO DUPLICATION) -----------------
# Global cache to store extracted images per document
_image_cache = {}

def extract_images(document, output_dir, q_index):
    """
    Extract images from document - but only once per document.
    Returns the image for the specific question index.
    """
    global _image_cache
    
    # Create a cache key based on document and output directory
    cache_key = f"{id(document)}_{output_dir}"
    
    # If we haven't extracted images for this document yet, do it once
    if cache_key not in _image_cache:
        os.makedirs(output_dir, exist_ok=True)
        images = []
        count = 0
        
        # Extract all images and store them in order
        for rel in document.part.rels.values():
            if rel.reltype == RT.IMAGE:
                count += 1
                ext = rel.target_ref.split('.')[-1]
                # Use generic names without question index
                filename = f"img_{count}.{ext}"
                filepath = os.path.join(output_dir, filename)
                with open(filepath, "wb") as f:
                    f.write(rel.target_part.blob)
                images.append(filename)
        
        _image_cache[cache_key] = images
    
    # Return the image for this question index (1-based)
    all_images = _image_cache[cache_key]
    if q_index <= len(all_images):
        return [all_images[q_index - 1]]  # Return as list to maintain compatibility
    else:
        return []  # No image for this question

# --------------------- LOAD DOCX ---------------------
def load_docx(path):
    return docx.Document(path)

# --------------------- HELPERS -----------------------
def is_question_line(text):
    return bool(re.match(r"^\d+[\.\)]\s*", text))

def is_option_line(text):
    return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

def is_answer_line(text):
    return text.lower().lstrip().startswith(("answer", "ans", "correct"))

def is_case_study_line(text):
    keywords = [
        "use the following information",
        "study the information",
        "refer to the following",
        "case study",
        "use the data below"
    ]
    t = text.lower()
    return any(k in t for k in keywords)

# ------------------- TABLE -> HTML -------------------
def make_html_table(cells):
    html = "<table class='table table-bordered'>"
    for row in cells:
        html += "<tr>"
        for c in row:
            # Escape minimal HTML-sensitive chars (basic)
            safe = (c.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
            html += f"<td>{safe}</td>"
        html += "</tr>"
    html += "</table>"
    return html

# ------------------- FLATTEN DOCX --------------------
def flatten_doc(document):
    """
    Returns a list of entries preserving paragraphs and tables in document order.
    Each entry is {"type":"text","content":...} or {"type":"table","cells": [...]}
    """
    lines = []
    for block in document.element.body:
        if block.tag.endswith('p'):
            para = docx.text.paragraph.Paragraph(block, document)
            text = para.text.strip()
            if text:
                lines.append({"type": "text", "content": text})
        elif block.tag.endswith('tbl'):
            table = docx.table.Table(block, document)
            rows = []
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                rows.append(cells)
            lines.append({"type": "table", "cells": rows})
    return lines

# --------------- PARSE DOCX QUESTIONS -----------------
def parse_docx_questions(path, image_output_dir=None):
    """
    Parse docx file and return list of questions.
    Each question dict has:
      - question (text with any inline tables as HTML)
      - instructions (case study/instructions that were appearing BEFORE the question)
      - a, b, c, d (option texts)
      - answer (single letter 'a'..'d', stored lower-case)
      - marks (int)
      - image (filename or None)
    Case studies/tables that appear before a question are attached to the next question's 'instructions'.
    """
    doc = load_docx(path)
    entries = flatten_doc(doc)

    questions = []
    current = None
    q_index = 0
    # holds instructions/case-study found before the *next* question
    pending_instructions = ""

    for entry in entries:
        if entry["type"] == "text":
            line = entry["content"].strip()

            # If line contains both an Answer and a new case-study after it (same paragraph),
            # we'll handle splitting later when looking for answer lines.
            # ---------- NEW QUESTION ----------
            if is_question_line(line):
                # push previous
                if current:
                    questions.append(current)

                q_index += 1
                raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

                # detect inline embedded case study text in same line as question (rare)
                embedded_case = ""
                for k in [
                    "use the following information",
                    "study the information",
                    "refer to the following",
                    "case study",
                    "use the data below"
                ]:
                    if k in raw_question.lower():
                        parts = re.split(k, raw_question, flags=re.IGNORECASE)
                        raw_question = parts[0].strip()
                        embedded_case = k + " " + parts[1].strip()
                        break

                # extract marks like (2 mks)
                mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
                marks = int(mk.group(1)) if mk else 1
                raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

                # create current question and attach pending_instructions to it
                current = {
                    "question": raw_question,
                    "instructions": pending_instructions.strip(),
                    "a": "",
                    "b": "",
                    "c": "",
                    "d": "",
                    "answer": "",   # keep the working scoring schema
                    "marks": marks,
                    "image": None
                }

                # extract image if requested (FIXED: now gets correct image per question)
                if image_output_dir:
                    imgs = extract_images(doc, image_output_dir, q_index)
                    if imgs:
                        current["image"] = imgs[0]

                # reset pending_instructions; if there was embedded_case put it back as pending for next question
                pending_instructions = embedded_case

            # ---------- OPTION (A-D) ----------
            elif current and is_option_line(line):
                letter = line[0].lower()
                text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
                current[letter] = text

            # ---------- ANSWER LINE (may appear anywhere after options) ----------
            elif current and is_answer_line(line):
                # handle patterns like "Answer: B Use the following..."
                # split into answer part and trailing case study if present
                parts = re.split(r"(use the following.*|study the information.*|refer to the following.*|case study.*|use the data below.*)",
                                 line, flags=re.IGNORECASE)
                answer_part = parts[0]
                trailing_case = ""
                if len(parts) > 1:
                    trailing_case = "".join(parts[1:]).strip()

                raw = answer_part.split(":")[-1].strip().lower()
                clean = re.sub(r"[^a-d]", "", raw)
                current["answer"] = clean

                # if trailing_case exists, attach it to pending_instructions for next question (rule B)
                if trailing_case:
                    # normalize
                    pending_instructions += ("<br>" if pending_instructions else "") + trailing_case

            # ---------- CASE STUDY LINE BEFORE ANY QUESTION (attach to pending_instructions) ----------
            elif is_case_study_line(line) and not current:
                pending_instructions += ("<br>" if pending_instructions else "") + line

            # ---------- CASE STUDY LINE AFTER A QUESTION (should be attached to next question per rule B) ----------
            elif is_case_study_line(line) and current:
                # attach to pending so next question will receive it
                pending_instructions += ("<br>" if pending_instructions else "") + line

            # ---------- OTHER TEXT: attach to current question text (question continuation) or to pending if no current ----------
            else:
                if current:
                    current["question"] += " " + line
                else:
                    pending_instructions += ("<br>" if pending_instructions else "") + line

        # ---------- TABLE ENTRY ----------
        elif entry["type"] == "table":
            html_table = make_html_table(entry["cells"])
            # If no current question exists, treat table as part of pending instructions (case study)
            if not current:
                pending_instructions += ("<br>" if pending_instructions else "") + html_table
            else:
                # attach table to current question text
                current["question"] += "<br>" + html_table

    # push last question
    if current:
        questions.append(current)

    return questions

# --------------- PREPARE STUDENT-FACING (HIDE ANSWERS) ---------------
def prepare_questions_for_student(questions, include_instructions_once=True):
    """
    Returns a list of questions safe to display to students (without answers).
    If include_instructions_once=True, returns instructions separately as 'page_instructions' and questions list.
    Otherwise each question includes its 'instructions' field.
    """
    if include_instructions_once:
        # gather first non-empty instructions (or concatenate distinct ones)
        page_instructions = ""
        for q in questions:
            if q.get("instructions"):
                if page_instructions:
                    page_instructions += "<hr>" + q["instructions"]
                else:
                    page_instructions = q["instructions"]
        # build student questions without answers
        student_questions = []
        for q in questions:
            student_questions.append({
                "question": q["question"],
                "a": q["a"],
                "b": q["b"],
                "c": q["c"],
                "d": q["d"],
                "marks": q["marks"],
                "image": q["image"]
            })
        return {"page_instructions": page_instructions, "questions": student_questions}
    else:
        student_questions = []
        for q in questions:
            student_questions.append({
                "question": q["question"],
                "instructions": q["instructions"],
                "a": q["a"],
                "b": q["b"],
                "c": q["c"],
                "d": q["d"],
                "marks": q["marks"],
                "image": q["image"]
            })
        return {"page_instructions": "", "questions": student_questions}

# ------------------ SCORING ENGINE (keeps your working format) ------------------
def compute_score(questions, student_answers):
    """
    questions: list produced by parse_docx_questions (uses question['answer'] as correct)
    student_answers: dict-like, expected keys 'q1','q2',... or '1','2',...
    returns dict with score, total, percentage, details list
    """
    score = 0
    total_marks = 0
    details = []

    for index, q in enumerate(questions, start=1):
        # normalize correct (from parser stored in 'answer')
        correct = (q.get("answer", "") or "").strip().lower()
        correct = re.sub(r"[^a-d]", "", correct)

        total_marks += q.get("marks", 1)

        # find student answer: prefer 'q{index}' then '{index}'
        student_answer = ""
        for key in (f"q{index}", str(index)):
            if key in student_answers:
                raw = (student_answers[key] or "").strip().lower()
                student_answer = re.sub(r"[^a-d]", "", raw)
                break

        # Safety: if both are empty, treat as unanswered (not correct)
        got_it = (student_answer != "" and student_answer == correct)

        if got_it:
            score += q.get("marks", 1)

        details.append({
            "question": q.get("question", ""),
            "correct": correct,
            "student_answer": student_answer,
            "marks": q.get("marks", 1),
            "earned": q.get("marks", 1) if got_it else 0
        })

    percentage = round((score / total_marks) * 100, 2) if total_marks else 0

    return {
        "score": score,
        "total": total_marks,
        "percentage": percentage,
        "details": details
    }

# ------------------ QUIZ STATUS ------------------
def get_quiz_status(questions, student_answers):
    status_list = []
    for index, q in enumerate(questions, start=1):
        correct = (q.get("answer", "") or "").strip().lower()
        correct = re.sub(r"[^a-d]", "", correct)

        student_answer = ""
        for key in (f"q{index}", str(index)):
            if key in student_answers:
                raw = (student_answers[key] or "").strip().lower()
                student_answer = re.sub(r"[^a-d]", "", raw)
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

# ------------------ GOOGLE DRIVE HELPERS ------------------
def extract_drive_id(url):
    patterns = [
        r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
        r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
    ]
    for pattern in patterns:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return url

def get_drive_embed_url(drive_url_or_id):
    file_id = extract_drive_id(drive_url_or_id)
    return f"https://drive.google.com/file/d/{file_id}/preview"