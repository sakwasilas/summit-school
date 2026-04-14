# # # # #import docx
# # # # #import re
# # # # # import os
# # # # # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # # # # ------------------- IMAGE EXTRACTION (FIXED - NO DUPLICATION) -----------------
# # # # # # Global cache to store extracted images per document
# # # # # _image_cache = {}

# # # # # def extract_images(document, output_dir, q_index):
# # # # #     """
# # # # #     Extract images from document - but only once per document.
# # # # #     Returns the image for the specific question index.
# # # # #     """
# # # # #     global _image_cache
    
# # # # #     # Create a cache key based on document and output directory
# # # # #     cache_key = f"{id(document)}_{output_dir}"
    
# # # # #     # If we haven't extracted images for this document yet, do it once
# # # # #     if cache_key not in _image_cache:
# # # # #         os.makedirs(output_dir, exist_ok=True)
# # # # #         images = []
# # # # #         count = 0
        
# # # # #         # Extract all images and store them in order
# # # # #         for rel in document.part.rels.values():
# # # # #             if rel.reltype == RT.IMAGE:
# # # # #                 count += 1
# # # # #                 ext = rel.target_ref.split('.')[-1]
# # # # #                 # Use generic names without question index
# # # # #                 filename = f"img_{count}.{ext}"
# # # # #                 filepath = os.path.join(output_dir, filename)
# # # # #                 with open(filepath, "wb") as f:
# # # # #                     f.write(rel.target_part.blob)
# # # # #                 images.append(filename)
        
# # # # #         _image_cache[cache_key] = images
    
# # # # #     # Return the image for this question index (1-based)
# # # # #     all_images = _image_cache[cache_key]
# # # # #     if q_index <= len(all_images):
# # # # #         return [all_images[q_index - 1]]  # Return as list to maintain compatibility
# # # # #     else:
# # # # #         return []  # No image for this question

# # # # # # --------------------- LOAD DOCX ---------------------
# # # # # def load_docx(path):
# # # # #     return docx.Document(path)

# # # # # # --------------------- HELPERS -----------------------
# # # # # def is_question_line(text):
# # # # #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # # # # def is_option_line(text):
# # # # #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # # # # def is_answer_line(text):
# # # # #     return text.lower().lstrip().startswith(("answer", "ans", "correct"))

# # # # # def is_case_study_line(text):
# # # # #     keywords = [
# # # # #         "use the following information",
# # # # #         "study the information",
# # # # #         "refer to the following",
# # # # #         "case study",
# # # # #         "use the data below"
# # # # #     ]
# # # # #     t = text.lower()
# # # # #     return any(k in t for k in keywords)

# # # # # # ------------------- TABLE -> HTML -------------------
# # # # # def make_html_table(cells):
# # # # #     html = "<table class='table table-bordered'>"
# # # # #     for row in cells:
# # # # #         html += "<tr>"
# # # # #         for c in row:
# # # # #             # Escape minimal HTML-sensitive chars (basic)
# # # # #             safe = (c.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
# # # # #             html += f"<td>{safe}</td>"
# # # # #         html += "</tr>"
# # # # #     html += "</table>"
# # # # #     return html

# # # # # # ------------------- FLATTEN DOCX --------------------
# # # # # def flatten_doc(document):
# # # # #     """
# # # # #     Returns a list of entries preserving paragraphs and tables in document order.
# # # # #     Each entry is {"type":"text","content":...} or {"type":"table","cells": [...]}
# # # # #     """
# # # # #     lines = []
# # # # #     for block in document.element.body:
# # # # #         if block.tag.endswith('p'):
# # # # #             para = docx.text.paragraph.Paragraph(block, document)
# # # # #             text = para.text.strip()
# # # # #             if text:
# # # # #                 lines.append({"type": "text", "content": text})
# # # # #         elif block.tag.endswith('tbl'):
# # # # #             table = docx.table.Table(block, document)
# # # # #             rows = []
# # # # #             for row in table.rows:
# # # # #                 cells = [cell.text.strip() for cell in row.cells]
# # # # #                 rows.append(cells)
# # # # #             lines.append({"type": "table", "cells": rows})
# # # # #     return lines

# # # # # # --------------- PARSE DOCX QUESTIONS -----------------
# # # # # def parse_docx_questions(path, image_output_dir=None):
# # # # #     """
# # # # #     Parse docx file and return list of questions.
# # # # #     Each question dict has:
# # # # #       - question (text with any inline tables as HTML)
# # # # #       - instructions (case study/instructions that were appearing BEFORE the question)
# # # # #       - a, b, c, d (option texts)
# # # # #       - answer (single letter 'a'..'d', stored lower-case)
# # # # #       - marks (int)
# # # # #       - image (filename or None)
# # # # #     Case studies/tables that appear before a question are attached to the next question's 'instructions'.
# # # # #     """
# # # # #     doc = load_docx(path)
# # # # #     entries = flatten_doc(doc)

# # # # #     questions = []
# # # # #     current = None
# # # # #     q_index = 0
# # # # #     # holds instructions/case-study found before the *next* question
# # # # #     pending_instructions = ""

# # # # #     for entry in entries:
# # # # #         if entry["type"] == "text":
# # # # #             line = entry["content"].strip()

# # # # #             # If line contains both an Answer and a new case-study after it (same paragraph),
# # # # #             # we'll handle splitting later when looking for answer lines.
# # # # #             # ---------- NEW QUESTION ----------
# # # # #             if is_question_line(line):
# # # # #                 # push previous
# # # # #                 if current:
# # # # #                     questions.append(current)

# # # # #                 q_index += 1
# # # # #                 raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# # # # #                 # detect inline embedded case study text in same line as question (rare)
# # # # #                 embedded_case = ""
# # # # #                 for k in [
# # # # #                     "use the following information",
# # # # #                     "study the information",
# # # # #                     "refer to the following",
# # # # #                     "case study",
# # # # #                     "use the data below"
# # # # #                 ]:
# # # # #                     if k in raw_question.lower():
# # # # #                         parts = re.split(k, raw_question, flags=re.IGNORECASE)
# # # # #                         raw_question = parts[0].strip()
# # # # #                         embedded_case = k + " " + parts[1].strip()
# # # # #                         break

# # # # #                 # extract marks like (2 mks)
# # # # #                 mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
# # # # #                 marks = int(mk.group(1)) if mk else 1
# # # # #                 raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

# # # # #                 # create current question and attach pending_instructions to it
# # # # #                 current = {
# # # # #                     "question": raw_question,
# # # # #                     "instructions": pending_instructions.strip(),
# # # # #                     "a": "",
# # # # #                     "b": "",
# # # # #                     "c": "",
# # # # #                     "d": "",
# # # # #                     "answer": "",   # keep the working scoring schema
# # # # #                     "marks": marks,
# # # # #                     "image": None
# # # # #                 }

# # # # #                 # extract image if requested (FIXED: now gets correct image per question)
# # # # #                 if image_output_dir:
# # # # #                     imgs = extract_images(doc, image_output_dir, q_index)
# # # # #                     if imgs:
# # # # #                         current["image"] = imgs[0]

# # # # #                 # reset pending_instructions; if there was embedded_case put it back as pending for next question
# # # # #                 pending_instructions = embedded_case

# # # # #             # ---------- OPTION (A-D) ----------
# # # # #             elif current and is_option_line(line):
# # # # #                 letter = line[0].lower()
# # # # #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# # # # #                 current[letter] = text

# # # # #             # ---------- ANSWER LINE (may appear anywhere after options) ----------
# # # # #             elif current and is_answer_line(line):
# # # # #                 # handle patterns like "Answer: B Use the following..."
# # # # #                 # split into answer part and trailing case study if present
# # # # #                 parts = re.split(r"(use the following.*|study the information.*|refer to the following.*|case study.*|use the data below.*)",
# # # # #                                  line, flags=re.IGNORECASE)
# # # # #                 answer_part = parts[0]
# # # # #                 trailing_case = ""
# # # # #                 if len(parts) > 1:
# # # # #                     trailing_case = "".join(parts[1:]).strip()

# # # # #                 raw = answer_part.split(":")[-1].strip().lower()
# # # # #                 clean = re.sub(r"[^a-d]", "", raw)
# # # # #                 current["answer"] = clean

# # # # #                 # if trailing_case exists, attach it to pending_instructions for next question (rule B)
# # # # #                 if trailing_case:
# # # # #                     # normalize
# # # # #                     pending_instructions += ("<br>" if pending_instructions else "") + trailing_case

# # # # #             # ---------- CASE STUDY LINE BEFORE ANY QUESTION (attach to pending_instructions) ----------
# # # # #             elif is_case_study_line(line) and not current:
# # # # #                 pending_instructions += ("<br>" if pending_instructions else "") + line

# # # # #             # ---------- CASE STUDY LINE AFTER A QUESTION (should be attached to next question per rule B) ----------
# # # # #             elif is_case_study_line(line) and current:
# # # # #                 # attach to pending so next question will receive it
# # # # #                 pending_instructions += ("<br>" if pending_instructions else "") + line

# # # # #             # ---------- OTHER TEXT: attach to current question text (question continuation) or to pending if no current ----------
# # # # #             else:
# # # # #                 if current:
# # # # #                     current["question"] += " " + line
# # # # #                 else:
# # # # #                     pending_instructions += ("<br>" if pending_instructions else "") + line

# # # # #         # ---------- TABLE ENTRY ----------
# # # # #         elif entry["type"] == "table":
# # # # #             html_table = make_html_table(entry["cells"])
# # # # #             # If no current question exists, treat table as part of pending instructions (case study)
# # # # #             if not current:
# # # # #                 pending_instructions += ("<br>" if pending_instructions else "") + html_table
# # # # #             else:
# # # # #                 # attach table to current question text
# # # # #                 current["question"] += "<br>" + html_table

# # # # #     # push last question
# # # # #     if current:
# # # # #         questions.append(current)

# # # # #     return questions

# # # # # # --------------- PREPARE STUDENT-FACING (HIDE ANSWERS) ---------------
# # # # # def prepare_questions_for_student(questions, include_instructions_once=True):
# # # # #     """
# # # # #     Returns a list of questions safe to display to students (without answers).
# # # # #     If include_instructions_once=True, returns instructions separately as 'page_instructions' and questions list.
# # # # #     Otherwise each question includes its 'instructions' field.
# # # # #     """
# # # # #     if include_instructions_once:
# # # # #         # gather first non-empty instructions (or concatenate distinct ones)
# # # # #         page_instructions = ""
# # # # #         for q in questions:
# # # # #             if q.get("instructions"):
# # # # #                 if page_instructions:
# # # # #                     page_instructions += "<hr>" + q["instructions"]
# # # # #                 else:
# # # # #                     page_instructions = q["instructions"]
# # # # #         # build student questions without answers
# # # # #         student_questions = []
# # # # #         for q in questions:
# # # # #             student_questions.append({
# # # # #                 "question": q["question"],
# # # # #                 "a": q["a"],
# # # # #                 "b": q["b"],
# # # # #                 "c": q["c"],
# # # # #                 "d": q["d"],
# # # # #                 "marks": q["marks"],
# # # # #                 "image": q["image"]
# # # # #             })
# # # # #         return {"page_instructions": page_instructions, "questions": student_questions}
# # # # #     else:
# # # # #         student_questions = []
# # # # #         for q in questions:
# # # # #             student_questions.append({
# # # # #                 "question": q["question"],
# # # # #                 "instructions": q["instructions"],
# # # # #                 "a": q["a"],
# # # # #                 "b": q["b"],
# # # # #                 "c": q["c"],
# # # # #                 "d": q["d"],
# # # # #                 "marks": q["marks"],
# # # # #                 "image": q["image"]
# # # # #             })
# # # # #         return {"page_instructions": "", "questions": student_questions}

# # # # # # ------------------ SCORING ENGINE (keeps your working format) ------------------
# # # # # def compute_score(questions, student_answers):
# # # # #     """
# # # # #     questions: list produced by parse_docx_questions (uses question['answer'] as correct)
# # # # #     student_answers: dict-like, expected keys 'q1','q2',... or '1','2',...
# # # # #     returns dict with score, total, percentage, details list
# # # # #     """
# # # # #     score = 0
# # # # #     total_marks = 0
# # # # #     details = []

# # # # #     for index, q in enumerate(questions, start=1):
# # # # #         # normalize correct (from parser stored in 'answer')
# # # # #         correct = (q.get("answer", "") or "").strip().lower()
# # # # #         correct = re.sub(r"[^a-d]", "", correct)

# # # # #         total_marks += q.get("marks", 1)

# # # # #         # find student answer: prefer 'q{index}' then '{index}'
# # # # #         student_answer = ""
# # # # #         for key in (f"q{index}", str(index)):
# # # # #             if key in student_answers:
# # # # #                 raw = (student_answers[key] or "").strip().lower()
# # # # #                 student_answer = re.sub(r"[^a-d]", "", raw)
# # # # #                 break

# # # # #         # Safety: if both are empty, treat as unanswered (not correct)
# # # # #         got_it = (student_answer != "" and student_answer == correct)

# # # # #         if got_it:
# # # # #             score += q.get("marks", 1)

# # # # #         details.append({
# # # # #             "question": q.get("question", ""),
# # # # #             "correct": correct,
# # # # #             "student_answer": student_answer,
# # # # #             "marks": q.get("marks", 1),
# # # # #             "earned": q.get("marks", 1) if got_it else 0
# # # # #         })

# # # # #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# # # # #     return {
# # # # #         "score": score,
# # # # #         "total": total_marks,
# # # # #         "percentage": percentage,
# # # # #         "details": details
# # # # #     }

# # # # # # ------------------ QUIZ STATUS ------------------
# # # # # def get_quiz_status(questions, student_answers):
# # # # #     status_list = []
# # # # #     for index, q in enumerate(questions, start=1):
# # # # #         correct = (q.get("answer", "") or "").strip().lower()
# # # # #         correct = re.sub(r"[^a-d]", "", correct)

# # # # #         student_answer = ""
# # # # #         for key in (f"q{index}", str(index)):
# # # # #             if key in student_answers:
# # # # #                 raw = (student_answers[key] or "").strip().lower()
# # # # #                 student_answer = re.sub(r"[^a-d]", "", raw)
# # # # #                 break

# # # # #         if not student_answer:
# # # # #             status = "unanswered"
# # # # #         elif student_answer == correct:
# # # # #             status = "correct"
# # # # #         else:
# # # # #             status = "incorrect"

# # # # #         status_list.append({
# # # # #             "question_index": index,
# # # # #             "status": status,
# # # # #             "student_answer": student_answer,
# # # # #             "correct_answer": correct
# # # # #         })
# # # # #     return status_list

# # # # # # ------------------ GOOGLE DRIVE HELPERS ------------------
# # # # # def extract_drive_id(url):
# # # # #     patterns = [
# # # # #         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
# # # # #         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
# # # # #     ]
# # # # #     for pattern in patterns:
# # # # #         m = re.search(pattern, url)
# # # # #         if m:
# # # # #             return m.group(1)
# # # # #     return url

# # # # # def get_drive_embed_url(drive_url_or_id):
# # # # #     file_id = extract_drive_id(drive_url_or_id)
# # # # #     return f"https://drive.google.com/file/d/{file_id}/preview"

# # # # import os
# # # # import re
# # # # from docx import Document
# # # # from docx.oxml.ns import qn
# # # # from docx.oxml.text.paragraph import CT_P
# # # # from docx.oxml.table import CT_Tbl
# # # # from docx.table import Table
# # # # from docx.text.paragraph import Paragraph

# # # # DEFAULT_IMAGE_DIR = "static/question_images"

# # # # # ------------------ TABLE TO HTML ------------------
# # # # def extract_table_html(table):
# # # #     html = "<table border='1' cellspacing='0' cellpadding='5'>"
# # # #     for row in table.rows:
# # # #         html += "<tr>"
# # # #         for cell in row.cells:
# # # #             html += f"<td>{cell.text.strip()}</td>"
# # # #         html += "</tr>"
# # # #     html += "</table>"
# # # #     return html


# # # # # ------------------ IMAGE EXTRACTION ------------------
# # # # def save_image_from_run(run, output_dir, image_counter):
# # # #     blip_elements = run._element.findall('.//a:blip', namespaces={
# # # #         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
# # # #     })

# # # #     if not blip_elements:
# # # #         return None

# # # #     rId = blip_elements[0].get(qn('r:embed'))
# # # #     image_part = run.part.related_parts[rId]
# # # #     image_data = image_part.blob

# # # #     image_filename = f"question_image_{image_counter}.png"
# # # #     image_path = os.path.join(output_dir, image_filename)

# # # #     with open(image_path, 'wb') as f:
# # # #         f.write(image_data)

# # # #     return image_filename


# # # # # ------------------ ITERATE DOCX BLOCKS ------------------
# # # # def iter_block_items(parent):
# # # #     for child in parent.element.body.iterchildren():
# # # #         if isinstance(child, CT_P):
# # # #             yield Paragraph(child, parent)
# # # #         elif isinstance(child, CT_Tbl):
# # # #             yield Table(child, parent)


# # # # # ------------------ MAIN PARSER ------------------
# # # # def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
# # # #     document = Document(file_stream)
# # # #     questions = []
# # # #     current_question = None
# # # #     extra_html_parts = []
# # # #     image_counter = 0
# # # #     skipped = 0

# # # #     os.makedirs(image_output_dir, exist_ok=True)

# # # #     for block in iter_block_items(document):

# # # #         # ---------------- PARAGRAPHS ----------------
# # # #         if isinstance(block, Paragraph):
# # # #             text = block.text.strip()

# # # #             # Attach image
# # # #             for run in block.runs:
# # # #                 image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
# # # #                 if image_name and current_question:
# # # #                     image_counter += 1
# # # #                     current_question["image"] = image_name

# # # #             if not text:
# # # #                 continue

# # # #             # ✅ FIX: Clean double numbering like "37. 29."
# # # #             text = re.sub(r"^\d+[\.\)]\s*\d+[\.\)]\s*", lambda m: m.group(0).split()[-1] + " ", text)

# # # #             # ---------------- NEW QUESTION ----------------
# # # #             if re.match(r"^\d+[\.\)]", text):

# # # #                 if current_question:
# # # #                     current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None

# # # #                     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# # # #                         questions.append(current_question)
# # # #                     else:
# # # #                         print("❌ Skipped question:", current_question.get("question"))
# # # #                         skipped += 1

# # # #                     extra_html_parts = []

# # # #                 # Extract marks
# # # #                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
# # # #                 marks = int(marks_match.group(1)) if marks_match else 1

# # # #                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
# # # #                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)

# # # #                 current_question = {
# # # #                     "question": question_text,
# # # #                     "a": "", "b": "", "c": "", "d": "",
# # # #                     "answer": "",
# # # #                     "extra_content": None,
# # # #                     "image": None,
# # # #                     "marks": marks
# # # #                 }

# # # #             # ---------------- OPTIONS ----------------
# # # #             elif re.match(r"^\(?[a-dA-D][\.\)]\s*", text):
# # # #                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
# # # #                 if match and current_question:
# # # #                     label = match.group(1).lower()
# # # #                     content = match.group(2).strip()
# # # #                     current_question[label] = content

# # # #             # ---------------- ANSWER ----------------
# # # #             elif re.match(r"^(answer|ans|correct answer)\s*:?", text, re.IGNORECASE):
# # # #                 match = re.search(r"([a-dA-D])", text)
# # # #                 if match:
# # # #                     if current_question:
# # # #                         current_question["answer"] = match.group(1).lower()
# # # #                     else:
# # # #                         print("⚠️ Answer found without a question.")

# # # #             # ---------------- EXTRA CONTENT ----------------
# # # #             else:
# # # #                 extra_html_parts.append(f"<p>{text}</p>")

# # # #         # ---------------- TABLE ----------------
# # # #         elif isinstance(block, Table):
# # # #             table_html = extract_table_html(block)

# # # #             if current_question:
# # # #                 current_question["extra_content"] = (current_question.get("extra_content") or '') + table_html
# # # #             else:
# # # #                 extra_html_parts.append(table_html)

# # # #     # ---------------- FINAL QUESTION ----------------
# # # #     if current_question:
# # # #         current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None

# # # #         if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# # # #             questions.append(current_question)
# # # #         else:
# # # #             print("❌ Skipped last question:", current_question.get("question"))
# # # #             skipped += 1

# # # #     print(f"✅ Parsed {len(questions)} valid questions.")
# # # #     if skipped:
# # # #         print(f"⚠️ Skipped {skipped} invalid question(s).")

# # # #     return questions


# # # # # ------------------ QUIZ STATUS ------------------
# # # # def get_quiz_status(user_id):
# # # #     return "active"


# # # # # ------------------ GOOGLE DRIVE HELPERS ------------------
# # # # def extract_drive_id(url):
# # # #     patterns = [
# # # #         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
# # # #         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
# # # #     ]
# # # #     for pattern in patterns:
# # # #         match = re.search(pattern, url)
# # # #         if match:
# # # #             return match.group(1)
# # # #     return url


# # # # def get_drive_embed_url(drive_url_or_id):
# # # #     file_id = extract_drive_id(drive_url_or_id)
# # # #     return f"https://drive.google.com/file/d/{file_id}/preview"

# # # # import os
# # # # import re
# # # # from docx import Document
# # # # from docx.oxml.ns import qn
# # # # from docx.oxml.text.paragraph import CT_P
# # # # from docx.oxml.table import CT_Tbl
# # # # from docx.table import Table
# # # # from docx.text.paragraph import Paragraph

# # # # DEFAULT_IMAGE_DIR = "static/question_images"

# # # # # ------------------ TABLE TO HTML ------------------
# # # # def extract_table_html(table):
# # # #     html = "<table border='1' cellspacing='0' cellpadding='5'>"
# # # #     for row in table.rows:
# # # #         html += "<tr>"
# # # #         for cell in row.cells:
# # # #             html += f"<td>{cell.text.strip()}</td>"
# # # #         html += "</tr>"
# # # #     html += "</table>"
# # # #     return html

# # # # # ------------------ IMAGE EXTRACTION ------------------
# # # # def save_image_from_run(run, output_dir, question_number, image_index):
# # # #     blip_elements = run._element.findall('.//a:blip', namespaces={
# # # #         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
# # # #     })
# # # #     if not blip_elements:
# # # #         return None

# # # #     rId = blip_elements[0].get(qn('r:embed'))
# # # #     image_part = run.part.related_parts[rId]
# # # #     image_data = image_part.blob

# # # #     image_filename = f"question_{question_number}_image_{image_index}.png"
# # # #     image_path = os.path.join(output_dir, image_filename)
# # # #     with open(image_path, 'wb') as f:
# # # #         f.write(image_data)

# # # #     return image_filename

# # # # # ------------------ ITERATE DOCX BLOCKS ------------------
# # # # def iter_block_items(parent):
# # # #     for child in parent.element.body.iterchildren():
# # # #         if isinstance(child, CT_P):
# # # #             yield Paragraph(child, parent)
# # # #         elif isinstance(child, CT_Tbl):
# # # #             yield Table(child, parent)

# # # # # ------------------ MAIN PARSER ------------------
# # # # def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
# # # #     document = Document(file_stream)
# # # #     questions = []
# # # #     current_question = None
# # # #     extra_html_parts = []
# # # #     question_counter = 0

# # # #     os.makedirs(image_output_dir, exist_ok=True)

# # # #     for block in iter_block_items(document):
# # # #         if isinstance(block, Paragraph):
# # # #             text = block.text.strip()
# # # #             if not text:
# # # #                 continue

# # # #             # ---------------- NEW QUESTION ----------------
# # # #             if re.match(r"^\d+[\.\)]", text):
# # # #                 # Save previous question
# # # #                 if current_question:
# # # #                     current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
# # # #                     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# # # #                         questions.append(current_question)
# # # #                     extra_html_parts = []

# # # #                 question_counter += 1

# # # #                 # Extract marks
# # # #                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
# # # #                 marks = int(marks_match.group(1)) if marks_match else 1

# # # #                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
# # # #                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)

# # # #                 current_question = {
# # # #                     "question_number": question_counter,
# # # #                     "question": question_text,
# # # #                     "a": "", "b": "", "c": "", "d": "",
# # # #                     "answer": "",
# # # #                     "extra_content": None,
# # # #                     "images": [],  # store multiple images per question
# # # #                     "marks": marks
# # # #                 }

# # # #             # ---------------- OPTIONS ----------------
# # # #             elif re.match(r"^\(?[a-dA-D][\.\)]\s*", text) and current_question:
# # # #                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
# # # #                 if match:
# # # #                     current_question[match.group(1).lower()] = match.group(2).strip()

# # # #             # ---------------- ANSWER ----------------
# # # #             elif re.match(r"^(answer|ans|correct answer)\s*:?", text, re.IGNORECASE) and current_question:
# # # #                 match = re.search(r"([a-dA-D])", text)
# # # #                 if match:
# # # #                     current_question["answer"] = match.group(1).lower().strip()

# # # #             # ---------------- EXTRA CONTENT ----------------
# # # #             else:
# # # #                 extra_html_parts.append(f"<p>{text}</p>")

# # # #             # ---------------- IMAGE ATTACHMENT ----------------
# # # #             if current_question:
# # # #                 image_index = 0
# # # #                 for run in block.runs:
# # # #                     image_name = save_image_from_run(run, image_output_dir, current_question["question_number"], image_index + 1)
# # # #                     if image_name:
# # # #                         image_index += 1
# # # #                         current_question["images"].append(image_name)

# # # #         # ---------------- TABLE ----------------
# # # #         elif isinstance(block, Table):
# # # #             table_html = extract_table_html(block)
# # # #             if current_question:
# # # #                 current_question["extra_content"] = (current_question.get("extra_content") or '') + table_html
# # # #             else:
# # # #                 extra_html_parts.append(table_html)

# # # #     # ---------------- FINAL QUESTION ----------------
# # # #     if current_question:
# # # #         current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
# # # #         if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# # # #             questions.append(current_question)

# # # #     print(f"✅ Parsed {len(questions)} questions successfully.")
# # # #     return questions
# # # import os
# # # import re
# # # from docx import Document
# # # from docx.oxml.ns import qn
# # # from docx.oxml.text.paragraph import CT_P
# # # from docx.oxml.table import CT_Tbl
# # # from docx.table import Table
# # # from docx.text.paragraph import Paragraph

# # # DEFAULT_IMAGE_DIR = "static/question_images"

# # # # ------------------ TABLE TO HTML ------------------
# # # def extract_table_html(table):
# # #     html = "<table border='1' cellspacing='0' cellpadding='5'>"
# # #     for row in table.rows:
# # #         html += "<tr>"
# # #         for cell in row.cells:
# # #             html += f"<td>{cell.text.strip()}</td>"
# # #         html += "</tr>"
# # #     html += "</table>"
# # #     return html

# # # # ------------------ IMAGE EXTRACTION ------------------
# # # def save_image_from_run(run, output_dir, question_number, image_index):
# # #     blip_elements = run._element.findall('.//a:blip', namespaces={
# # #         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
# # #     })
# # #     if not blip_elements:
# # #         return None

# # #     rId = blip_elements[0].get(qn('r:embed'))
# # #     image_part = run.part.related_parts[rId]
# # #     image_data = image_part.blob

# # #     image_filename = f"question_{question_number}_image_{image_index}.png"
# # #     image_path = os.path.join(output_dir, image_filename)
# # #     with open(image_path, 'wb') as f:
# # #         f.write(image_data)

# # #     return image_filename

# # # # ------------------ ITERATE DOCX BLOCKS ------------------
# # # def iter_block_items(parent):
# # #     for child in parent.element.body.iterchildren():
# # #         if isinstance(child, CT_P):
# # #             yield Paragraph(child, parent)
# # #         elif isinstance(child, CT_Tbl):
# # #             yield Table(child, parent)

# # # # ------------------ MAIN PARSER ------------------
# # # def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
# # #     document = Document(file_stream)
# # #     questions = []
# # #     current_question = None
# # #     extra_html_parts = []
# # #     question_counter = 0

# # #     os.makedirs(image_output_dir, exist_ok=True)

# # #     for block in iter_block_items(document):
# # #         if isinstance(block, Paragraph):
# # #             text = block.text.strip()
# # #             if not text:
# # #                 continue

# # #             # ---------------- NEW QUESTION ----------------
# # #             if re.match(r"^\d+[\.\)]", text):
# # #                 # Save previous question
# # #                 if current_question:
# # #                     current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
# # #                     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# # #                         questions.append(current_question)
# # #                     extra_html_parts = []

# # #                 question_counter += 1

# # #                 # Extract marks
# # #                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
# # #                 marks = int(marks_match.group(1)) if marks_match else 1

# # #                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
# # #                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)

# # #                 current_question = {
# # #                     "question_number": question_counter,
# # #                     "question": question_text,
# # #                     "a": "", "b": "", "c": "", "d": "",
# # #                     "answer": "",
# # #                     "extra_content": None,
# # #                     "images": [],
# # #                     "marks": marks
# # #                 }

# # #             # ---------------- OPTIONS ----------------
# # #             elif re.match(r"^\(?[a-dA-D][\.\)]\s*", text) and current_question:
# # #                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
# # #                 if match:
# # #                     current_question[match.group(1).lower()] = match.group(2).strip()

# # #             # ---------------- ANSWER ----------------
# # #             elif re.match(r"^(answer|ans|correct answer)\s*:?", text, re.IGNORECASE) and current_question:
# # #                 match = re.search(r"([a-dA-D])", text)
# # #                 if match:
# # #                     current_question["answer"] = match.group(1).lower().strip()

# # #             # ---------------- EXTRA CONTENT ----------------
# # #             else:
# # #                 extra_html_parts.append(f"<p>{text}</p>")

# # #             # ---------------- IMAGE ATTACHMENT ----------------
# # #             if current_question:
# # #                 image_index = 0
# # #                 for run in block.runs:
# # #                     image_name = save_image_from_run(run, image_output_dir, current_question["question_number"], image_index + 1)
# # #                     if image_name:
# # #                         image_index += 1
# # #                         current_question["images"].append(image_name)

# # #         # ---------------- TABLE ----------------
# # #         elif isinstance(block, Table):
# # #             table_html = extract_table_html(block)
# # #             if current_question:
# # #                 current_question["extra_content"] = (current_question.get("extra_content") or '') + table_html
# # #             else:
# # #                 extra_html_parts.append(table_html)

# # #     # ---------------- FINAL QUESTION ----------------
# # #     if current_question:
# # #         current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
# # #         if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# # #             questions.append(current_question)

# # #     print(f"✅ Parsed {len(questions)} questions successfully.")
# # #     return questions

# # # # ------------------ QUIZ STATUS ------------------
# # # def get_quiz_status(user_id):
# # #     """Return quiz status for a user."""
# # #     return "active"

# # # import os
# # # import re
# # # from docx import Document
# # # from docx.oxml.ns import qn
# # # from docx.oxml.text.paragraph import CT_P
# # # from docx.oxml.table import CT_Tbl
# # # from docx.table import Table
# # # from docx.text.paragraph import Paragraph

# # # DEFAULT_IMAGE_DIR = "static/question_images"

# # # # ------------------ TABLE TO HTML ------------------
# # # def extract_table_html(table):
# # #     html = "<table border='1' cellspacing='0' cellpadding='5'>"
# # #     for row in table.rows:
# # #         html += "<tr>"
# # #         for cell in row.cells:
# # #             html += f"<td>{cell.text.strip()}</td>"
# # #         html += "</tr>"
# # #     html += "</table>"
# # #     return html

# # # # ------------------ IMAGE EXTRACTION ------------------
# # # def save_image_from_run(run, output_dir, question_number, image_index):
# # #     blip_elements = run._element.findall('.//a:blip', namespaces={
# # #         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
# # #     })
# # #     if not blip_elements:
# # #         return None

# # #     rId = blip_elements[0].get(qn('r:embed'))
# # #     image_part = run.part.related_parts[rId]
# # #     image_data = image_part.blob

# # #     image_filename = f"question_{question_number}_image_{image_index}.png"
# # #     image_path = os.path.join(output_dir, image_filename)
# # #     os.makedirs(output_dir, exist_ok=True)
# # #     with open(image_path, 'wb') as f:
# # #         f.write(image_data)

# # #     return image_filename

# # # # ------------------ ITERATE DOCX BLOCKS ------------------
# # # def iter_block_items(parent):
# # #     for child in parent.element.body.iterchildren():
# # #         if isinstance(child, CT_P):
# # #             yield Paragraph(child, parent)
# # #         elif isinstance(child, CT_Tbl):
# # #             yield Table(child, parent)

# # # # ------------------ MAIN PARSER ------------------
# # # def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
# # #     document = Document(file_stream)
# # #     questions = []
# # #     current_question = None
# # #     extra_html_parts = []
# # #     question_counter = 0

# # #     os.makedirs(image_output_dir, exist_ok=True)

# # #     for block in iter_block_items(document):
# # #         if isinstance(block, Paragraph):
# # #             text = block.text.strip()
# # #             if not text:
# # #                 continue

# # #             # ---------------- NEW QUESTION ----------------
# # #             if re.match(r"^\d+[\.\)]", text):
# # #                 if current_question:
# # #                     current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
# # #                     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# # #                         questions.append(current_question)
# # #                     extra_html_parts = []

# # #                 question_counter += 1

# # #                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
# # #                 marks = int(marks_match.group(1)) if marks_match else 1

# # #                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
# # #                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)

# # #                 current_question = {
# # #                     "question_number": question_counter,
# # #                     "question": question_text,
# # #                     "a": "", "b": "", "c": "", "d": "",
# # #                     "answer": "",
# # #                     "extra_content": None,
# # #                     "images": [],
# # #                     "marks": marks
# # #                 }

# # #             # ---------------- OPTIONS ----------------
# # #             elif re.match(r"^\(?[a-dA-D][\.\)]\s*", text) and current_question:
# # #                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
# # #                 if match:
# # #                     current_question[match.group(1).lower()] = match.group(2).strip()

# # #             # ---------------- ANSWER ----------------
# # #             elif re.match(r"^(answer|ans|correct answer)\s*:?", text, re.IGNORECASE) and current_question:
# # #                 match = re.search(r"([a-dA-D])", text)
# # #                 if match:
# # #                     current_question["answer"] = match.group(1).lower().strip()

# # #             # ---------------- EXTRA CONTENT ----------------
# # #             else:
# # #                 extra_html_parts.append(f"<p>{text}</p>")

# # #             # ---------------- IMAGE ATTACHMENT ----------------
# # #             if current_question:
# # #                 image_index = 0
# # #                 for run in block.runs:
# # #                     image_name = save_image_from_run(run, image_output_dir, current_question["question_number"], image_index + 1)
# # #                     if image_name:
# # #                         image_index += 1
# # #                         current_question["images"].append(image_name)

# # #         # ---------------- TABLE ----------------
# # #         elif isinstance(block, Table):
# # #             table_html = extract_table_html(block)
# # #             if current_question:
# # #                 current_question["extra_content"] = (current_question.get("extra_content") or '') + table_html
# # #             else:
# # #                 extra_html_parts.append(table_html)

# # #     # ---------------- FINAL QUESTION ----------------
# # #     if current_question:
# # #         current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
# # #         if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# # #             questions.append(current_question)

# # #     print(f"✅ Parsed {len(questions)} questions successfully.")
# # #     return questions

# # # # ------------------ QUIZ STATUS ------------------
# # # def get_quiz_status(user_id):
# # #     return "active"

# # # # ------------------ GOOGLE DRIVE HELPERS ------------------
# # # def extract_drive_id(url):
# # #     patterns = [
# # #         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
# # #         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
# # #     ]
# # #     for pattern in patterns:
# # #         match = re.search(pattern, url)
# # #         if match:
# # #             return match.group(1)
# # #     return url

# # # def get_drive_embed_url(drive_url_or_id):
# # #     file_id = extract_drive_id(drive_url_or_id)
# # #     return f"https://drive.google.com/file/d/{file_id}/preview"

# # import os
# # import re
# # from docx import Document
# # from docx.oxml.ns import qn
# # from docx.oxml.text.paragraph import CT_P
# # from docx.oxml.table import CT_Tbl
# # from docx.table import Table
# # from docx.text.paragraph import Paragraph

# # DEFAULT_IMAGE_DIR = "static/question_images"

# # def extract_table_html(table):
# #     html = "<table border='1' cellspacing='0' cellpadding='5'>"
# #     for row in table.rows:
# #         html += "<tr>"
# #         for cell in row.cells:
# #             html += f"<td>{cell.text.strip()}</td>"
# #         html += "</tr>"
# #     html += "</table>"
# #     return html

# # def save_image_from_run(run, output_dir, question_number, image_index):
# #     blip_elements = run._element.findall('.//a:blip', namespaces={
# #         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
# #     })
# #     if not blip_elements:
# #         return None

# #     rId = blip_elements[0].get(qn('r:embed'))
# #     image_part = run.part.related_parts[rId]
# #     image_data = image_part.blob

# #     image_filename = f"question_{question_number}_image_{image_index}.png"
# #     image_path = os.path.join(output_dir, image_filename)
# #     os.makedirs(output_dir, exist_ok=True)
# #     with open(image_path, 'wb') as f:
# #         f.write(image_data)

# #     return image_filename

# # def iter_block_items(parent):
# #     for child in parent.element.body.iterchildren():
# #         if isinstance(child, CT_P):
# #             yield Paragraph(child, parent)
# #         elif isinstance(child, CT_Tbl):
# #             yield Table(child, parent)

# # def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
# #     document = Document(file_stream)
# #     questions = []
# #     current_question = None
# #     extra_html_parts = []
# #     question_counter = 0

# #     os.makedirs(image_output_dir, exist_ok=True)

# #     for block in iter_block_items(document):
# #         if isinstance(block, Paragraph):
# #             text = block.text.strip()
# #             if not text:
# #                 continue

# #             if re.match(r"^\d+[\.\)]", text):
# #                 if current_question:
# #                     current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
# #                     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# #                         questions.append(current_question)
# #                     extra_html_parts = []

# #                 question_counter += 1

# #                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
# #                 marks = int(marks_match.group(1)) if marks_match else 1

# #                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
# #                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text).strip()

# #                 current_question = {
# #                     "question_number": question_counter,
# #                     "question": question_text,
# #                     "a": "", "b": "", "c": "", "d": "",
# #                     "answer": "",
# #                     "extra_content": None,
# #                     "images": [],
# #                     "marks": marks
# #                 }

# #             elif re.match(r"^\(?[a-dA-D][\.\)]\s*", text) and current_question:
# #                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
# #                 if match:
# #                     current_question[match.group(1).lower()] = match.group(2).strip()

# #             elif re.match(r"^(answer|ans|correct answer)\s*:?", text, re.IGNORECASE) and current_question:
# #                 match = re.search(
# #                     r"^(?:answer|ans|correct answer)\s*:?\s*[\(\[]?([a-dA-D])[\)\]]?\b",
# #                     text,
# #                     re.IGNORECASE
# #                 )
# #                 if match:
# #                     current_question["answer"] = match.group(1).lower().strip()
# #                     print("RAW ANSWER LINE:", text, "=> PARSED:", current_question["answer"])

# #             else:
# #                 extra_html_parts.append(f"<p>{text}</p>")

# #             if current_question:
# #                 image_index = 0
# #                 for run in block.runs:
# #                     image_name = save_image_from_run(
# #                         run,
# #                         image_output_dir,
# #                         current_question["question_number"],
# #                         image_index + 1
# #                     )
# #                     if image_name:
# #                         image_index += 1
# #                         current_question["images"].append(image_name)

# #         elif isinstance(block, Table):
# #             table_html = extract_table_html(block)
# #             if current_question:
# #                 current_question["extra_content"] = (current_question.get("extra_content") or '') + table_html
# #             else:
# #                 extra_html_parts.append(table_html)

# #     if current_question:
# #         current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
# #         if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
# #             questions.append(current_question)

# #     print(f"✅ Parsed {len(questions)} questions successfully.")
# #     return questions

# import os
# import re
# from docx import Document
# from docx.oxml.ns import qn
# from docx.oxml.text.paragraph import CT_P
# from docx.oxml.table import CT_Tbl
# from docx.table import Table
# from docx.text.paragraph import Paragraph

# DEFAULT_IMAGE_DIR = "static/question_images"

# # ------------------ TABLE TO HTML ------------------
# def extract_table_html(table):
#     html = "<table border='1' cellspacing='0' cellpadding='5'>"
#     for row in table.rows:
#         html += "<tr>"
#         for cell in row.cells:
#             html += f"<td>{cell.text.strip()}</td>"
#         html += "</tr>"
#     html += "</table>"
#     return html

# # ------------------ IMAGE EXTRACTION ------------------
# def save_image_from_run(run, output_dir, question_number, image_index):
#     blip_elements = run._element.findall('.//a:blip', namespaces={
#         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
#     })
#     if not blip_elements:
#         return None

#     rId = blip_elements[0].get(qn('r:embed'))
#     image_part = run.part.related_parts[rId]
#     image_data = image_part.blob

#     image_filename = f"question_{question_number}_image_{image_index}.png"
#     image_path = os.path.join(output_dir, image_filename)
#     os.makedirs(output_dir, exist_ok=True)

#     with open(image_path, 'wb') as f:
#         f.write(image_data)

#     return image_filename

# # ------------------ ITERATE DOCX BLOCKS ------------------
# def iter_block_items(parent):
#     for child in parent.element.body.iterchildren():
#         if isinstance(child, CT_P):
#             yield Paragraph(child, parent)
#         elif isinstance(child, CT_Tbl):
#             yield Table(child, parent)

# # ------------------ MAIN PARSER ------------------
# def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
#     document = Document(file_stream)
#     questions = []
#     current_question = None
#     extra_html_parts = []
#     question_counter = 0

#     os.makedirs(image_output_dir, exist_ok=True)

#     for block in iter_block_items(document):
#         if isinstance(block, Paragraph):
#             text = block.text.strip()
#             if not text:
#                 continue

#             # ---------------- NEW QUESTION ----------------
#             if re.match(r"^\d+[\.\)]", text):
#                 if current_question:
#                     current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
#                     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
#                         questions.append(current_question)
#                     extra_html_parts = []

#                 question_counter += 1

#                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
#                 marks = int(marks_match.group(1)) if marks_match else 1

#                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
#                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text).strip()

#                 current_question = {
#                     "question_number": question_counter,
#                     "question": question_text,
#                     "a": "",
#                     "b": "",
#                     "c": "",
#                     "d": "",
#                     "answer": "",
#                     "extra_content": None,
#                     "images": [],
#                     "marks": marks
#                 }

#             # ---------------- OPTIONS ----------------
#             elif re.match(r"^\(?[a-dA-D][\.\)]\s*", text) and current_question:
#                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
#                 if match:
#                     current_question[match.group(1).lower()] = match.group(2).strip()

#             # ---------------- ANSWER ----------------
#             elif re.match(r"^(answer|ans|correct answer)\s*:?", text, re.IGNORECASE) and current_question:
#                 match = re.search(
#                     r"^(?:answer|ans|correct answer)\s*[:\-]?\s*[\(\[]?([a-dA-D])[\)\]]?\b",
#                     text,
#                     re.IGNORECASE
#                 )
#                 if match:
#                     current_question["answer"] = match.group(1).lower().strip()

#             # ---------------- EXTRA CONTENT ----------------
#             else:
#                 extra_html_parts.append(f"<p>{text}</p>")

#             # ---------------- IMAGE ATTACHMENT ----------------
#             if current_question:
#                 image_index = len(current_question["images"])
#                 for run in block.runs:
#                     image_name = save_image_from_run(
#                         run,
#                         image_output_dir,
#                         current_question["question_number"],
#                         image_index + 1
#                     )
#                     if image_name:
#                         image_index += 1
#                         current_question["images"].append(image_name)

#         # ---------------- TABLE ----------------
#         elif isinstance(block, Table):
#             table_html = extract_table_html(block)
#             if current_question:
#                 current_question["extra_content"] = (current_question.get("extra_content") or "") + table_html
#             else:
#                 extra_html_parts.append(table_html)

#     # ---------------- FINAL QUESTION ----------------
#     if current_question:
#         current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
#         if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
#             questions.append(current_question)

#     print(f"✅ Parsed {len(questions)} questions successfully.")
#     return questions

# # ------------------ QUIZ STATUS ------------------
# def get_quiz_status(user_id):
#     return "active"

# # ------------------ GOOGLE DRIVE HELPERS ------------------
# def extract_drive_id(url):
#     patterns = [
#         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
#         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
#     ]
#     for pattern in patterns:
#         match = re.search(pattern, url)
#         if match:
#             return match.group(1)
#     return url

# def get_drive_embed_url(drive_url_or_id):
#     file_id = extract_drive_id(drive_url_or_id)
#     return f"https://drive.google.com/file/d/{file_id}/preview"

#-------------------------working utils---------------------------------------------------

# import os
# import re
# from docx import Document
# from docx.oxml.ns import qn
# from docx.oxml.text.paragraph import CT_P
# from docx.oxml.table import CT_Tbl
# from docx.table import Table
# from docx.text.paragraph import Paragraph

# DEFAULT_IMAGE_DIR = "static/question_images"

# def extract_table_html(table):
#     html = "<table border='1' cellspacing='0' cellpadding='5'>"
#     for row in table.rows:
#         html += "<tr>"
#         for cell in row.cells:
#             html += f"<td>{cell.text.strip()}</td>"
#         html += "</tr>"
#     html += "</table>"
#     return html

# def save_image_from_run(run, output_dir, image_counter):
#     blip_elements = run._element.findall('.//a:blip', namespaces={
#         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
#     })

#     if not blip_elements:
#         return None

#     rId = blip_elements[0].get(qn('r:embed'))
#     image_part = run.part.related_parts[rId]
#     image_data = image_part.blob

#     image_filename = f"question_image_{image_counter}.png"
#     image_path = os.path.join(output_dir, image_filename)

#     with open(image_path, 'wb') as f:
#         f.write(image_data)

#     return image_filename

# def iter_block_items(parent):
#     """
#     Generator that yields paragraphs and tables in order from a docx document.
#     """
#     for child in parent.element.body.iterchildren():
#         if isinstance(child, CT_P):
#             yield Paragraph(child, parent)
#         elif isinstance(child, CT_Tbl):
#             yield Table(child, parent)

# def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
#     document = Document(file_stream)
#     questions = []
#     current_question = None
#     extra_html_parts = []
#     image_counter = 0
#     skipped = 0

#     os.makedirs(image_output_dir, exist_ok=True)

#     for block in iter_block_items(document):
#         if isinstance(block, Paragraph):
#             para = block
#             text = para.text.strip()

#             # Attach image to the current question
#             for run in para.runs:
#                 image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
#                 if image_name and current_question:
#                     image_counter += 1
#                     current_question["image"] = image_name

#             if not text:
#                 continue

#             # ✅ New question starts
#             if re.match(r"^\d+[\.\)]", text):
#                 if current_question:
#                     current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
#                     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
#                         questions.append(current_question)
#                     else:
#                         skipped += 1
#                     extra_html_parts = []

#                 # Extract marks
#                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
#                 marks = int(marks_match.group(1)) if marks_match else 1
#                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)

#                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)
#                 current_question = {
#                     "question": question_text,
#                     "a": "", "b": "", "c": "", "d": "",
#                     "answer": "",
#                     "extra_content": None,
#                     "image": None,
#                     "marks": marks
#                 }

#             # ✅ Option line (A., B., etc.)
#             elif re.match(r"^\(?[a-dA-D][\.\)]", text):
#                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
#                 if match and current_question:
#                     label = match.group(1).lower()
#                     content = match.group(2).strip()
#                     current_question[label] = content

#             # ✅ Answer line (e.g., Answer: B)
#             elif re.match(r"^(answer|correct answer):", text, re.IGNORECASE):
#                 match = re.search(r":\s*([a-dA-D])", text, re.IGNORECASE)
#                 if match:
#                     if current_question:
#                         current_question["answer"] = match.group(1).lower()
#                     else:
#                         print("⚠️ Found answer but no current question defined.")

#             # ✅ Extra content (instruction, explanation, etc.)
#             else:
#                 extra_html_parts.append(f"<p>{text}</p>")

#         elif isinstance(block, Table):
#             table_html = extract_table_html(block)
#             if current_question:
#                 current_question["extra_content"] = (current_question.get("extra_content") or '') + table_html
#             else:
#                 # No question yet, treat table as part of initial instruction
#                 extra_html_parts.append(table_html)

#     # ✅ Save final question
#     if current_question:
#         current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
#         if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
#             questions.append(current_question)
#         else:
#             skipped += 1

#     print(f"✅ Parsed {len(questions)} valid questions.")
#     if skipped > 0:
#         print(f"⚠️ Skipped {skipped} question(s) due to missing answers or invalid format.")

#     return questions

# # (Optional) Sample usage
# # with open("your_question.docx", "rb") as f:
# #     questions = parse_docx_questions(f)
# #     for q in questions:
# #         print(q["question"])
# def get_quiz_status(user_id):
#     # Placeholder implementation
#         return "active"
# # ------------------ GOOGLE DRIVE HELPERS ------------------
# def extract_drive_id(url):
#     patterns = [
#         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
#         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
#     ]
#     for pattern in patterns:
#         match = re.search(pattern, url)
#         if match:
#             return match.group(1)
#     return url


# def get_drive_embed_url(drive_url_or_id):
#     file_id = extract_drive_id(drive_url_or_id)
#     return f"https://drive.google.com/file/d/{file_id}/preview"

#---------------------------new utills -----------------------------------------

# import os
# import re
# from docx import Document
# from docx.oxml.ns import qn
# from docx.oxml.text.paragraph import CT_P
# from docx.oxml.table import CT_Tbl
# from docx.table import Table
# from docx.text.paragraph import Paragraph

# DEFAULT_IMAGE_DIR = "static/question_images"

# def extract_table_html(table):
#     html = "<table border='1' cellspacing='0' cellpadding='5'>"
#     for row in table.rows:
#         html += "<tr>"
#         for cell in row.cells:
#             # Preserve cell content including possible nested structure
#             cell_text = cell.text.strip() if cell.text else ""
#             html += f"<td>{cell_text}</td>"
#         html += "</tr>"
#     html += "</table>"
#     return html

# def save_image_from_run(run, output_dir, image_counter):
#     blip_elements = run._element.findall('.//a:blip', namespaces={
#         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
#     })

#     if not blip_elements:
#         return None

#     rId = blip_elements[0].get(qn('r:embed'))
#     image_part = run.part.related_parts[rId]
#     image_data = image_part.blob

#     image_filename = f"question_image_{image_counter}.png"
#     image_path = os.path.join(output_dir, image_filename)

#     with open(image_path, 'wb') as f:
#         f.write(image_data)

#     return image_filename

# def iter_block_items(document):
#     """
#     Generator that yields paragraphs and tables in order from a docx document.
#     FIXED: Properly iterate through document body elements
#     """
#     # Access the document body correctly
#     body = document.element.body
    
#     for child in body.iterchildren():
#         if child.tag.endswith('p'):  # Paragraph
#             yield Paragraph(child, document)
#         elif child.tag.endswith('tbl'):  # Table
#             yield Table(child, document)

# def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
#     document = Document(file_stream)
#     questions = []
#     current_question = None
#     extra_html_parts = []
#     image_counter = 0
#     skipped = 0

#     os.makedirs(image_output_dir, exist_ok=True)

#     # Debug: Print all blocks found
#     blocks = list(iter_block_items(document))
#     print(f"📄 Found {len(blocks)} total blocks (paragraphs + tables)")
    
#     for idx, block in enumerate(blocks):
#         print(f"Processing block {idx}: {type(block).__name__}")
        
#         if isinstance(block, Paragraph):
#             para = block
#             text = para.text.strip()
            
#             print(f"  Paragraph text: '{text[:50] if text else 'EMPTY'}'")

#             # Attach image to the current question
#             for run in para.runs:
#                 image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
#                 if image_name and current_question:
#                     image_counter += 1
#                     current_question["image"] = image_name
#                     print(f"  📸 Saved image: {image_name}")

#             if not text:
#                 continue

#             # ✅ New question starts
#             if re.match(r"^\d+[\.\)]", text):
#                 print(f"  🆕 New question detected")
#                 if current_question:
#                     current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
#                     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
#                         questions.append(current_question)
#                     else:
#                         skipped += 1
#                     extra_html_parts = []

#                 # Extract marks
#                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
#                 marks = int(marks_match.group(1)) if marks_match else 1
#                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)

#                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)
#                 current_question = {
#                     "question": question_text,
#                     "a": "", "b": "", "c": "", "d": "",
#                     "answer": "",
#                     "extra_content": None,
#                     "image": None,
#                     "marks": marks
#                 }
#                 print(f"  Question: {question_text[:50]}...")

#             # ✅ Option line (A., B., etc.)
#             elif re.match(r"^\(?[a-dA-D][\.\)]", text):
#                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
#                 if match and current_question:
#                     label = match.group(1).lower()
#                     content = match.group(2).strip()
#                     current_question[label] = content
#                     print(f"  Option {label}: {content[:30]}...")

#             # ✅ Answer line (e.g., Answer: B)
#             elif re.match(r"^(answer|correct answer):", text, re.IGNORECASE):
#                 match = re.search(r":\s*([a-dA-D])", text, re.IGNORECASE)
#                 if match:
#                     if current_question:
#                         current_question["answer"] = match.group(1).lower()
#                         print(f"  ✓ Answer: {match.group(1).lower()}")
#                     else:
#                         print("⚠️ Found answer but no current question defined.")

#             # ✅ Extra content (instruction, explanation, etc.)
#             else:
#                 print(f"  📝 Extra content added")
#                 extra_html_parts.append(f"<p>{text}</p>")

#         elif isinstance(block, Table):
#             print(f"  📊 Table detected")
#             table_html = extract_table_html(block)
#             if current_question:
#                 current_question["extra_content"] = (current_question.get("extra_content") or '') + table_html
#                 print(f"  Table added to current question")
#             else:
#                 # No question yet, treat table as part of initial instruction
#                 extra_html_parts.append(table_html)
#                 print(f"  Table added to extra content (no active question)")

#     # ✅ Save final question
#     if current_question:
#         current_question["extra_content"] = ''.join(extra_html_parts) if extra_html_parts else None
#         if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
#             questions.append(current_question)
#         else:
#             skipped += 1

#     print(f"\n✅ Parsed {len(questions)} valid questions.")
#     if skipped > 0:
#         print(f"⚠️ Skipped {skipped} question(s) due to missing answers or invalid format.")
    
#     # Print summary of parsed questions
#     for i, q in enumerate(questions, 1):
#         print(f"\nQuestion {i}: {q['question'][:50]}...")
#         print(f"  Options: a={q['a'][:30]}, b={q['b'][:30]}, c={q['c'][:30]}, d={q['d'][:30]}")
#         print(f"  Answer: {q['answer']}")
#         print(f"  Has extra: {bool(q['extra_content'])}")
#         print(f"  Has image: {bool(q['image'])}")

#     return questions

# def get_quiz_status(user_id):
#     # Placeholder implementation
#     return "active"

# # ------------------ GOOGLE DRIVE HELPERS ------------------
# def extract_drive_id(url):
#     patterns = [
#         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
#         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
#     ]
#     for pattern in patterns:
#         match = re.search(pattern, url)
#         if match:
#             return match.group(1)
#     return url

# def get_drive_embed_url(drive_url_or_id):
#     file_id = extract_drive_id(drive_url_or_id)
#     return f"https://drive.google.com/file/d/{file_id}/preview"

# import os
# import re
# import html
# from docx import Document
# from docx.oxml.ns import qn
# from docx.table import Table
# from docx.text.paragraph import Paragraph

# DEFAULT_IMAGE_DIR = "static/question_images"


# def extract_table_html(table):
#     """
#     Convert a python-docx table to simple HTML.
#     """
#     html_output = "<table border='1' cellspacing='0' cellpadding='5'>"

#     for row in table.rows:
#         html_output += "<tr>"
#         for cell in row.cells:
#             cell_text = cell.text.strip() if cell.text else ""
#             cell_text = html.escape(cell_text).replace("\n", "<br>")
#             html_output += f"<td>{cell_text}</td>"
#         html_output += "</tr>"

#     html_output += "</table>"
#     return html_output


# def save_image_from_run(run, output_dir, image_counter):
#     """
#     Extract image from a run and save it to disk.
#     Returns the saved filename, or None if no image is found.
#     """
#     blip_elements = run._element.findall(
#         './/a:blip',
#         namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
#     )

#     if not blip_elements:
#         return None

#     r_id = blip_elements[0].get(qn('r:embed'))
#     if not r_id:
#         return None

#     image_part = run.part.related_parts[r_id]
#     image_data = image_part.blob

#     image_filename = f"question_image_{image_counter}.png"
#     image_path = os.path.join(output_dir, image_filename)

#     with open(image_path, 'wb') as f:
#         f.write(image_data)

#     return image_filename


# def iter_block_items(document):
#     """
#     Yield paragraphs and tables in document order.
#     """
#     body = document.element.body

#     for child in body.iterchildren():
#         if child.tag.endswith('p'):
#             yield Paragraph(child, document)
#         elif child.tag.endswith('tbl'):
#             yield Table(child, document)


# def merge_extra_content(current_question, extra_html_parts):
#     """
#     Merge already-stored extra_content with new paragraph/table content.
#     This prevents tables from being overwritten later.
#     """
#     existing_extra = current_question.get("extra_content") or ""
#     new_extra = ''.join(extra_html_parts) if extra_html_parts else ""
#     merged = existing_extra + new_extra
#     current_question["extra_content"] = merged if merged else None


# def finalize_question(current_question, extra_html_parts, questions):
#     """
#     Finalize the current question:
#     - merge extra content
#     - validate required fields
#     - append to questions if valid
#     Returns True if saved, False if skipped.
#     """
#     if not current_question:
#         return False

#     merge_extra_content(current_question, extra_html_parts)

#     if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
#         questions.append(current_question)
#         return True

#     return False


# def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
#     document = Document(file_stream)
#     questions = []
#     current_question = None
#     extra_html_parts = []
#     image_counter = 0
#     skipped = 0

#     os.makedirs(image_output_dir, exist_ok=True)

#     blocks = list(iter_block_items(document))
#     print(f"📄 Found {len(blocks)} total blocks (paragraphs + tables)")

#     for idx, block in enumerate(blocks):
#         print(f"Processing block {idx}: {type(block).__name__}")

#         if isinstance(block, Paragraph):
#             para = block
#             text = para.text.strip()

#             print(f"  Paragraph text: '{text[:50] if text else 'EMPTY'}'")

#             # Save image(s) attached to current paragraph
#             for run in para.runs:
#                 image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
#                 if image_name and current_question:
#                     image_counter += 1
#                     current_question["image"] = image_name
#                     print(f"  📸 Saved image: {image_name}")

#             if not text:
#                 continue

#             # New question
#             if re.match(r"^\d+[\.\)]", text):
#                 print("  🆕 New question detected")

#                 if current_question:
#                     saved = finalize_question(current_question, extra_html_parts, questions)
#                     if not saved:
#                         skipped += 1
#                     extra_html_parts = []

#                 # Extract marks
#                 marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
#                 marks = int(marks_match.group(1)) if marks_match else 1

#                 # Remove marks from displayed question text
#                 clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text, flags=re.IGNORECASE)
#                 question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text).strip()

#                 current_question = {
#                     "question": question_text,
#                     "a": "",
#                     "b": "",
#                     "c": "",
#                     "d": "",
#                     "answer": "",
#                     "extra_content": None,
#                     "image": None,
#                     "marks": marks
#                 }

#                 print(f"  Question: {question_text[:50]}...")

#             # Option line: A. / A) / (A) etc.
#             elif re.match(r"^\(?[a-dA-D][\.\)]", text):
#                 match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
#                 if match and current_question:
#                     label = match.group(1).lower()
#                     content = match.group(2).strip()
#                     current_question[label] = content
#                     print(f"  Option {label}: {content[:30]}...")

#             # Answer line
#             elif re.match(r"^(answer|correct answer):", text, re.IGNORECASE):
#                 match = re.search(r":\s*([a-dA-D])", text, re.IGNORECASE)
#                 if match:
#                     if current_question:
#                         current_question["answer"] = match.group(1).lower()
#                         print(f"  ✓ Answer: {match.group(1).lower()}")
#                     else:
#                         print("⚠️ Found answer but no current question defined.")

#             # Any other paragraph becomes extra content
#             else:
#                 print("  📝 Extra content added")
#                 extra_html_parts.append(f"<p>{html.escape(text)}</p>")

#         elif isinstance(block, Table):
#             print("  📊 Table detected")
#             table_html = extract_table_html(block)

#             if current_question:
#                 current_question["extra_content"] = (current_question.get("extra_content") or "") + table_html
#                 print("  Table added to current question")
#             else:
#                 extra_html_parts.append(table_html)
#                 print("  Table added to extra content (no active question)")

#     # Final question
#     if current_question:
#         saved = finalize_question(current_question, extra_html_parts, questions)
#         if not saved:
#             skipped += 1

#     print(f"\n✅ Parsed {len(questions)} valid questions.")
#     if skipped > 0:
#         print(f"⚠️ Skipped {skipped} question(s) due to missing answers or invalid format.")

#     # Debug summary
#     for i, q in enumerate(questions, 1):
#         print(f"\nQuestion {i}: {q['question'][:50]}...")
#         print(f"  Options: a={q['a'][:30]}, b={q['b'][:30]}, c={q['c'][:30]}, d={q['d'][:30]}")
#         print(f"  Answer: {q['answer']}")
#         print(f"  Has extra: {bool(q['extra_content'])}")
#         print(f"  Has image: {bool(q['image'])}")

#     return questions


# def get_quiz_status(user_id):
#     # Placeholder implementation
#     return "active"


# # ------------------ GOOGLE DRIVE HELPERS ------------------

# def extract_drive_id(url):
#     patterns = [
#         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
#         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
#     ]

#     for pattern in patterns:
#         match = re.search(pattern, url)
#         if match:
#             return match.group(1)

#     return url


# def get_drive_embed_url(drive_url_or_id):
#     file_id = extract_drive_id(drive_url_or_id)
#     return f"https://drive.google.com/file/d/{file_id}/preview"


# import os
# import re
# import html
# from docx import Document
# from docx.oxml.ns import qn
# from docx.table import Table
# from docx.text.paragraph import Paragraph

# DEFAULT_IMAGE_DIR = "static/question_images"


# def normalize_text(text):
#     if text is None:
#         return ""
#     return text.replace("\xa0", " ").strip()


# def extract_table_text(table):
#     """
#     Convert a python-docx table to readable plain text.
#     Safer for parsers than HTML.
#     """
#     rows = []
#     for row in table.rows:
#         cells = [normalize_text(cell.text) for cell in row.cells]
#         rows.append(" | ".join(cells))
#     return "\n".join(rows).strip()


# def extract_table_html(table):
#     """
#     Optional HTML representation if you still want to display tables in the UI.
#     """
#     html_output = "<table border='1' cellspacing='0' cellpadding='5'>"
#     for row in table.rows:
#         html_output += "<tr>"
#         for cell in row.cells:
#             cell_text = normalize_text(cell.text)
#             cell_text = html.escape(cell_text).replace("\n", "<br>")
#             html_output += f"<td>{cell_text}</td>"
#         html_output += "</tr>"
#     html_output += "</table>"
#     return html_output


# def save_image_from_run(run, output_dir, image_counter):
#     """
#     Extract image from a run and save it to disk.
#     Returns the saved filename, or None if no image is found.
#     """
#     blip_elements = run._element.findall(
#         './/a:blip',
#         namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
#     )

#     if not blip_elements:
#         return None

#     r_id = blip_elements[0].get(qn('r:embed'))
#     if not r_id:
#         return None

#     image_part = run.part.related_parts[r_id]
#     image_data = image_part.blob

#     image_filename = f"question_image_{image_counter}.png"
#     image_path = os.path.join(output_dir, image_filename)

#     with open(image_path, 'wb') as f:
#         f.write(image_data)

#     return image_filename


# def iter_block_items(document):
#     """
#     Yield paragraphs and tables in document order.
#     """
#     body = document.element.body
#     for child in body.iterchildren():
#         if child.tag.endswith('p'):
#             yield Paragraph(child, document)
#         elif child.tag.endswith('tbl'):
#             yield Table(child, document)


# def is_question_start(text):
#     return bool(re.match(r"^\d+[\.\)]\s*", text))


# def is_option_start(text):
#     return bool(re.match(r"^\(?[A-Da-d][\.\)]\s*", text))


# def is_answer_line(text):
#     return bool(re.match(r"^(answer|correct answer|ans|answr)\s*:", text, re.IGNORECASE))


# def is_instruction_line(text):
#     """
#     Detect grouped instruction/case-study lead-ins.
#     """
#     lowered = text.lower().strip()

#     instruction_starts = (
#         "use the following",
#         "use matrices",
#         "use the information below",
#         "use the information to answer",
#         "refer to the following",
#         "answer question",
#         "answer questions",
#     )

#     return lowered.startswith(instruction_starts)


# def parse_answer_letter(text):
#     """
#     Supports:
#     Answer: A
#     Correct answer: B
#     Ans: C
#     Answr: D
#     """
#     match = re.search(
#         r"^(?:answer|correct answer|ans|answr)\s*:\s*([A-Da-d])\b",
#         text.strip(),
#         re.IGNORECASE
#     )
#     return match.group(1).lower() if match else ""


# def get_marks_from_text(text):
#     match = re.search(r"\((\d+)\s*(?:mks|marks?)\)", text, re.IGNORECASE)
#     if match:
#         return int(match.group(1))
#     return None


# def append_with_newline(existing, new_text):
#     new_text = normalize_text(new_text)
#     if not new_text:
#         return existing
#     if not existing:
#         return new_text
#     return existing + "\n" + new_text


# def start_new_question(raw_text, active_instruction):
#     marks = get_marks_from_text(raw_text)

#     return {
#         "question": normalize_text(raw_text),          # keep original question line
#         "question_full": normalize_text(raw_text),     # stem + continuation lines/tables
#         "instruction": normalize_text(active_instruction),
#         "a": "",
#         "b": "",
#         "c": "",
#         "d": "",
#         "answer": "",
#         "marks": marks if marks is not None else 1,
#         "extra_content": None,
#         "extra_html": None,
#         "image": None,
#         "images": [],
#     }


# def finalize_question(question, questions):
#     """
#     Save even if answer is blank.
#     Only require that some question text exists.
#     """
#     if not question:
#         return False

#     if normalize_text(question.get("question_full")):
#         # Build extra_content from instruction if you want one combined field
#         extra_parts = []

#         if question.get("instruction"):
#             extra_parts.append(question["instruction"])

#         if question.get("extra_content"):
#             extra_parts.append(question["extra_content"])

#         question["extra_content"] = "\n".join([p for p in extra_parts if p]).strip() or None
#         question["extra_html"] = question.get("extra_html") or None

#         questions.append(question)
#         return True

#     return False


# def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
#     document = Document(file_stream)
#     questions = []

#     current_question = None
#     current_option = None

#     # instruction blocks before a question group
#     active_instruction = ""

#     # loose intro text before first question
#     intro_buffer = []

#     image_counter = 0
#     skipped = 0

#     os.makedirs(image_output_dir, exist_ok=True)

#     blocks = list(iter_block_items(document))
#     print(f"📄 Found {len(blocks)} total blocks (paragraphs + tables)")

#     def close_current_question():
#         nonlocal current_question, current_option, skipped
#         if current_question:
#             saved = finalize_question(current_question, questions)
#             if not saved:
#                 skipped += 1
#         current_question = None
#         current_option = None

#     for idx, block in enumerate(blocks):
#         print(f"\nProcessing block {idx}: {type(block).__name__}")

#         if isinstance(block, Paragraph):
#             para = block
#             raw_text = para.text or ""
#             text = normalize_text(raw_text)

#             # Save paragraph images
#             for run in para.runs:
#                 image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
#                 if image_name:
#                     image_counter += 1
#                     if current_question:
#                         current_question["images"].append(image_name)
#                         if not current_question["image"]:
#                             current_question["image"] = image_name
#                         print(f"  📸 Saved image for current question: {image_name}")

#             if not text:
#                 continue

#             print(f"  Paragraph text: {text[:100]}")

#             # 1. New question
#             if is_question_start(text):
#                 print("  🆕 New question detected")
#                 close_current_question()

#                 # If there was intro text and no active instruction yet, attach it
#                 if intro_buffer and not active_instruction:
#                     active_instruction = "\n".join(intro_buffer).strip()
#                     intro_buffer = []

#                 current_question = start_new_question(text, active_instruction)
#                 current_option = None
#                 continue

#             # 2. New instruction block
#             if is_instruction_line(text):
#                 print("  📘 Instruction block detected")
#                 close_current_question()
#                 active_instruction = text
#                 current_option = None
#                 continue

#             # 3. Option line
#             option_match = re.match(r"^\(?([A-Da-d])[\.\)]\s*(.*)$", text)
#             if option_match and current_question:
#                 label = option_match.group(1).lower()
#                 content = option_match.group(2).strip()

#                 # Preserve the original option form as closely as possible
#                 current_question[label] = text
#                 current_option = label
#                 print(f"  🔠 Option {label.upper()} detected")
#                 continue

#             # 4. Answer line
#             if is_answer_line(text) and current_question:
#                 current_question["answer"] = parse_answer_letter(text)
#                 current_option = None
#                 print(f"  ✅ Answer detected: {current_question['answer'] or '(blank)'}")
#                 continue

#             # 5. Continuation lines
#             if current_question:
#                 if current_option:
#                     # line continues the current option
#                     current_question[current_option] = append_with_newline(
#                         current_question[current_option],
#                         text
#                     )
#                     print(f"  ↪ Continued option {current_option.upper()}")
#                 else:
#                     # line continues question stem
#                     current_question["question_full"] = append_with_newline(
#                         current_question["question_full"],
#                         text
#                     )
#                     print("  ↪ Continued question stem")
#             else:
#                 # no active question yet; treat as intro or instruction accumulation
#                 if active_instruction:
#                     active_instruction = append_with_newline(active_instruction, text)
#                     print("  ↪ Continued instruction block")
#                 else:
#                     intro_buffer.append(text)
#                     print("  ↪ Intro text buffered")

#         elif isinstance(block, Table):
#             print("  📊 Table detected")
#             table_text = extract_table_text(block)
#             table_html = extract_table_html(block)

#             if current_question:
#                 if current_option:
#                     current_question[current_option] = append_with_newline(
#                         current_question[current_option],
#                         table_text
#                     )
#                     print(f"  ↪ Table attached to option {current_option.upper()}")
#                 else:
#                     current_question["question_full"] = append_with_newline(
#                         current_question["question_full"],
#                         table_text
#                     )
#                     print("  ↪ Table attached to question stem")

#                 current_question["extra_html"] = (
#                     (current_question.get("extra_html") or "") + table_html
#                 )
#             else:
#                 if active_instruction:
#                     active_instruction = append_with_newline(active_instruction, table_text)
#                     print("  ↪ Table attached to instruction block")
#                 else:
#                     intro_buffer.append(table_text)
#                     print("  ↪ Table buffered as intro text")

#     close_current_question()

#     print(f"\n✅ Parsed {len(questions)} question(s)")
#     if skipped:
#         print(f"⚠️ Skipped {skipped} invalid question(s)")

#     for i, q in enumerate(questions, 1):
#         print(f"\nQuestion {i}")
#         print(f"  Question: {q['question'][:80]}")
#         print(f"  Full stem exists: {bool(q['question_full'])}")
#         print(f"  Instruction exists: {bool(q['instruction'])}")
#         print(f"  A: {q['a'][:60]}")
#         print(f"  B: {q['b'][:60]}")
#         print(f"  C: {q['c'][:60]}")
#         print(f"  D: {q['d'][:60]}")
#         print(f"  Answer: {q['answer'] or '(blank)'}")
#         print(f"  Images: {len(q['images'])}")

#     return questions


# def get_quiz_status(user_id):
#     return "active"


# # ------------------ GOOGLE DRIVE HELPERS ------------------

# def extract_drive_id(url):
#     patterns = [
#         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
#         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
#     ]

#     for pattern in patterns:
#         match = re.search(pattern, url)
#         if match:
#             return match.group(1)

#     return url


# def get_drive_embed_url(drive_url_or_id):
#     file_id = extract_drive_id(drive_url_or_id)
#     return f"https://drive.google.com/file/d/{file_id}/preview"


# import os
# import re
# import html
# from docx import Document
# from docx.oxml.ns import qn
# from docx.table import Table
# from docx.text.paragraph import Paragraph

# DEFAULT_IMAGE_DIR = "static/question_images"


# def normalize_text(text):
#     if text is None:
#         return ""
#     return text.replace("\xa0", " ").strip()


# def extract_table_text(table):
#     rows = []
#     for row in table.rows:
#         cells = [normalize_text(cell.text) for cell in row.cells]
#         rows.append(" | ".join(cells))
#     return "\n".join(rows).strip()


# def extract_table_html(table):
#     html_output = "<table border='1' cellspacing='0' cellpadding='5'>"
#     for row in table.rows:
#         html_output += "<tr>"
#         for cell in row.cells:
#             cell_text = normalize_text(cell.text)
#             cell_text = html.escape(cell_text).replace("\n", "<br>")
#             html_output += f"<td>{cell_text}</td>"
#         html_output += "</tr>"
#     html_output += "</table>"
#     return html_output


# def save_image_from_run(run, output_dir, image_counter):
#     blip_elements = run._element.findall(
#         './/a:blip',
#         namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
#     )

#     if not blip_elements:
#         return None

#     r_id = blip_elements[0].get(qn('r:embed'))
#     if not r_id:
#         return None

#     image_part = run.part.related_parts[r_id]
#     image_data = image_part.blob

#     image_filename = f"question_image_{image_counter}.png"
#     image_path = os.path.join(output_dir, image_filename)

#     with open(image_path, 'wb') as f:
#         f.write(image_data)

#     return image_filename


# def iter_block_items(document):
#     body = document.element.body
#     for child in body.iterchildren():
#         if child.tag.endswith('p'):
#             yield Paragraph(child, document)
#         elif child.tag.endswith('tbl'):
#             yield Table(child, document)


# def append_with_newline(existing, new_text):
#     new_text = normalize_text(new_text)
#     if not new_text:
#         return existing
#     if not existing:
#         return new_text
#     return existing + "\n" + new_text


# def is_question_start(text):
#     return bool(re.match(r"^\d+[\.\)]\s*", text))


# def is_option_start(text):
#     return bool(re.match(r"^\(?[A-Da-d][\.\)]\s*", text))


# def is_answer_line(text):
#     return bool(re.match(r"^(answer|correct answer|ans|answr)\s*:", text, re.IGNORECASE))


# def is_instruction_line(text):
#     lowered = text.lower().strip()
#     starters = (
#         "use the following",
#         "use matrices",
#         "use the information below",
#         "use the information to answer",
#         "refer to the following",
#         "answer question",
#         "answer questions",
#     )
#     return lowered.startswith(starters)


# def parse_answer_letter(text):
#     match = re.search(
#         r"^(?:answer|correct answer|ans|answr)\s*:\s*([A-Da-d])\b",
#         text.strip(),
#         re.IGNORECASE
#     )
#     return match.group(1).lower() if match else ""


# def get_marks_from_text(text):
#     match = re.search(r"\((\d+)\s*(?:mks|marks?)\)", text, re.IGNORECASE)
#     if match:
#         return int(match.group(1))
#     return 1


# def start_new_question(question_line, active_instruction):
#     full_question = question_line
#     if active_instruction:
#         full_question = active_instruction.strip() + "\n\n" + question_line.strip()

#     return {
#         "question": full_question.strip(),   # IMPORTANT: save everything in existing field
#         "a": "",
#         "b": "",
#         "c": "",
#         "d": "",
#         "answer": "",
#         "extra_content": None,
#         "image": None,
#         "marks": get_marks_from_text(question_line),
#     }


# def finalize_question(question, questions):
#     if not question:
#         return False

#     if normalize_text(question.get("question")):
#         questions.append(question)
#         return True

#     return False


# def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
#     document = Document(file_stream)
#     questions = []

#     current_question = None
#     current_option = None
#     active_instruction = ""
#     intro_buffer = []
#     image_counter = 0
#     skipped = 0

#     os.makedirs(image_output_dir, exist_ok=True)

#     blocks = list(iter_block_items(document))
#     print(f"📄 Found {len(blocks)} total blocks")

#     def close_current_question():
#         nonlocal current_question, current_option, skipped
#         if current_question:
#             saved = finalize_question(current_question, questions)
#             if not saved:
#                 skipped += 1
#         current_question = None
#         current_option = None

#     for idx, block in enumerate(blocks):
#         print(f"\nProcessing block {idx}: {type(block).__name__}")

#         if isinstance(block, Paragraph):
#             para = block
#             text = normalize_text(para.text)

#             for run in para.runs:
#                 image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
#                 if image_name and current_question:
#                     image_counter += 1
#                     current_question["image"] = image_name

#             if not text:
#                 continue

#             print(f"  Text: {text[:100]}")

#             # New grouped instruction
#             if is_instruction_line(text):
#                 close_current_question()
#                 active_instruction = text
#                 current_option = None
#                 print("  📘 Instruction detected")
#                 continue

#             # New question
#             if is_question_start(text):
#                 close_current_question()

#                 if intro_buffer and not active_instruction:
#                     active_instruction = "\n".join(intro_buffer).strip()
#                     intro_buffer = []

#                 current_question = start_new_question(text, active_instruction)
#                 current_option = None
#                 print("  🆕 Question detected")
#                 continue

#             # Option
#             option_match = re.match(r"^\(?([A-Da-d])[\.\)]\s*(.*)$", text)
#             if option_match and current_question:
#                 label = option_match.group(1).lower()

#                 # keep option exactly as it appears
#                 current_question[label] = text
#                 current_option = label
#                 print(f"  🔠 Option {label.upper()} detected")
#                 continue

#             # Answer
#             if is_answer_line(text) and current_question:
#                 current_question["answer"] = parse_answer_letter(text)
#                 current_option = None
#                 print(f"  ✅ Answer: {current_question['answer'] or '(blank)'}")
#                 continue

#             # Continuation lines
#             if current_question:
#                 if current_option:
#                     current_question[current_option] = append_with_newline(
#                         current_question[current_option],
#                         text
#                     )
#                     print(f"  ↪ Continued option {current_option.upper()}")
#                 else:
#                     current_question["question"] = append_with_newline(
#                         current_question["question"],
#                         text
#                     )
#                     print("  ↪ Continued question")
#             else:
#                 if active_instruction:
#                     active_instruction = append_with_newline(active_instruction, text)
#                     print("  ↪ Continued instruction")
#                 else:
#                     intro_buffer.append(text)
#                     print("  ↪ Buffered intro text")

#         elif isinstance(block, Table):
#             table_text = extract_table_text(block)
#             print("  📊 Table detected")

#             if current_question:
#                 if current_option:
#                     current_question[current_option] = append_with_newline(
#                         current_question[current_option],
#                         table_text
#                     )
#                     print(f"  ↪ Table added to option {current_option.upper()}")
#                 else:
#                     current_question["question"] = append_with_newline(
#                         current_question["question"],
#                         table_text
#                     )
#                     print("  ↪ Table added to question")
#             else:
#                 if active_instruction:
#                     active_instruction = append_with_newline(active_instruction, table_text)
#                     print("  ↪ Table added to instruction")
#                 else:
#                     intro_buffer.append(table_text)
#                     print("  ↪ Table buffered")

#     close_current_question()

#     print(f"\n✅ Parsed {len(questions)} question(s)")
#     if skipped:
#         print(f"⚠️ Skipped {skipped} question(s)")

#     for i, q in enumerate(questions, 1):
#         print(f"\nQuestion {i}")
#         print(f"QUESTION:\n{q['question'][:200]}")
#         print(f"A: {q['a'][:80]}")
#         print(f"B: {q['b'][:80]}")
#         print(f"C: {q['c'][:80]}")
#         print(f"D: {q['d'][:80]}")
#         print(f"ANSWER: {q['answer'] or '(blank)'}")

#     return questions


# def get_quiz_status(user_id):
#     return "active"


# def extract_drive_id(url):
#     patterns = [
#         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
#         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
#     ]

#     for pattern in patterns:
#         match = re.search(pattern, url)
#         if match:
#             return match.group(1)

#     return url


# def get_drive_embed_url(drive_url_or_id):
#     file_id = extract_drive_id(drive_url_or_id)
#     return f"https://drive.google.com/file/d/{file_id}/preview"


# import os
# import re
# from docx import Document
# from docx.oxml.ns import qn
# from docx.table import Table
# from docx.text.paragraph import Paragraph

# DEFAULT_IMAGE_DIR = "static/question_images"


# def normalize_text(text):
#     if text is None:
#         return ""
#     return text.replace("\xa0", " ").strip()


# def extract_table_text(table):
#     rows = []
#     for row in table.rows:
#         cells = [normalize_text(cell.text) for cell in row.cells]
#         rows.append(" | ".join(cells))
#     return "\n".join(rows).strip()


# def save_image_from_run(run, output_dir, image_counter):
#     blip_elements = run._element.findall(
#         './/a:blip',
#         namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
#     )

#     if not blip_elements:
#         return None

#     r_id = blip_elements[0].get(qn('r:embed'))
#     if not r_id:
#         return None

#     image_part = run.part.related_parts[r_id]
#     image_data = image_part.blob

#     image_filename = f"question_image_{image_counter}.png"
#     image_path = os.path.join(output_dir, image_filename)

#     with open(image_path, 'wb') as f:
#         f.write(image_data)

#     return image_filename


# def iter_block_items(document):
#     body = document.element.body
#     for child in body.iterchildren():
#         if child.tag.endswith('p'):
#             yield Paragraph(child, document)
#         elif child.tag.endswith('tbl'):
#             yield Table(child, document)


# def append_with_newline(existing, new_text):
#     new_text = normalize_text(new_text)
#     if not new_text:
#         return existing
#     if not existing:
#         return new_text
#     return existing + "\n" + new_text


# def is_question_start(text):
#     return bool(re.match(r"^\d+[\.\)]\s*", text))


# def is_answer_line(text):
#     return bool(re.match(r"^(answer|correct answer|ans|answr)\s*:", text, re.IGNORECASE))


# def is_instruction_line(text):
#     lowered = text.lower().strip()
#     starters = (
#         "use the following",
#         "use matrices",
#         "use the information below",
#         "use the information to answer",
#         "refer to the following",
#         "answer question",
#         "answer questions",
#     )
#     return lowered.startswith(starters)


# def parse_answer_letter(text):
#     match = re.search(
#         r"^(?:answer|correct answer|ans|answr)\s*:\s*([A-Da-d])\b",
#         text.strip(),
#         re.IGNORECASE
#     )
#     return match.group(1).lower() if match else ""


# def get_marks_from_text(text):
#     match = re.search(r"\((\d+)\s*(?:mks|marks?)\)", text, re.IGNORECASE)
#     if match:
#         return int(match.group(1))
#     return 1


# def finalize_question(question, questions):
#     if question and normalize_text(question.get("question")):
#         questions.append(question)
#         return True
#     return False


# def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
#     document = Document(file_stream)
#     questions = []

#     current_question = None
#     current_option = None
#     current_shared_block = ""   # instruction + table block for upcoming questions
#     image_counter = 0
#     skipped = 0

#     os.makedirs(image_output_dir, exist_ok=True)

#     blocks = list(iter_block_items(document))
#     print(f"📄 Found {len(blocks)} total blocks")

#     def close_current_question():
#         nonlocal current_question, current_option, skipped
#         if current_question:
#             saved = finalize_question(current_question, questions)
#             if not saved:
#                 skipped += 1
#         current_question = None
#         current_option = None

#     for idx, block in enumerate(blocks):
#         print(f"\nProcessing block {idx}: {type(block).__name__}")

#         if isinstance(block, Paragraph):
#             text = normalize_text(block.text)

#             # images
#             for run in block.runs:
#                 image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
#                 if image_name and current_question:
#                     image_counter += 1
#                     current_question["image"] = image_name

#             if not text:
#                 continue

#             print("TEXT:", text)

#             # Start of a new shared instruction block
#             if is_instruction_line(text):
#                 close_current_question()
#                 current_shared_block = text
#                 current_option = None
#                 print("📘 Started shared block")
#                 continue

#             # New question
#             if is_question_start(text):
#                 close_current_question()

#                 full_question = text
#                 if current_shared_block:
#                     full_question = current_shared_block + "\n\n" + text

#                 current_question = {
#                     "question": full_question,
#                     "a": "",
#                     "b": "",
#                     "c": "",
#                     "d": "",
#                     "answer": "",
#                     "extra_content": None,
#                     "image": None,
#                     "marks": get_marks_from_text(text),
#                 }
#                 current_option = None
#                 print("🆕 New question started")
#                 continue

#             # Option
#             option_match = re.match(r"^\(?([A-Da-d])[\.\)]\s*(.*)$", text)
#             if option_match and current_question:
#                 label = option_match.group(1).lower()
#                 current_question[label] = text
#                 current_option = label
#                 print(f"🔠 Option {label.upper()}")
#                 continue

#             # Answer
#             if is_answer_line(text) and current_question:
#                 current_question["answer"] = parse_answer_letter(text)
#                 current_option = None
#                 print(f"✅ Answer: {current_question['answer'] or '(blank)'}")
#                 continue

#             # Continuation
#             if current_question:
#                 if current_option:
#                     current_question[current_option] = append_with_newline(
#                         current_question[current_option], text
#                     )
#                     print(f"↪ Continued option {current_option.upper()}")
#                 else:
#                     current_question["question"] = append_with_newline(
#                         current_question["question"], text
#                     )
#                     print("↪ Continued current question")
#             else:
#                 # THIS IS THE IMPORTANT PART:
#                 # any text before the next numbered question belongs to the shared block
#                 current_shared_block = append_with_newline(current_shared_block, text)
#                 print("↪ Added to shared block")

#         elif isinstance(block, Table):
#             table_text = extract_table_text(block)
#             if not table_text:
#                 continue

#             print("📊 Table detected")

#             if current_question:
#                 if current_option:
#                     current_question[current_option] = append_with_newline(
#                         current_question[current_option], table_text
#                     )
#                     print(f"↪ Table added to option {current_option.upper()}")
#                 else:
#                     current_question["question"] = append_with_newline(
#                         current_question["question"], table_text
#                     )
#                     print("↪ Table added to current question")
#             else:
#                 # THIS IS ALSO IMPORTANT:
#                 # table before next question belongs to shared block
#                 current_shared_block = append_with_newline(current_shared_block, table_text)
#                 print("↪ Table added to shared block")

#     close_current_question()

#     print(f"\n✅ Parsed {len(questions)} question(s)")
#     if skipped:
#         print(f"⚠️ Skipped {skipped} question(s)")

#     for i, q in enumerate(questions, 1):
#         print(f"\nQuestion {i}")
#         print(q["question"][:300])
#         print("A:", q["a"])
#         print("B:", q["b"])
#         print("C:", q["c"])
#         print("D:", q["d"])
#         print("Answer:", q["answer"] or "(blank)")

#     return questions


# def get_quiz_status(user_id):
#     return "active"


# def extract_drive_id(url):
#     patterns = [
#         r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
#         r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
#     ]

#     for pattern in patterns:
#         match = re.search(pattern, url)
#         if match:
#             return match.group(1)

#     return url


# def get_drive_embed_url(drive_url_or_id):
#     file_id = extract_drive_id(drive_url_or_id)
#     return f"https://drive.google.com/file/d/{file_id}/preview"

import os
import re
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

DEFAULT_IMAGE_DIR = "static/question_images"


def normalize_text(text):
    if text is None:
        return ""
    return text.replace("\xa0", " ").strip()


def extract_table_text(table):
    rows = []
    for row in table.rows:
        cells = [normalize_text(cell.text) for cell in row.cells]
        rows.append(" | ".join(cells))
    return "\n".join(rows).strip()


def save_image_from_run(run, output_dir, image_counter):
    blip_elements = run._element.findall(
        './/a:blip',
        namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    )

    if not blip_elements:
        return None

    r_id = blip_elements[0].get(qn('r:embed'))
    if not r_id:
        return None

    image_part = run.part.related_parts[r_id]
    image_data = image_part.blob

    image_filename = f"question_image_{image_counter}.png"
    image_path = os.path.join(output_dir, image_filename)

    with open(image_path, 'wb') as f:
        f.write(image_data)

    return image_filename


def iter_block_items(document):
    body = document.element.body
    for child in body.iterchildren():
        if child.tag.endswith('p'):
            yield Paragraph(child, document)
        elif child.tag.endswith('tbl'):
            yield Table(child, document)


def append_with_newline(existing, new_text):
    new_text = normalize_text(new_text)
    if not new_text:
        return existing
    if not existing:
        return new_text
    return existing + "\n" + new_text


def is_question_start(text):
    return bool(re.match(r"^\d+[\.\)]\s*", text))


def is_answer_line(text):
    return bool(re.match(r"^(answer|correct answer|ans|answr)\s*:", text, re.IGNORECASE))


def is_instruction_line(text):
    lowered = text.lower().strip()
    starters = (
        "use the following",
        "use matrices",
        "use the information below",
        "use the information to answer",
        "refer to the following",
        "answer question",
        "answer questions",
    )
    return lowered.startswith(starters)


def parse_answer_letter(text):
    match = re.search(
        r"^(?:answer|correct answer|ans|answr)\s*:\s*([A-Da-d])\b",
        text.strip(),
        re.IGNORECASE
    )
    return match.group(1).lower() if match else ""


def get_marks_from_text(text):
    match = re.search(r"\((\d+)\s*(?:mks|marks?)\)", text, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return 1


def finalize_question(question, questions):
    if question and normalize_text(question.get("question")):
        questions.append(question)
        return True
    return False


def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
    document = Document(file_stream)
    questions = []

    current_question = None
    current_option = None
    current_shared_block = ""
    image_counter = 0
    skipped = 0

    os.makedirs(image_output_dir, exist_ok=True)

    blocks = list(iter_block_items(document))
    print(f"📄 Found {len(blocks)} total blocks")

    def close_current_question():
        nonlocal current_question, current_option, skipped
        if current_question:
            saved = finalize_question(current_question, questions)
            if not saved:
                skipped += 1
        current_question = None
        current_option = None

    for idx, block in enumerate(blocks):
        print(f"\nProcessing block {idx}: {type(block).__name__}")

        if isinstance(block, Paragraph):
            text = normalize_text(block.text)

            # Save images
            for run in block.runs:
                image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
                if image_name and current_question:
                    image_counter += 1
                    current_question["image"] = image_name
                    print(f"📸 Saved image: {image_name}")

            if not text:
                continue

            print("TEXT:", text)

            # Start shared instruction block
            if is_instruction_line(text):
                close_current_question()
                current_shared_block = text
                current_option = None
                print("📘 Started shared block")
                continue

            # New question
            if is_question_start(text):
                close_current_question()

                full_question = text
                if current_shared_block:
                    full_question = current_shared_block + "\n\n" + text
                    current_shared_block = ""   # clear after first use

                current_question = {
                    "question": full_question,
                    "a": "",
                    "b": "",
                    "c": "",
                    "d": "",
                    "answer": "",
                    "extra_content": None,
                    "image": None,
                    "marks": get_marks_from_text(text),
                }
                current_option = None
                print("🆕 New question started")
                continue

            # Option
            option_match = re.match(r"^\(?([A-Da-d])[\.\)]\s*(.*)$", text)
            if option_match and current_question:
                label = option_match.group(1).lower()
                current_question[label] = text
                current_option = label
                print(f"🔠 Option {label.upper()}")
                continue

            # Answer
            if is_answer_line(text) and current_question:
                current_question["answer"] = parse_answer_letter(text)
                current_option = None
                print(f"✅ Answer: {current_question['answer'] or '(blank)'}")
                continue

            # Continuation
            if current_question:
                if current_option:
                    current_question[current_option] = append_with_newline(
                        current_question[current_option], text
                    )
                    print(f"↪ Continued option {current_option.upper()}")
                else:
                    current_question["question"] = append_with_newline(
                        current_question["question"], text
                    )
                    print("↪ Continued current question")
            else:
                current_shared_block = append_with_newline(current_shared_block, text)
                print("↪ Added to shared block")

        elif isinstance(block, Table):
            table_text = extract_table_text(block)
            if not table_text:
                continue

            print("📊 Table detected")

            if current_question:
                if current_option:
                    current_question[current_option] = append_with_newline(
                        current_question[current_option], table_text
                    )
                    print(f"↪ Table added to option {current_option.upper()}")
                else:
                    current_question["question"] = append_with_newline(
                        current_question["question"], table_text
                    )
                    print("↪ Table added to current question")
            else:
                current_shared_block = append_with_newline(current_shared_block, table_text)
                print("↪ Table added to shared block")

    close_current_question()

    print(f"\n✅ Parsed {len(questions)} question(s)")
    if skipped:
        print(f"⚠️ Skipped {skipped} question(s)")

    for i, q in enumerate(questions, 1):
        print(f"\nQuestion {i}")
        print(q["question"][:300])
        print("A:", q["a"])
        print("B:", q["b"])
        print("C:", q["c"])
        print("D:", q["d"])
        print("Answer:", q["answer"] or "(blank)")

    return questions


def get_quiz_status(user_id):
    return "active"


def extract_drive_id(url):
    patterns = [
        r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
        r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)"
    ]

    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)

    return url


def get_drive_embed_url(drive_url_or_id):
    file_id = extract_drive_id(drive_url_or_id)
    return f"https://drive.google.com/file/d/{file_id}/preview"