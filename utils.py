# # # # # # # # # # 
# # # # # # # # # import docx
# # # # # # # # # import re
# # # # # # # # # import os
# # # # # # # # # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # LOAD DOCX
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def load_docx(path):
# # # # # # # # #     return docx.Document(path)

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # HELPER MATCHERS
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def is_question_line(text):
# # # # # # # # #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # # # # # # # # def is_option_line(text):
# # # # # # # # #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # # # # # # # # def is_answer_line(text):
# # # # # # # # #     return text.lower().startswith(("answer", "ans", "correct"))

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # IMAGE EXTRACTION
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def extract_images(document, output_dir, q_index):
# # # # # # # # #     os.makedirs(output_dir, exist_ok=True)
# # # # # # # # #     images = []
# # # # # # # # #     count = 0
# # # # # # # # #     for rel in document.part.rels.values():
# # # # # # # # #         if rel.reltype == RT.IMAGE:
# # # # # # # # #             count += 1
# # # # # # # # #             ext = rel.target_ref.split('.')[-1]
# # # # # # # # #             filename = f"q{q_index}_img{count}.{ext}"
# # # # # # # # #             filepath = os.path.join(output_dir, filename)
# # # # # # # # #             with open(filepath, "wb") as f:
# # # # # # # # #                 f.write(rel.target_part.blob)
# # # # # # # # #             images.append(filename)
# # # # # # # # #     return images

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # HTML TABLE BUILDER
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def make_html_table(cells):
# # # # # # # # #     html = "<table class='table table-bordered'><tr>"
# # # # # # # # #     for c in cells:
# # # # # # # # #         html += f"<td>{c}</td>"
# # # # # # # # #     html += "</tr></table>"
# # # # # # # # #     return html

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # FLATTEN DOCX (PARAGRAPHS + TABLES)
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def flatten_doc(document):
# # # # # # # # #     lines = []

# # # # # # # # #     # Preserve exact order of paragraphs + tables
# # # # # # # # #     for block in document.element.body:
# # # # # # # # #         # Paragraph
# # # # # # # # #         if block.tag.endswith('p'):
# # # # # # # # #             para = docx.text.paragraph.Paragraph(block, document)
# # # # # # # # #             text = para.text.strip()
# # # # # # # # #             if text:
# # # # # # # # #                 lines.append({"type": "text", "content": text})

# # # # # # # # #         # Table
# # # # # # # # #         elif block.tag.endswith('tbl'):
# # # # # # # # #             table = docx.table.Table(block, document)
# # # # # # # # #             rows = []
# # # # # # # # #             for row in table.rows:
# # # # # # # # #                 cells = [c.text.strip() for c in row.cells if c.text.strip()]
# # # # # # # # #                 if cells:
# # # # # # # # #                     rows.append(cells)
# # # # # # # # #             if rows:
# # # # # # # # #                 lines.append({"type": "table", "cells": rows})

# # # # # # # # #     return lines(document):
# # # # # # # # #     lines = []
# # # # # # # # #     # Paragraphs
# # # # # # # # #     for p in document.paragraphs:
# # # # # # # # #         if p.text.strip():
# # # # # # # # #             lines.append({"type": "text", "content": p.text.strip()})
# # # # # # # # #     # Tables
# # # # # # # # #     for table in document.tables:
# # # # # # # # #         for row in table.rows:
# # # # # # # # #             row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
# # # # # # # # #             if row_text:
# # # # # # # # #                 lines.append({"type": "table", "cells": row_text})
# # # # # # # # #     return lines

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # PARSE DOCX QUESTIONS (WITH STRICT CASE STUDY SEPARATION)
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # NEW RULE:
# # # # # # # # # # 1) Anything after a question line BUT before first option is still part of QUESTION STEM.
# # # # # # # # # # 2) Case study begins ONLY when line starts with known keywords OR appears BETWEEN questions
# # # # # # # # # #    such as: "Use the following information to answer...".
# # # # # # # # # # 3) Case study must attach to NEXT question, not previous.
# # # # # # # # # # ---------------------------------------------------------

# # # # # # # # #  (WITH CASE STUDY + TABLES)
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def is_case_study_line(text):
# # # # # # # # #     keywords = [
# # # # # # # # #         "use the following information",
# # # # # # # # #         "study the information",
# # # # # # # # #         "refer to the following",
# # # # # # # # #         "case study",
# # # # # # # # #         "use the data below"
# # # # # # # # #     ]
# # # # # # # # #     t = text.lower()
# # # # # # # # #     return any(k in t for k in keywords)


# # # # # # # # # def parse_docx_questions(path, image_output_dir=None):(path, image_output_dir=None):
# # # # # # # # #     doc = load_docx(path)
# # # # # # # # #     entries = flatten_doc(doc)

# # # # # # # # #     questions = []
# # # # # # # # #     current = None
# # # # # # # # #     q_index = 0
# # # # # # # # #     current_case_study = ""

# # # # # # # # #     for entry in entries:

# # # # # # # # #         # --------------------------------------------------
# # # # # # # # #         # TEXT ENTRY
# # # # # # # # #         # --------------------------------------------------
# # # # # # # # #         if entry["type"] == "text":
# # # # # # # # #             line = entry["content"]

# # # # # # # # #             # ---------- NEW QUESTION ----------
# # # # # # # # #             if is_question_line(line):
# # # # # # # # #                 if current:
# # # # # # # # #                     questions.append(current)

# # # # # # # # #                 q_index += 1
# # # # # # # # #                 question_text = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# # # # # # # # #                 mk = re.search(r"\((\d+)\s*mks?\)", line, re.IGNORECASE)
# # # # # # # # #                 marks = int(mk.group(1)) if mk else 1

# # # # # # # # #                 current = {
# # # # # # # # #                     "question": question_text,
# # # # # # # # #                     "instructions": current_case_study.strip(),
# # # # # # # # #                     "a": "",
# # # # # # # # #                     "b": "",
# # # # # # # # #                     "c": "",
# # # # # # # # #                     "d": "",
# # # # # # # # #                     "answer": "",
# # # # # # # # #                     "marks": marks,
# # # # # # # # #                     "image": None
# # # # # # # # #                 }

# # # # # # # # #                 # Extract images
# # # # # # # # #                 if image_output_dir:
# # # # # # # # #                     imgs = extract_images(doc, image_output_dir, q_index)
# # # # # # # # #                     if imgs:
# # # # # # # # #                         current["image"] = imgs[0]

# # # # # # # # #                 # Reset case study
# # # # # # # # #                 current_case_study = ""

# # # # # # # # #             # ---------- OPTIONS ----------
# # # # # # # # #             elif current and is_option_line(line):
# # # # # # # # #                 letter = line[0].lower()
# # # # # # # # #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# # # # # # # # #                 current[letter] = text

# # # # # # # # #             # ---------- ANSWER ----------
# # # # # # # # #             elif current and is_answer_line(line):
# # # # # # # # #                 raw = line.split(":")[-1].strip().lower()
# # # # # # # # #                 clean = re.sub(r"[^a-d]", "", raw)
# # # # # # # # #                 current["answer"] = clean

# # # # # # # # #             # ---------- CASE STUDY OR EXTRA TEXT ----------
# # # # # # # # #             else:
# # # # # # # # #                 # BEFORE first question → case study
# # # # # # # # #                 if not current:
# # # # # # # # #                     if current_case_study:
# # # # # # # # #                         current_case_study += "<br>" + line.strip()
# # # # # # # # #                     else:
# # # # # # # # #                         current_case_study = line.strip()
# # # # # # # # #                 else:
# # # # # # # # #                     # text inside question
# # # # # # # # #                     current["question"] += " " + line.strip()

# # # # # # # # #         # --------------------------------------------------
# # # # # # # # #         # TABLE ENTRY
# # # # # # # # #         # --------------------------------------------------
# # # # # # # # #         elif entry["type"] == "table":
# # # # # # # # #             html_table = make_html_table(entry["cells"])

# # # # # # # # #             if not current:
# # # # # # # # #                 if current_case_study:
# # # # # # # # #                     current_case_study += "<br>" + html_table
# # # # # # # # #                 else:
# # # # # # # # #                     current_case_study = html_table
# # # # # # # # #             else:
# # # # # # # # #                 current["question"] += "<br>" + html_table

# # # # # # # # #     # Push last question
# # # # # # # # #     if current:
# # # # # # # # #         questions.append(current)

# # # # # # # # #     return questions

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # SCORING ENGINE
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def compute_score(questions, student_answers):
# # # # # # # # #     score = 0
# # # # # # # # #     total_marks = 0
# # # # # # # # #     details = []

# # # # # # # # #     for index, q in enumerate(questions, start=1):
# # # # # # # # #         correct = q["answer"].strip().lower()
# # # # # # # # #         total_marks += q["marks"]

# # # # # # # # #         student_answer = ""
# # # # # # # # #         for key in (index, f"q{index}", q["question"]):
# # # # # # # # #             if key in student_answers:
# # # # # # # # #                 student_answer = student_answers[key].strip().lower()
# # # # # # # # #                 break

# # # # # # # # #         got_it = student_answer == correct
# # # # # # # # #         if got_it:
# # # # # # # # #             score += q["marks"]

# # # # # # # # #         details.append({
# # # # # # # # #             "question": q["question"],
# # # # # # # # #             "correct": correct,
# # # # # # # # #             "student_answer": student_answer,
# # # # # # # # #             "marks": q["marks"],
# # # # # # # # #             "earned": q["marks"] if got_it else 0
# # # # # # # # #         })

# # # # # # # # #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# # # # # # # # #     return {
# # # # # # # # #         "score": score,
# # # # # # # # #         "total": total_marks,
# # # # # # # # #         "percentage": percentage,
# # # # # # # # #         "details": details
# # # # # # # # #     }

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # QUIZ STATUS
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def get_quiz_status(questions, student_answers):
# # # # # # # # #     status_list = []

# # # # # # # # #     for index, q in enumerate(questions, start=1):
# # # # # # # # #         correct = q["answer"].strip().lower()

# # # # # # # # #         student_answer = ""
# # # # # # # # #         for key in (index, f"q{index}", q["question"]):
# # # # # # # # #             if key in student_answers:
# # # # # # # # #                 student_answer = student_answers[key].strip().lower()
# # # # # # # # #                 break

# # # # # # # # #         if not student_answer:
# # # # # # # # #             status = "unanswered"
# # # # # # # # #         elif student_answer == correct:
# # # # # # # # #             status = "correct"
# # # # # # # # #         else:
# # # # # # # # #             status = "incorrect"

# # # # # # # # #         status_list.append({
# # # # # # # # #             "question_index": index,
# # # # # # # # #             "status": status,
# # # # # # # # #             "student_answer": student_answer,
# # # # # # # # #             "correct_answer": correct
# # # # # # # # #         })

# # # # # # # # #     return status_list

# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # # GOOGLE DRIVE HELPERS
# # # # # # # # # # ---------------------------------------------------------
# # # # # # # # # def extract_drive_id(url):
# # # # # # # # #     patterns = [
# # # # # # # # #         r"https://drive\\.google\\.com/file/d/([a-zA-Z0-9_-]+)",
# # # # # # # # #         r"https://drive\\.google\\.com/open\\?id=([a-zA-Z0-9_-]+)"
# # # # # # # # #     ]
# # # # # # # # #     for pattern in patterns:
# # # # # # # # #         m = re.search(pattern, url)
# # # # # # # # #         if m:
# # # # # # # # #             return m.group(1)
# # # # # # # # #     return url

# # # # # # # # # def get_drive_embed_url(drive_url_or_id):
# # # # # # # # #     file_id = extract_drive_id(drive_url_or_id)
# # # # # # # # #     return f"https://drive.google.com/file/d/{file_id}/preview"

# # # # # # # # import docx
# # # # # # # # import re
# # # # # # # # import os
# # # # # # # # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # LOAD DOCX
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def load_docx(path):
# # # # # # # #     return docx.Document(path)

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # HELPER MATCHERS
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def is_question_line(text):
# # # # # # # #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # # # # # # # def is_option_line(text):
# # # # # # # #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # # # # # # # def is_answer_line(text):
# # # # # # # #     return text.lower().startswith(("answer", "ans", "correct"))

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # IMAGE EXTRACTION
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def extract_images(document, output_dir, q_index):
# # # # # # # #     os.makedirs(output_dir, exist_ok=True)
# # # # # # # #     images = []
# # # # # # # #     count = 0
# # # # # # # #     for rel in document.part.rels.values():
# # # # # # # #         if rel.reltype == RT.IMAGE:
# # # # # # # #             count += 1
# # # # # # # #             ext = rel.target_ref.split('.')[-1]
# # # # # # # #             filename = f"q{q_index}_img{count}.{ext}"
# # # # # # # #             filepath = os.path.join(output_dir, filename)
# # # # # # # #             with open(filepath, "wb") as f:
# # # # # # # #                 f.write(rel.target_part.blob)
# # # # # # # #             images.append(filename)
# # # # # # # #     return images

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # HTML TABLE BUILDER
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def make_html_table(cells):
# # # # # # # #     html = "<table class='table table-bordered'>"
# # # # # # # #     for row in cells:
# # # # # # # #         html += "<tr>"
# # # # # # # #         for c in row:
# # # # # # # #             html += f"<td>{c}</td>"
# # # # # # # #         html += "</tr>"
# # # # # # # #     html += "</table>"
# # # # # # # #     return html

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # FLATTEN DOCX (PARAGRAPHS + TABLES) – FIXED
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def flatten_doc(document):
# # # # # # # #     lines = []

# # # # # # # #     for block in document.element.body:
# # # # # # # #         # Paragraph
# # # # # # # #         if block.tag.endswith('p'):
# # # # # # # #             para = docx.text.paragraph.Paragraph(block, document)
# # # # # # # #             text = para.text.strip()
# # # # # # # #             if text:
# # # # # # # #                 lines.append({"type": "text", "content": text})

# # # # # # # #         # Table
# # # # # # # #         elif block.tag.endswith('tbl'):
# # # # # # # #             table = docx.table.Table(block, document)
# # # # # # # #             rows = []
# # # # # # # #             for row in table.rows:
# # # # # # # #                 cells = [c.text.strip() for c in row.cells]
# # # # # # # #                 rows.append(cells)
# # # # # # # #             lines.append({"type": "table", "cells": rows})

# # # # # # # #     return lines

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # CASE STUDY CHECKER
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def is_case_study_line(text):
# # # # # # # #     keywords = [
# # # # # # # #         "use the following information",
# # # # # # # #         "study the information",
# # # # # # # #         "refer to the following",
# # # # # # # #         "case study",
# # # # # # # #         "use the data below"
# # # # # # # #     ]
# # # # # # # #     t = text.lower()
# # # # # # # #     return any(k in t for k in keywords)

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # PARSE DOCX QUESTIONS – FIXED
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def parse_docx_questions(path, image_output_dir=None):
# # # # # # # #     doc = load_docx(path)
# # # # # # # #     entries = flatten_doc(doc)

# # # # # # # #     questions = []
# # # # # # # #     current = None
# # # # # # # #     q_index = 0
# # # # # # # #     current_case_study = ""

# # # # # # # #     for entry in entries:

# # # # # # # #         # --------------------------------------------------
# # # # # # # #         # TEXT ENTRY
# # # # # # # #         # --------------------------------------------------
# # # # # # # #         if entry["type"] == "text":
# # # # # # # #             line = entry["content"]

# # # # # # # #             # ---------- NEW QUESTION ----------
# # # # # # # #             if is_question_line(line):
# # # # # # # #                 if current:
# # # # # # # #                     questions.append(current)

# # # # # # # #                 q_index += 1
# # # # # # # #                 question_text = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# # # # # # # #                 mk = re.search(r"\((\d+)\s*mks?\)", line, re.IGNORECASE)
# # # # # # # #                 marks = int(mk.group(1)) if mk else 1

# # # # # # # #                 current = {
# # # # # # # #                     "question": question_text,
# # # # # # # #                     "instructions": current_case_study.strip(),
# # # # # # # #                     "a": "",
# # # # # # # #                     "b": "",
# # # # # # # #                     "c": "",
# # # # # # # #                     "d": "",
# # # # # # # #                     "answer": "",
# # # # # # # #                     "marks": marks,
# # # # # # # #                     "image": None
# # # # # # # #                 }

# # # # # # # #                 if image_output_dir:
# # # # # # # #                     imgs = extract_images(doc, image_output_dir, q_index)
# # # # # # # #                     if imgs:
# # # # # # # #                         current["image"] = imgs[0]

# # # # # # # #                 # Reset case study
# # # # # # # #                 current_case_study = ""

# # # # # # # #             # ---------- OPTIONS ----------
# # # # # # # #             elif current and is_option_line(line):
# # # # # # # #                 letter = line[0].lower()
# # # # # # # #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# # # # # # # #                 current[letter] = text

# # # # # # # #             # ---------- ANSWER ----------
# # # # # # # #             elif current and is_answer_line(line):
# # # # # # # #                 raw = line.split(":")[-1].strip().lower()
# # # # # # # #                 clean = re.sub(r"[^a-d]", "", raw)
# # # # # # # #                 current["answer"] = clean

# # # # # # # #             # ---------- CASE STUDY or EXTRA TEXT ----------
# # # # # # # #             else:
# # # # # # # #                 if not current:
# # # # # # # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
# # # # # # # #                 else:
# # # # # # # #                     current["question"] += " " + line.strip()

# # # # # # # #         # --------------------------------------------------
# # # # # # # #         # TABLE ENTRY – FIXED
# # # # # # # #         # --------------------------------------------------
# # # # # # # #         elif entry["type"] == "table":
# # # # # # # #             html_table = make_html_table(entry["cells"])

# # # # # # # #             if not current:
# # # # # # # #                 current_case_study += ("<br>" if current_case_study else "") + html_table
# # # # # # # #             else:
# # # # # # # #                 current["question"] += "<br>" + html_table

# # # # # # # #     if current:
# # # # # # # #         questions.append(current)

# # # # # # # #     return questions

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # SCORING ENGINE (UNCHANGED)
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def compute_score(questions, student_answers):
# # # # # # # #     score = 0
# # # # # # # #     total_marks = 0
# # # # # # # #     details = []

# # # # # # # #     for index, q in enumerate(questions, start=1):
# # # # # # # #         correct = q["answer"].strip().lower()
# # # # # # # #         total_marks += q["marks"]

# # # # # # # #         student_answer = ""
# # # # # # # #         for key in (index, f"q{index}", q["question"]):
# # # # # # # #             if key in student_answers:
# # # # # # # #                 student_answer = student_answers[key].strip().lower()
# # # # # # # #                 break

# # # # # # # #         got_it = student_answer == correct
# # # # # # # #         if got_it:
# # # # # # # #             score += q["marks"]

# # # # # # # #         details.append({
# # # # # # # #             "question": q["question"],
# # # # # # # #             "correct": correct,
# # # # # # # #             "student_answer": student_answer,
# # # # # # # #             "marks": q["marks"],
# # # # # # # #             "earned": q["marks"] if got_it else 0
# # # # # # # #         })

# # # # # # # #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# # # # # # # #     return {
# # # # # # # #         "score": score,
# # # # # # # #         "total": total_marks,
# # # # # # # #         "percentage": percentage,
# # # # # # # #         "details": details
# # # # # # # #     }

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # QUIZ STATUS (UNCHANGED)
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def get_quiz_status(questions, student_answers):
# # # # # # # #     status_list = []

# # # # # # # #     for index, q in enumerate(questions, start=1):
# # # # # # # #         correct = q["answer"].strip().lower()

# # # # # # # #         student_answer = ""
# # # # # # # #         for key in (index, f"q{index}", q["question"]):
# # # # # # # #             if key in student_answers:
# # # # # # # #                 student_answer = student_answers[key].strip().lower()
# # # # # # # #                 break

# # # # # # # #         if not student_answer:
# # # # # # # #             status = "unanswered"
# # # # # # # #         elif student_answer == correct:
# # # # # # # #             status = "correct"
# # # # # # # #         else:
# # # # # # # #             status = "incorrect"

# # # # # # # #         status_list.append({
# # # # # # # #             "question_index": index,
# # # # # # # #             "status": status,
# # # # # # # #             "student_answer": student_answer,
# # # # # # # #             "correct_answer": correct
# # # # # # # #         })

# # # # # # # #     return status_list

# # # # # # # # # ---------------------------------------------------------
# # # # # # # # # GOOGLE DRIVE HELPERS
# # # # # # # # # ---------------------------------------------------------
# # # # # # # # def extract_drive_id(url):
# # # # # # # #     patterns = [
# # # # # # # #         r"https://drive\\.google\\.com/file/d/([a-zA-Z0-9_-]+)",
# # # # # # # #         r"https://drive\\.google\\.com/open\\?id=([a-zA-Z0-9_-]+)"
# # # # # # # #     ]
# # # # # # # #     for pattern in patterns:
# # # # # # # #         m = re.search(pattern, url)
# # # # # # # #         if m:
# # # # # # # #             return m.group(1)
# # # # # # # #     return url

# # # # # # # # def get_drive_embed_url(drive_url_or_id):
# # # # # # # #     file_id = extract_drive_id(drive_url_or_id)
# # # # # # # #     return f"https://drive.google.com/file/d/{file_id}/preview"


# # # # # # # import docx
# # # # # # # import re
# # # # # # # import os
# # # # # # # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # # # # # # ---------------------------------------------------------
# # # # # # # # LOAD DOCX
# # # # # # # # ---------------------------------------------------------
# # # # # # # def load_docx(path):
# # # # # # #     return docx.Document(path)

# # # # # # # # ---------------------------------------------------------
# # # # # # # # HELPER MATCHERS
# # # # # # # # ---------------------------------------------------------
# # # # # # # def is_question_line(text):
# # # # # # #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # # # # # # def is_option_line(text):
# # # # # # #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # # # # # # def is_answer_line(text):
# # # # # # #     return text.lower().startswith(("answer", "ans", "correct"))

# # # # # # # # ---------------------------------------------------------
# # # # # # # # IMAGE EXTRACTION
# # # # # # # # ---------------------------------------------------------
# # # # # # # def extract_images(document, output_dir, q_index):
# # # # # # #     os.makedirs(output_dir, exist_ok=True)
# # # # # # #     images = []
# # # # # # #     count = 0
# # # # # # #     for rel in document.part.rels.values():
# # # # # # #         if rel.reltype == RT.IMAGE:
# # # # # # #             count += 1
# # # # # # #             ext = rel.target_ref.split('.')[-1]
# # # # # # #             filename = f"q{q_index}_img{count}.{ext}"
# # # # # # #             filepath = os.path.join(output_dir, filename)
# # # # # # #             with open(filepath, "wb") as f:
# # # # # # #                 f.write(rel.target_part.blob)
# # # # # # #             images.append(filename)
# # # # # # #     return images

# # # # # # # # ---------------------------------------------------------
# # # # # # # # HTML TABLE BUILDER
# # # # # # # # ---------------------------------------------------------
# # # # # # # def make_html_table(cells):
# # # # # # #     html = "<table class='table table-bordered'>"
# # # # # # #     for row in cells:
# # # # # # #         html += "<tr>"
# # # # # # #         for c in row:
# # # # # # #             html += f"<td>{c}</td>"
# # # # # # #         html += "</tr>"
# # # # # # #     html += "</table>"
# # # # # # #     return html

# # # # # # # # ---------------------------------------------------------
# # # # # # # # FLATTEN DOCX (PARAGRAPHS + TABLES)
# # # # # # # # ---------------------------------------------------------
# # # # # # # def flatten_doc(document):
# # # # # # #     lines = []

# # # # # # #     for block in document.element.body:
# # # # # # #         # Paragraph
# # # # # # #         if block.tag.endswith('p'):
# # # # # # #             para = docx.text.paragraph.Paragraph(block, document)
# # # # # # #             text = para.text.strip()
# # # # # # #             if text:
# # # # # # #                 lines.append({"type": "text", "content": text})

# # # # # # #         # Table
# # # # # # #         elif block.tag.endswith('tbl'):
# # # # # # #             table = docx.table.Table(block, document)
# # # # # # #             rows = []
# # # # # # #             for row in table.rows:
# # # # # # #                 cells = [c.text.strip() for c in row.cells]
# # # # # # #                 rows.append(cells)
# # # # # # #             lines.append({"type": "table", "cells": rows})

# # # # # # #     return lines

# # # # # # # # ---------------------------------------------------------
# # # # # # # # CASE STUDY CHECKER
# # # # # # # # ---------------------------------------------------------
# # # # # # # def is_case_study_line(text):
# # # # # # #     keywords = [
# # # # # # #         "use the following information",
# # # # # # #         "study the information",
# # # # # # #         "refer to the following",
# # # # # # #         "case study",
# # # # # # #         "use the data below"
# # # # # # #     ]
# # # # # # #     t = text.lower()
# # # # # # #     return any(k in t for k in keywords)

# # # # # # # # ---------------------------------------------------------
# # # # # # # # PARSE DOCX QUESTIONS – FIXED WITH INLINE CASE STUDY SPLIT
# # # # # # # # ---------------------------------------------------------
# # # # # # # def parse_docx_questions(path, image_output_dir=None):
# # # # # # #     doc = load_docx(path)
# # # # # # #     entries = flatten_doc(doc)

# # # # # # #     questions = []
# # # # # # #     current = None
# # # # # # #     q_index = 0
# # # # # # #     current_case_study = ""

# # # # # # #     for entry in entries:

# # # # # # #         # --------------------------------------------------
# # # # # # #         # TEXT ENTRY
# # # # # # #         # --------------------------------------------------
# # # # # # #         if entry["type"] == "text":
# # # # # # #             line = entry["content"]

# # # # # # #             # ---------- NEW QUESTION ----------
# # # # # # #             if is_question_line(line):
# # # # # # #                 if current:
# # # # # # #                     questions.append(current)

# # # # # # #                 q_index += 1

# # # # # # #                 # --- Remove number prefix ---
# # # # # # #                 raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# # # # # # #                 # --- Detect embedded case study inline ---
# # # # # # #                 embedded_case = ""
# # # # # # #                 for k in [
# # # # # # #                     "use the following information",
# # # # # # #                     "study the information",
# # # # # # #                     "refer to the following",
# # # # # # #                     "case study",
# # # # # # #                     "use the data below"
# # # # # # #                 ]:
# # # # # # #                     if k in raw_question.lower():
# # # # # # #                         parts = re.split(k, raw_question, flags=re.IGNORECASE)
# # # # # # #                         raw_question = parts[0].strip()
# # # # # # #                         embedded_case = k + " " + parts[1].strip()
# # # # # # #                         break

# # # # # # #                 # --- Extract marks ---
# # # # # # #                 mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
# # # # # # #                 marks = int(mk.group(1)) if mk else 1
# # # # # # #                 raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

# # # # # # #                 # Save question text
# # # # # # #                 question_text = raw_question

# # # # # # #                 # Create question entry
# # # # # # #                 current = {
# # # # # # #                     "question": question_text,
# # # # # # #                     "instructions": current_case_study.strip(),
# # # # # # #                     "a": "",
# # # # # # #                     "b": "",
# # # # # # #                     "c": "",
# # # # # # #                     "d": "",
# # # # # # #                     "answer": "",
# # # # # # #                     "marks": marks,
# # # # # # #                     "image": None
# # # # # # #                 }

# # # # # # #                 # Extract images
# # # # # # #                 if image_output_dir:
# # # # # # #                     imgs = extract_images(doc, image_output_dir, q_index)
# # # # # # #                     if imgs:
# # # # # # #                         current["image"] = imgs[0]

# # # # # # #                 # Reset case study for next question if embedded
# # # # # # #                 current_case_study = embedded_case

# # # # # # #             # ---------- OPTIONS ----------
# # # # # # #             elif current and is_option_line(line):
# # # # # # #                 letter = line[0].lower()
# # # # # # #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# # # # # # #                 current[letter] = text

# # # # # # #             # ---------- ANSWER ----------
# # # # # # #             elif current and is_answer_line(line):
# # # # # # #                 raw = line.split(":")[-1].strip().lower()
# # # # # # #                 clean = re.sub(r"[^a-d]", "", raw)
# # # # # # #                 current["answer"] = clean

# # # # # # #             # ---------- CASE STUDY OR EXTRA TEXT ----------
# # # # # # #             else:
# # # # # # #                 if not current:
# # # # # # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
# # # # # # #                 else:
# # # # # # #                     current["question"] += " " + line.strip()

# # # # # # #         # --------------------------------------------------
# # # # # # #         # TABLE ENTRY
# # # # # # #         # --------------------------------------------------
# # # # # # #         elif entry["type"] == "table":
# # # # # # #             html_table = make_html_table(entry["cells"])

# # # # # # #             if not current:
# # # # # # #                 current_case_study += ("<br>" if current_case_study else "") + html_table
# # # # # # #             else:
# # # # # # #                 current["question"] += "<br>" + html_table

# # # # # # #     if current:
# # # # # # #         questions.append(current)

# # # # # # #     return questions

# # # # # # # # ---------------------------------------------------------
# # # # # # # # SCORING ENGINE
# # # # # # # # ---------------------------------------------------------
# # # # # # # def compute_score(questions, student_answers):
# # # # # # #     score = 0
# # # # # # #     total_marks = 0
# # # # # # #     details = []

# # # # # # #     for index, q in enumerate(questions, start=1):
# # # # # # #         correct = q["answer"].strip().lower()
# # # # # # #         total_marks += q["marks"]

# # # # # # #         student_answer = ""
# # # # # # #         for key in (index, f"q{index}", q["question"]):
# # # # # # #             if key in student_answers:
# # # # # # #                 student_answer = student_answers[key].strip().lower()
# # # # # # #                 break

# # # # # # #         got_it = student_answer == correct
# # # # # # #         if got_it:
# # # # # # #             score += q["marks"]

# # # # # # #         details.append({
# # # # # # #             "question": q["question"],
# # # # # # #             "correct": correct,
# # # # # # #             "student_answer": student_answer,
# # # # # # #             "marks": q["marks"],
# # # # # # #             "earned": q["marks"] if got_it else 0
# # # # # # #         })

# # # # # # #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# # # # # # #     return {
# # # # # # #         "score": score,
# # # # # # #         "total": total_marks,
# # # # # # #         "percentage": percentage,
# # # # # # #         "details": details
# # # # # # #     }

# # # # # # # # ---------------------------------------------------------
# # # # # # # # QUIZ STATUS
# # # # # # # # ---------------------------------------------------------
# # # # # # # def get_quiz_status(questions, student_answers):
# # # # # # #     status_list = []

# # # # # # #     for index, q in enumerate(questions, start=1):
# # # # # # #         correct = q["answer"].strip().lower()

# # # # # # #         student_answer = ""
# # # # # # #         for key in (index, f"q{index}", q["question"]):
# # # # # # #             if key in student_answers:
# # # # # # #                 student_answer = student_answers[key].strip().lower()
# # # # # # #                 break

# # # # # # #         if not student_answer:
# # # # # # #             status = "unanswered"
# # # # # # #         elif student_answer == correct:
# # # # # # #             status = "correct"
# # # # # # #         else:
# # # # # # #             status = "incorrect"

# # # # # # #         status_list.append({
# # # # # # #             "question_index": index,
# # # # # # #             "status": status,
# # # # # # #             "student_answer": student_answer,
# # # # # # #             "correct_answer": correct
# # # # # # #         })

# # # # # # #     return status_list

# # # # # # # # ---------------------------------------------------------
# # # # # # # # GOOGLE DRIVE HELPERS
# # # # # # # # ---------------------------------------------------------
# # # # # # # def extract_drive_id(url):
# # # # # # #     patterns = [
# # # # # # #         r"https://drive\\.google\\.com/file/d/([a-zA-Z0-9_-]+)",
# # # # # # #         r"https://drive\\.google\\.com/open\\?id=([a-zA-Z0-9_-]+)"
# # # # # # #     ]
# # # # # # #     for pattern in patterns:
# # # # # # #         m = re.search(pattern, url)
# # # # # # #         if m:
# # # # # # #             return m.group(1)
# # # # # # #     return url

# # # # # # # def get_drive_embed_url(drive_url_or_id):
# # # # # # #     file_id = extract_drive_id(drive_url_or_id)
# # # # # # #     return f"https://drive.google.com/file/d/{file_id}/preview"


# # # # # # import docx
# # # # # # import re
# # # # # # import os
# # # # # # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # # # # # ---------------------------------------------------------
# # # # # # # LOAD DOCX
# # # # # # # ---------------------------------------------------------
# # # # # # def load_docx(path):
# # # # # #     return docx.Document(path)

# # # # # # # ---------------------------------------------------------
# # # # # # # HELPER MATCHERS
# # # # # # # ---------------------------------------------------------
# # # # # # def is_question_line(text):
# # # # # #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # # # # # def is_option_line(text):
# # # # # #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # # # # # def is_answer_line(text):
# # # # # #     return text.lower().startswith(("answer", "ans", "correct"))

# # # # # # # ---------------------------------------------------------
# # # # # # # IMAGE EXTRACTION
# # # # # # # ---------------------------------------------------------
# # # # # # def extract_images(document, output_dir, q_index):
# # # # # #     os.makedirs(output_dir, exist_ok=True)
# # # # # #     images = []
# # # # # #     count = 0
# # # # # #     for rel in document.part.rels.values():
# # # # # #         if rel.reltype == RT.IMAGE:
# # # # # #             count += 1
# # # # # #             ext = rel.target_ref.split('.')[-1]
# # # # # #             filename = f"q{q_index}_img{count}.{ext}"
# # # # # #             filepath = os.path.join(output_dir, filename)
# # # # # #             with open(filepath, "wb") as f:
# # # # # #                 f.write(rel.target_part.blob)
# # # # # #             images.append(filename)
# # # # # #     return images

# # # # # # # ---------------------------------------------------------
# # # # # # # HTML TABLE BUILDER
# # # # # # # ---------------------------------------------------------
# # # # # # def make_html_table(cells):
# # # # # #     html = "<table class='table table-bordered'>"
# # # # # #     for row in cells:
# # # # # #         html += "<tr>"
# # # # # #         for c in row:
# # # # # #             html += f"<td>{c}</td>"
# # # # # #         html += "</tr>"
# # # # # #     html += "</table>"
# # # # # #     return html

# # # # # # # ---------------------------------------------------------
# # # # # # # FLATTEN DOCX (PARAGRAPHS + TABLES)
# # # # # # # ---------------------------------------------------------
# # # # # # def flatten_doc(document):
# # # # # #     lines = []

# # # # # #     for block in document.element.body:
# # # # # #         # Paragraph
# # # # # #         if block.tag.endswith('p'):
# # # # # #             para = docx.text.paragraph.Paragraph(block, document)
# # # # # #             text = para.text.strip()
# # # # # #             if text:
# # # # # #                 lines.append({"type": "text", "content": text})

# # # # # #         # Table
# # # # # #         elif block.tag.endswith('tbl'):
# # # # # #             table = docx.table.Table(block, document)
# # # # # #             rows = []
# # # # # #             for row in table.rows:
# # # # # #                 cells = [c.text.strip() for c in row.cells]
# # # # # #                 rows.append(cells)
# # # # # #             lines.append({"type": "table", "cells": rows})

# # # # # #     return lines

# # # # # # # ---------------------------------------------------------
# # # # # # # CASE STUDY CHECKER
# # # # # # # ---------------------------------------------------------
# # # # # # def is_case_study_line(text):
# # # # # #     keywords = [
# # # # # #         "use the following information",
# # # # # #         "study the information",
# # # # # #         "refer to the following",
# # # # # #         "case study",
# # # # # #         "use the data below"
# # # # # #     ]
# # # # # #     t = text.lower()
# # # # # #     return any(k in t for k in keywords)

# # # # # # # ---------------------------------------------------------
# # # # # # # PARSE DOCX QUESTIONS – FULL FIXED
# # # # # # # ---------------------------------------------------------
# # # # # # def parse_docx_questions(path, image_output_dir=None):
# # # # # #     doc = load_docx(path)
# # # # # #     entries = flatten_doc(doc)

# # # # # #     questions = []
# # # # # #     current = None
# # # # # #     q_index = 0
# # # # # #     current_case_study = ""

# # # # # #     for entry in entries:

# # # # # #         # --------------------------------------------------
# # # # # #         # TEXT ENTRY
# # # # # #         # --------------------------------------------------
# # # # # #         if entry["type"] == "text":
# # # # # #             line = entry["content"]

# # # # # #             # ---------- NEW QUESTION ----------
# # # # # #             if is_question_line(line):
# # # # # #                 if current:
# # # # # #                     questions.append(current)

# # # # # #                 q_index += 1

# # # # # #                 # --- Remove number prefix ---
# # # # # #                 raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# # # # # #                 # --- Detect inline case study inside same line ---
# # # # # #                 embedded_case = ""
# # # # # #                 for k in [
# # # # # #                     "use the following information",
# # # # # #                     "study the information",
# # # # # #                     "refer to the following",
# # # # # #                     "case study",
# # # # # #                     "use the data below"
# # # # # #                 ]:
# # # # # #                     if k in raw_question.lower():
# # # # # #                         parts = re.split(k, raw_question, flags=re.IGNORECASE)
# # # # # #                         raw_question = parts[0].strip()
# # # # # #                         embedded_case = k + " " + parts[1].strip()
# # # # # #                         break

# # # # # #                 # --- Extract marks ---
# # # # # #                 mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
# # # # # #                 marks = int(mk.group(1)) if mk else 1
# # # # # #                 raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

# # # # # #                 # Save question text
# # # # # #                 question_text = raw_question

# # # # # #                 # Create question entry
# # # # # #                 current = {
# # # # # #                     "question": question_text,
# # # # # #                     "instructions": current_case_study.strip(),
# # # # # #                     "a": "",
# # # # # #                     "b": "",
# # # # # #                     "c": "",
# # # # # #                     "d": "",
# # # # # #                     "answer": "",
# # # # # #                     "marks": marks,
# # # # # #                     "image": None
# # # # # #                 }

# # # # # #                 # Extract images
# # # # # #                 if image_output_dir:
# # # # # #                     imgs = extract_images(doc, image_output_dir, q_index)
# # # # # #                     if imgs:
# # # # # #                         current["image"] = imgs[0]

# # # # # #                 # Reset case study for next question if embedded or already stored
# # # # # #                 current_case_study = embedded_case

# # # # # #             # ---------- OPTIONS ----------
# # # # # #             elif current and is_option_line(line):
# # # # # #                 letter = line[0].lower()
# # # # # #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# # # # # #                 current[letter] = text

# # # # # #             # ---------- ANSWER ----------
# # # # # #             elif current and is_answer_line(line):
# # # # # #                 raw = line.split(":")[-1].strip().lower()
# # # # # #                 clean = re.sub(r"[^a-d]", "", raw)
# # # # # #                 current["answer"] = clean

# # # # # #             # ---------- CASE STUDY OR EXTRA TEXT ----------
# # # # # #             else:
# # # # # #                 # If the line is a case study, attach it to the next question
# # # # # #                 if is_case_study_line(line):
# # # # # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
# # # # # #                 else:
# # # # # #                     # Otherwise, extra text belongs to current question if exists
# # # # # #                     if current:
# # # # # #                         current["question"] += " " + line.strip()
# # # # # #                     else:
# # # # # #                         # Text before any question → part of case study
# # # # # #                         current_case_study += ("<br>" if current_case_study else "") + line.strip()

# # # # # #         # --------------------------------------------------
# # # # # #         # TABLE ENTRY
# # # # # #         # --------------------------------------------------
# # # # # #         elif entry["type"] == "table":
# # # # # #             html_table = make_html_table(entry["cells"])

# # # # # #             if not current:
# # # # # #                 current_case_study += ("<br>" if current_case_study else "") + html_table
# # # # # #             else:
# # # # # #                 current["question"] += "<br>" + html_table

# # # # # #     if current:
# # # # # #         questions.append(current)

# # # # # #     return questions

# # # # # # # ---------------------------------------------------------
# # # # # # # SCORING ENGINE
# # # # # # # ---------------------------------------------------------
# # # # # # def compute_score(questions, student_answers):
# # # # # #     score = 0
# # # # # #     total_marks = 0
# # # # # #     details = []

# # # # # #     for index, q in enumerate(questions, start=1):
# # # # # #         correct = q["answer"].strip().lower()
# # # # # #         total_marks += q["marks"]

# # # # # #         student_answer = ""
# # # # # #         for key in (index, f"q{index}", q["question"]):
# # # # # #             if key in student_answers:
# # # # # #                 student_answer = student_answers[key].strip().lower()
# # # # # #                 break

# # # # # #         got_it = student_answer == correct
# # # # # #         if got_it:
# # # # # #             score += q["marks"]

# # # # # #         details.append({
# # # # # #             "question": q["question"],
# # # # # #             "correct": correct,
# # # # # #             "student_answer": student_answer,
# # # # # #             "marks": q["marks"],
# # # # # #             "earned": q["marks"] if got_it else 0
# # # # # #         })

# # # # # #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# # # # # #     return {
# # # # # #         "score": score,
# # # # # #         "total": total_marks,
# # # # # #         "percentage": percentage,
# # # # # #         "details": details
# # # # # #     }

# # # # # # # ---------------------------------------------------------
# # # # # # # QUIZ STATUS
# # # # # # # ---------------------------------------------------------
# # # # # # def get_quiz_status(questions, student_answers):
# # # # # #     status_list = []

# # # # # #     for index, q in enumerate(questions, start=1):
# # # # # #         correct = q["answer"].strip().lower()

# # # # # #         student_answer = ""
# # # # # #         for key in (index, f"q{index}", q["question"]):
# # # # # #             if key in student_answers:
# # # # # #                 student_answer = student_answers[key].strip().lower()
# # # # # #                 break

# # # # # #         if not student_answer:
# # # # # #             status = "unanswered"
# # # # # #         elif student_answer == correct:
# # # # # #             status = "correct"
# # # # # #         else:
# # # # # #             status = "incorrect"

# # # # # #         status_list.append({
# # # # # #             "question_index": index,
# # # # # #             "status": status,
# # # # # #             "student_answer": student_answer,
# # # # # #             "correct_answer": correct
# # # # # #         })

# # # # # #     return status_list

# # # # # # # ---------------------------------------------------------
# # # # # # # GOOGLE DRIVE HELPERS
# # # # # # # ---------------------------------------------------------
# # # # # # def extract_drive_id(url):
# # # # # #     patterns = [
# # # # # #         r"https://drive\\.google\\.com/file/d/([a-zA-Z0-9_-]+)",
# # # # # #         r"https://drive\\.google\\.com/open\\?id=([a-zA-Z0-9_-]+)"
# # # # # #     ]
# # # # # #     for pattern in patterns:
# # # # # #         m = re.search(pattern, url)
# # # # # #         if m:
# # # # # #             return m.group(1)
# # # # # #     return url

# # # # # # def get_drive_embed_url(drive_url_or_id):
# # # # # #     file_id = extract_drive_id(drive_url_or_id)
# # # # # #     return f"https://drive.google.com/file/d/{file_id}/preview"

# # # # # import docx
# # # # # import re
# # # # # import os
# # # # # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # # # # ---------------------------------------------------------
# # # # # # LOAD DOCX
# # # # # # ---------------------------------------------------------
# # # # # def load_docx(path):
# # # # #     return docx.Document(path)

# # # # # # ---------------------------------------------------------
# # # # # # HELPER MATCHERS
# # # # # # ---------------------------------------------------------
# # # # # def is_question_line(text):
# # # # #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # # # # def is_option_line(text):
# # # # #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # # # # def is_answer_line(text):
# # # # #     return text.lower().startswith(("answer", "ans", "correct"))

# # # # # # ---------------------------------------------------------
# # # # # # IMAGE EXTRACTION
# # # # # # ---------------------------------------------------------
# # # # # def extract_images(document, output_dir, q_index):
# # # # #     os.makedirs(output_dir, exist_ok=True)
# # # # #     images = []
# # # # #     count = 0
# # # # #     for rel in document.part.rels.values():
# # # # #         if rel.reltype == RT.IMAGE:
# # # # #             count += 1
# # # # #             ext = rel.target_ref.split('.')[-1]
# # # # #             filename = f"q{q_index}_img{count}.{ext}"
# # # # #             filepath = os.path.join(output_dir, filename)
# # # # #             with open(filepath, "wb") as f:
# # # # #                 f.write(rel.target_part.blob)
# # # # #             images.append(filename)
# # # # #     return images

# # # # # # ---------------------------------------------------------
# # # # # # HTML TABLE BUILDER
# # # # # # ---------------------------------------------------------
# # # # # def make_html_table(cells):
# # # # #     html = "<table class='table table-bordered'>"
# # # # #     for row in cells:
# # # # #         html += "<tr>"
# # # # #         for c in row:
# # # # #             html += f"<td>{c}</td>"
# # # # #         html += "</tr>"
# # # # #     html += "</table>"
# # # # #     return html

# # # # # # ---------------------------------------------------------
# # # # # # FLATTEN DOCX (PARAGRAPHS + TABLES)
# # # # # # ---------------------------------------------------------
# # # # # def flatten_doc(document):
# # # # #     lines = []

# # # # #     for block in document.element.body:
# # # # #         # Paragraph
# # # # #         if block.tag.endswith('p'):
# # # # #             para = docx.text.paragraph.Paragraph(block, document)
# # # # #             text = para.text.strip()
# # # # #             if text:
# # # # #                 lines.append({"type": "text", "content": text})

# # # # #         # Table
# # # # #         elif block.tag.endswith('tbl'):
# # # # #             table = docx.table.Table(block, document)
# # # # #             rows = []
# # # # #             for row in table.rows:
# # # # #                 cells = [c.text.strip() for c in row.cells]
# # # # #                 rows.append(cells)
# # # # #             lines.append({"type": "table", "cells": rows})

# # # # #     return lines

# # # # # # ---------------------------------------------------------
# # # # # # CASE STUDY CHECKER
# # # # # # ---------------------------------------------------------
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

# # # # # # ---------------------------------------------------------
# # # # # # PARSE DOCX QUESTIONS – FULL FIXED WITH POST-ANSWER CASE STUDY
# # # # # # ---------------------------------------------------------
# # # # # def parse_docx_questions(path, image_output_dir=None):
# # # # #     doc = load_docx(path)
# # # # #     entries = flatten_doc(doc)

# # # # #     questions = []
# # # # #     current = None
# # # # #     q_index = 0
# # # # #     current_case_study = ""

# # # # #     for entry in entries:

# # # # #         # --------------------------------------------------
# # # # #         # TEXT ENTRY
# # # # #         # --------------------------------------------------
# # # # #         if entry["type"] == "text":
# # # # #             line = entry["content"]

# # # # #             # ---------- NEW QUESTION ----------
# # # # #             if is_question_line(line):
# # # # #                 if current:
# # # # #                     questions.append(current)

# # # # #                 q_index += 1

# # # # #                 # --- Remove number prefix ---
# # # # #                 raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# # # # #                 # --- Detect inline case study inside same line ---
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

# # # # #                 # --- Extract marks ---
# # # # #                 mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
# # # # #                 marks = int(mk.group(1)) if mk else 1
# # # # #                 raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

# # # # #                 # Save question text
# # # # #                 question_text = raw_question

# # # # #                 # Create question entry
# # # # #                 current = {
# # # # #                     "question": question_text,
# # # # #                     "instructions": current_case_study.strip(),
# # # # #                     "a": "",
# # # # #                     "b": "",
# # # # #                     "c": "",
# # # # #                     "d": "",
# # # # #                    #"answer": "",
# # # # #                     "marks": marks,
# # # # #                     "image": None
# # # # #                 }

# # # # #                 # Extract images
# # # # #                 if image_output_dir:
# # # # #                     imgs = extract_images(doc, image_output_dir, q_index)
# # # # #                     if imgs:
# # # # #                         current["image"] = imgs[0]

# # # # #                 # Reset case study for next question if embedded or already stored
# # # # #                 current_case_study = embedded_case

# # # # #             # ---------- OPTIONS ----------
# # # # #             elif current and is_option_line(line):
# # # # #                 letter = line[0].lower()
# # # # #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# # # # #                 current[letter] = text

# # # # #             # ---------- ANSWER / POST-ANSWER CASE STUDY ----------
# # # # #             else:
# # # # #                 # Handle lines with both Answer and case study in same paragraph
# # # # #                 if current and "answer" in line.lower() and is_case_study_line(line):
# # # # #                     parts = re.split(
# # # # #                         r"(Use the following.*|Study the information.*|Refer to the following.*|Case study.*|Use the data below.*)",
# # # # #                         line,
# # # # #                         flags=re.IGNORECASE
# # # # #                     )
# # # # #                     answer_part = parts[0].strip()
# # # # #                     case_study_part = parts[1].strip() if len(parts) > 1 else ""

# # # # #                     # Process answer normally
# # # # #                     raw = answer_part.split(":")[-1].strip().lower()
# # # # #                     clean = re.sub(r"[^a-d]", "", raw)
# # # # #                     current["answer"] = clean

# # # # #                     # Store case study for next question
# # # # #                     if case_study_part:
# # # # #                         current_case_study += ("<br>" if current_case_study else "") + case_study_part

# # # # #                 # Regular case study line before any question
# # # # #                 elif is_case_study_line(line) and not current:
# # # # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()

# # # # #                 # Regular case study line after a question → attach to next
# # # # #                 elif is_case_study_line(line) and current:
# # # # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()

# # # # #                 # Otherwise, attach text to current question if exists
# # # # #                 else:
# # # # #                     if current:
# # # # #                         current["question"] += " " + line.strip()
# # # # #                     else:
# # # # #                         current_case_study += ("<br>" if current_case_study else "") + line.strip()

# # # # #         # --------------------------------------------------
# # # # #         # TABLE ENTRY
# # # # #         # --------------------------------------------------
# # # # #         elif entry["type"] == "table":
# # # # #             html_table = make_html_table(entry["cells"])

# # # # #             if not current:
# # # # #                 current_case_study += ("<br>" if current_case_study else "") + html_table
# # # # #             else:
# # # # #                 current["question"] += "<br>" + html_table

# # # # #     if current:
# # # # #         questions.append(current)

# # # # #     return questions

# # # # # # ---------------------------------------------------------
# # # # # # SCORING ENGINE
# # # # # # ---------------------------------------------------------
# # # # # def compute_score(questions, student_answers):
# # # # #     score = 0
# # # # #     total_marks = 0
# # # # #     details = []

# # # # #     for index, q in enumerate(questions, start=1):
# # # # #         correct = q["answer"].strip().lower()
# # # # #         total_marks += q["marks"]

# # # # #         student_answer = ""
# # # # #         for key in (index, f"q{index}", q["question"]):
# # # # #             if key in student_answers:
# # # # #                 student_answer = student_answers[key].strip().lower()
# # # # #                 break

# # # # #         got_it = student_answer == correct
# # # # #         if got_it:
# # # # #             score += q["marks"]

# # # # #         details.append({
# # # # #             "question": q["question"],
# # # # #             "correct": correct,
# # # # #             "student_answer": student_answer,
# # # # #             "marks": q["marks"],
# # # # #             "earned": q["marks"] if got_it else 0
# # # # #         })

# # # # #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# # # # #     return {
# # # # #         "score": score,
# # # # #         "total": total_marks,
# # # # #         "percentage": percentage,
# # # # #         "details": details
# # # # #     }

# # # # # # ---------------------------------------------------------
# # # # # # QUIZ STATUS
# # # # # # ---------------------------------------------------------
# # # # # def get_quiz_status(questions, student_answers):
# # # # #     status_list = []

# # # # #     for index, q in enumerate(questions, start=1):
# # # # #         correct = q["answer"].strip().lower()

# # # # #         student_answer = ""
# # # # #         for key in (index, f"q{index}", q["question"]):
# # # # #             if key in student_answers:
# # # # #                 student_answer = student_answers[key].strip().lower()
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

# # # # # # ---------------------------------------------------------
# # # # # # GOOGLE DRIVE HELPERS
# # # # # # ---------------------------------------------------------
# # # # # def extract_drive_id(url):
# # # # #     patterns = [
# # # # #         r"https://drive\\.google\\.com/file/d/([a-zA-Z0-9_-]+)",
# # # # #         r"https://drive\\.google\\.com/open\\?id=([a-zA-Z0-9_-]+)"
# # # # #     ]
# # # # #     for pattern in patterns:
# # # # #         m = re.search(pattern, url)
# # # # #         if m:
# # # # #             return m.group(1)
# # # # #     return url

# # # # # def get_drive_embed_url(drive_url_or_id):
# # # # #     file_id = extract_drive_id(drive_url_or_id)
# # # # #     return f"https://drive.google.com/file/d/{file_id}/preview"


# # # # import docx
# # # # import re
# # # # import os
# # # # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # # # ---------------------------------------------------------
# # # # # LOAD DOCX
# # # # # ---------------------------------------------------------
# # # # def load_docx(path):
# # # #     return docx.Document(path)

# # # # # ---------------------------------------------------------
# # # # # HELPER MATCHERS
# # # # # ---------------------------------------------------------
# # # # def is_question_line(text):
# # # #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # # # def is_option_line(text):
# # # #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # # # def is_answer_line(text):
# # # #     return text.lower().startswith(("answer", "ans", "correct"))

# # # # # ---------------------------------------------------------
# # # # # IMAGE EXTRACTION
# # # # # ---------------------------------------------------------
# # # # def extract_images(document, output_dir, q_index):
# # # #     os.makedirs(output_dir, exist_ok=True)
# # # #     images = []
# # # #     count = 0
# # # #     for rel in document.part.rels.values():
# # # #         if rel.reltype == RT.IMAGE:
# # # #             count += 1
# # # #             ext = rel.target_ref.split('.')[-1]
# # # #             filename = f"q{q_index}_img{count}.{ext}"
# # # #             filepath = os.path.join(output_dir, filename)
# # # #             with open(filepath, "wb") as f:
# # # #                 f.write(rel.target_part.blob)
# # # #             images.append(filename)
# # # #     return images

# # # # # ---------------------------------------------------------
# # # # # HTML TABLE BUILDER
# # # # # ---------------------------------------------------------
# # # # def make_html_table(cells):
# # # #     html = "<table class='table table-bordered'>"
# # # #     for row in cells:
# # # #         html += "<tr>"
# # # #         for c in row:
# # # #             html += f"<td>{c}</td>"
# # # #         html += "</tr>"
# # # #     html += "</table>"
# # # #     return html

# # # # # ---------------------------------------------------------
# # # # # FLATTEN DOCX (PARAGRAPHS + TABLES)
# # # # # ---------------------------------------------------------
# # # # def flatten_doc(document):
# # # #     lines = []
# # # #     for block in document.element.body:
# # # #         # Paragraph
# # # #         if block.tag.endswith('p'):
# # # #             para = docx.text.paragraph.Paragraph(block, document)
# # # #             text = para.text.strip()
# # # #             if text:
# # # #                 lines.append({"type": "text", "content": text})
# # # #         # Table
# # # #         elif block.tag.endswith('tbl'):
# # # #             table = docx.table.Table(block, document)
# # # #             rows = []
# # # #             for row in table.rows:
# # # #                 cells = [c.text.strip() for c in row.cells]
# # # #                 rows.append(cells)
# # # #             lines.append({"type": "table", "cells": rows})
# # # #     return lines

# # # # # ---------------------------------------------------------
# # # # # CASE STUDY CHECKER
# # # # # ---------------------------------------------------------
# # # # def is_case_study_line(text):
# # # #     keywords = [
# # # #         "use the following information",
# # # #         "study the information",
# # # #         "refer to the following",
# # # #         "case study",
# # # #         "use the data below"
# # # #     ]
# # # #     t = text.lower()
# # # #     return any(k in t for k in keywords)

# # # # # ---------------------------------------------------------
# # # # # PARSE DOCX QUESTIONS – ANSWERS KEPT INTERNALLY
# # # # # ---------------------------------------------------------
# # # # def parse_docx_questions(path, image_output_dir=None):
# # # #     doc = load_docx(path)
# # # #     entries = flatten_doc(doc)

# # # #     questions = []
# # # #     current = None
# # # #     q_index = 0
# # # #     current_case_study = ""

# # # #     for entry in entries:
# # # #         if entry["type"] == "text":
# # # #             line = entry["content"]

# # # #             # ---------- NEW QUESTION ----------
# # # #             if is_question_line(line):
# # # #                 if current:
# # # #                     questions.append(current)
# # # #                 q_index += 1

# # # #                 raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# # # #                 # Detect inline case study inside same line
# # # #                 embedded_case = ""
# # # #                 for k in [
# # # #                     "use the following information",
# # # #                     "study the information",
# # # #                     "refer to the following",
# # # #                     "case study",
# # # #                     "use the data below"
# # # #                 ]:
# # # #                     if k in raw_question.lower():
# # # #                         parts = re.split(k, raw_question, flags=re.IGNORECASE)
# # # #                         raw_question = parts[0].strip()
# # # #                         embedded_case = k + " " + parts[1].strip()
# # # #                         break

# # # #                 # Extract marks
# # # #                 mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
# # # #                 marks = int(mk.group(1)) if mk else 1
# # # #                 raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

# # # #                 # Create question object with internal answer placeholder
# # # #                 current = {
# # # #                     "question": raw_question,
# # # #                     "instructions": current_case_study.strip(),
# # # #                     "a": "",
# # # #                     "b": "",
# # # #                     "c": "",
# # # #                     "d": "",
# # # #                     "answer_internal": "",  # internal only
# # # #                     "marks": marks,
# # # #                     "image": None
# # # #                 }

# # # #                 if image_output_dir:
# # # #                     imgs = extract_images(doc, image_output_dir, q_index)
# # # #                     if imgs:
# # # #                         current["image"] = imgs[0]

# # # #                 current_case_study = embedded_case

# # # #             # ---------- OPTIONS ----------
# # # #             elif current and is_option_line(line):
# # # #                 letter = line[0].lower()
# # # #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# # # #                 current[letter] = text

# # # #             # ---------- ANSWER / POST-ANSWER CASE STUDY ----------
# # # #             else:
# # # #                 if current and is_answer_line(line):
# # # #                     raw = line.split(":")[-1].strip().lower()
# # # #                     clean = re.sub(r"[^a-d]", "", raw)
# # # #                     current["answer_internal"] = clean
# # # #                 elif is_case_study_line(line) and not current:
# # # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
# # # #                 elif is_case_study_line(line) and current:
# # # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
# # # #                 else:
# # # #                     if current:
# # # #                         current["question"] += " " + line.strip()
# # # #                     else:
# # # #                         current_case_study += ("<br>" if current_case_study else "") + line.strip()

# # # #         # ---------- TABLE ENTRY ----------
# # # #         elif entry["type"] == "table":
# # # #             html_table = make_html_table(entry["cells"])
# # # #             if not current:
# # # #                 current_case_study += ("<br>" if current_case_study else "") + html_table
# # # #             else:
# # # #                 current["question"] += "<br>" + html_table

# # # #     if current:
# # # #         questions.append(current)

# # # #     return questions

# # # # # ---------------------------------------------------------
# # # # # PREPARE STUDENT-FACING QUESTIONS (ANSWERS HIDDEN)
# # # # # ---------------------------------------------------------
# # # # def prepare_questions_for_student(questions):
# # # #     student_questions = []
# # # #     for q in questions:
# # # #         student_questions.append({
# # # #             "question": q["question"],
# # # #             "instructions": q["instructions"],
# # # #             "a": q["a"],
# # # #             "b": q["b"],
# # # #             "c": q["c"],
# # # #             "d": q["d"],
# # # #             "marks": q["marks"],
# # # #             "image": q["image"]
# # # #         })
# # # #     return student_questions

# # # # # ---------------------------------------------------------
# # # # # SCORING ENGINE
# # # # # ---------------------------------------------------------
# # # # def compute_score(questions, student_answers):
# # # #     score = 0
# # # #     total_marks = 0
# # # #     details = []

# # # #     for index, q in enumerate(questions, start=1):
# # # #         correct = q.get("answer_internal", "").strip().lower()
# # # #         total_marks += q["marks"]

# # # #         student_answer = ""
# # # #         for key in (index, f"q{index}", q["question"]):
# # # #             if key in student_answers:
# # # #                 student_answer = student_answers[key].strip().lower()
# # # #                 break

# # # #         got_it = student_answer == correct
# # # #         if got_it:
# # # #             score += q["marks"]

# # # #         details.append({
# # # #             "question": q["question"],
# # # #             "correct": correct,
# # # #             "student_answer": student_answer,
# # # #             "marks": q["marks"],
# # # #             "earned": q["marks"] if got_it else 0
# # # #         })

# # # #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# # # #     return {
# # # #         "score": score,
# # # #         "total": total_marks,
# # # #         "percentage": percentage,
# # # #         "details": details
# # # #     }

# # # # # ---------------------------------------------------------
# # # # # QUIZ STATUS
# # # # # ---------------------------------------------------------
# # # # def get_quiz_status(questions, student_answers):
# # # #     status_list = []

# # # #     for index, q in enumerate(questions, start=1):
# # # #         correct = q.get("answer_internal", "").strip().lower()

# # # #         student_answer = ""
# # # #         for key in (index, f"q{index}", q["question"]):
# # # #             if key in student_answers:
# # # #                 student_answer = student_answers[key].strip().lower()
# # # #                 break

# # # #         if not student_answer:
# # # #             status = "unanswered"
# # # #         elif student_answer == correct:
# # # #             status = "correct"
# # # #         else:
# # # #             status = "incorrect"

# # # #         status_list.append({
# # # #             "question_index": index,
# # # #             "status": status,
# # # #             "student_answer": student_answer,
# # # #             "correct_answer": correct
# # # #         })

# # # #     return status_list

# # # # # ---------------------------------------------------------
# # # # # GOOGLE DRIVE HELPERS
# # # # # ---------------------------------------------------------
# # # # def extract_drive_id(url):
# # # #     patterns = [
# # # #         r"https://drive\\.google\\.com/file/d/([a-zA-Z0-9_-]+)",
# # # #         r"https://drive\\.google\\.com/open\\?id=([a-zA-Z0-9_-]+)"
# # # #     ]
# # # #     for pattern in patterns:
# # # #         m = re.search(pattern, url)
# # # #         if m:
# # # #             return m.group(1)
# # # #     return url

# # # # def get_drive_embed_url(drive_url_or_id):
# # # #     file_id = extract_drive_id(drive_url_or_id)
# # # #     return f"https://drive.google.com/file/d/{file_id}/preview"

# # # import docx
# # # import re
# # # import os
# # # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # # ---------------------------------------------------------
# # # # LOAD DOCX
# # # # ---------------------------------------------------------
# # # def load_docx(path):
# # #     return docx.Document(path)

# # # # ---------------------------------------------------------
# # # # HELPER MATCHERS
# # # # ---------------------------------------------------------
# # # def is_question_line(text):
# # #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # # def is_option_line(text):
# # #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # # def is_answer_line(text):
# # #     return text.lower().startswith(("answer", "ans", "correct"))

# # # def is_case_study_line(text):
# # #     keywords = [
# # #         "use the following information",
# # #         "study the information",
# # #         "refer to the following",
# # #         "case study",
# # #         "use the data below"
# # #     ]
# # #     t = text.lower()
# # #     return any(k in t for k in keywords)

# # # # ---------------------------------------------------------
# # # # IMAGE EXTRACTION
# # # # ---------------------------------------------------------
# # # def extract_images(document, output_dir, q_index):
# # #     os.makedirs(output_dir, exist_ok=True)
# # #     images = []
# # #     count = 0
# # #     for rel in document.part.rels.values():
# # #         if rel.reltype == RT.IMAGE:
# # #             count += 1
# # #             ext = rel.target_ref.split('.')[-1]
# # #             filename = f"q{q_index}_img{count}.{ext}"
# # #             filepath = os.path.join(output_dir, filename)
# # #             with open(filepath, "wb") as f:
# # #                 f.write(rel.target_part.blob)
# # #             images.append(filename)
# # #     return images

# # # # ---------------------------------------------------------
# # # # HTML TABLE BUILDER
# # # # ---------------------------------------------------------
# # # def make_html_table(cells):
# # #     html = "<table class='table table-bordered'>"
# # #     for row in cells:
# # #         html += "<tr>"
# # #         for c in row:
# # #             html += f"<td>{c}</td>"
# # #         html += "</tr>"
# # #     html += "</table>"
# # #     return html

# # # # ---------------------------------------------------------
# # # # FLATTEN DOCX (PARAGRAPHS + TABLES)
# # # # ---------------------------------------------------------
# # # def flatten_doc(document):
# # #     lines = []
# # #     for block in document.element.body:
# # #         # Paragraph
# # #         if block.tag.endswith('p'):
# # #             para = docx.text.paragraph.Paragraph(block, document)
# # #             text = para.text.strip()
# # #             if text:
# # #                 lines.append({"type": "text", "content": text})
# # #         # Table
# # #         elif block.tag.endswith('tbl'):
# # #             table = docx.table.Table(block, document)
# # #             rows = []
# # #             for row in table.rows:
# # #                 cells = [c.text.strip() for c in row.cells]
# # #                 rows.append(cells)
# # #             lines.append({"type": "table", "cells": rows})
# # #     return lines

# # # # ---------------------------------------------------------
# # # # PARSE DOCX QUESTIONS WITH CASE STUDY AND TABLES
# # # # ---------------------------------------------------------
# # # def parse_docx_questions(path, image_output_dir=None):
# # #     doc = load_docx(path)
# # #     entries = flatten_doc(doc)

# # #     questions = []
# # #     current = None
# # #     q_index = 0
# # #     current_case_study = ""

# # #     for entry in entries:
# # #         if entry["type"] == "text":
# # #             line = entry["content"]

# # #             # ---------- NEW QUESTION ----------
# # #             if is_question_line(line):
# # #                 if current:
# # #                     questions.append(current)
# # #                 q_index += 1

# # #                 raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# # #                 # Detect inline case study inside same line
# # #                 embedded_case = ""
# # #                 for k in [
# # #                     "use the following information",
# # #                     "study the information",
# # #                     "refer to the following",
# # #                     "case study",
# # #                     "use the data below"
# # #                 ]:
# # #                     if k in raw_question.lower():
# # #                         parts = re.split(k, raw_question, flags=re.IGNORECASE)
# # #                         raw_question = parts[0].strip()
# # #                         embedded_case = k + " " + parts[1].strip()
# # #                         break

# # #                 # Extract marks
# # #                 mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
# # #                 marks = int(mk.group(1)) if mk else 1
# # #                 raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

# # #                 # Create question object
# # #                 current = {
# # #                     "question": raw_question,
# # #                     "instructions": current_case_study.strip(),
# # #                     "a": "",
# # #                     "b": "",
# # #                     "c": "",
# # #                     "d": "",
# # #                     "answer_internal": "",  # for scoring
# # #                     "marks": marks,
# # #                     "image": None
# # #                 }

# # #                 if image_output_dir:
# # #                     imgs = extract_images(doc, image_output_dir, q_index)
# # #                     if imgs:
# # #                         current["image"] = imgs[0]

# # #                 current_case_study = embedded_case

# # #             # ---------- OPTIONS ----------
# # #             elif current and is_option_line(line):
# # #                 letter = line[0].lower()
# # #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# # #                 current[letter] = text

# # #             # ---------- ANSWER / CASE STUDY ----------
# # #             else:
# # #                 if current and is_answer_line(line):
# # #                     raw = line.split(":")[-1].strip().lower()
# # #                     clean = re.sub(r"[^a-d]", "", raw)
# # #                     current["answer_internal"] = clean
# # #                 elif is_case_study_line(line) and not current:
# # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
# # #                 elif is_case_study_line(line) and current:
# # #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
# # #                 else:
# # #                     if current:
# # #                         current["question"] += " " + line.strip()
# # #                     else:
# # #                         current_case_study += ("<br>" if current_case_study else "") + line.strip()

# # #         # ---------- TABLE ENTRY ----------
# # #         elif entry["type"] == "table":
# # #             html_table = make_html_table(entry["cells"])
# # #             if not current:
# # #                 current_case_study += ("<br>" if current_case_study else "") + html_table
# # #             else:
# # #                 current["question"] += "<br>" + html_table

# # #     if current:
# # #         questions.append(current)

# # #     return questions

# # # # ---------------------------------------------------------
# # # # PREPARE STUDENT-FACING QUESTIONS (ANSWERS HIDDEN)
# # # # ---------------------------------------------------------
# # # def prepare_questions_for_student(questions):
# # #     student_questions = []
# # #     for q in questions:
# # #         student_questions.append({
# # #             "question": q["question"],
# # #             "instructions": q["instructions"],
# # #             "a": q["a"],
# # #             "b": q["b"],
# # #             "c": q["c"],
# # #             "d": q["d"],
# # #             "marks": q["marks"],
# # #             "image": q["image"]
# # #         })
# # #     return student_questions

# # # # ---------------------------------------------------------
# # # # SCORING ENGINE (WORKS NOW)
# # # # ---------------------------------------------------------
# # # def compute_score(questions, student_answers):
# # #     score = 0
# # #     total_marks = 0
# # #     details = []

# # #     for index, q in enumerate(questions, start=1):
# # #         correct = q.get("answer_internal", "").strip().lower()
# # #         total_marks += q["marks"]

# # #         # Student answer lookup only by q index or q1/q2...
# # #         student_answer = ""
# # #         for key in (index, f"q{index}"):
# # #             if key in student_answers:
# # #                 student_answer = student_answers[key].strip().lower()
# # #                 break

# # #         got_it = student_answer == correct
# # #         if got_it:
# # #             score += q["marks"]

# # #         details.append({
# # #             "question": q["question"],
# # #             "correct": correct,
# # #             "student_answer": student_answer,
# # #             "marks": q["marks"],
# # #             "earned": q["marks"] if got_it else 0
# # #         })

# # #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# # #     return {
# # #         "score": score,
# # #         "total": total_marks,
# # #         "percentage": percentage,
# # #         "details": details
# # #     }

# # # # ---------------------------------------------------------
# # # # QUIZ STATUS
# # # # ---------------------------------------------------------
# # # def get_quiz_status(questions, student_answers):
# # #     status_list = []

# # #     for index, q in enumerate(questions, start=1):
# # #         correct = q.get("answer_internal", "").strip().lower()
# # #         student_answer = ""
# # #         for key in (index, f"q{index}"):
# # #             if key in student_answers:
# # #                 student_answer = student_answers[key].strip().lower()
# # #                 break

# # #         if not student_answer:
# # #             status = "unanswered"
# # #         elif student_answer == correct:
# # #             status = "correct"
# # #         else:
# # #             status = "incorrect"

# # #         status_list.append({
# # #             "question_index": index,
# # #             "status": status,
# # #             "student_answer": student_answer,
# # #             "correct_answer": correct
# # #         })

# # #     return status_list

# # import docx
# # import re
# # import os
# # from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # # ------------------- LOAD DOCX -------------------
# # def load_docx(path):
# #     return docx.Document(path)

# # # ------------------- HELPERS -------------------
# # def is_question_line(text):
# #     return bool(re.match(r"^\d+[\.\)]\s*", text))

# # def is_option_line(text):
# #     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# # def is_answer_line(text):
# #     return text.lower().startswith(("answer", "ans", "correct"))

# # def is_case_study_line(text):
# #     keywords = [
# #         "use the following information",
# #         "study the information",
# #         "refer to the following",
# #         "case study",
# #         "use the data below"
# #     ]
# #     t = text.lower()
# #     return any(k in t for k in keywords)

# # # ------------------- IMAGE EXTRACTION -------------------
# # def extract_images(document, output_dir, q_index):
# #     os.makedirs(output_dir, exist_ok=True)
# #     images = []
# #     count = 0
# #     for rel in document.part.rels.values():
# #         if rel.reltype == RT.IMAGE:
# #             count += 1
# #             ext = rel.target_ref.split('.')[-1]
# #             filename = f"q{q_index}_img{count}.{ext}"
# #             filepath = os.path.join(output_dir, filename)
# #             with open(filepath, "wb") as f:
# #                 f.write(rel.target_part.blob)
# #             images.append(filename)
# #     return images

# # # ------------------- HTML TABLE BUILDER -------------------
# # def make_html_table(cells):
# #     html = "<table class='table table-bordered'>"
# #     for row in cells:
# #         html += "<tr>"
# #         for c in row:
# #             html += f"<td>{c}</td>"
# #         html += "</tr>"
# #     html += "</table>"
# #     return html

# # # ------------------- FLATTEN DOCX -------------------
# # def flatten_doc(document):
# #     lines = []
# #     for block in document.element.body:
# #         if block.tag.endswith('p'):
# #             para = docx.text.paragraph.Paragraph(block, document)
# #             text = para.text.strip()
# #             if text:
# #                 lines.append({"type": "text", "content": text})
# #         elif block.tag.endswith('tbl'):
# #             table = docx.table.Table(block, document)
# #             rows = []
# #             for row in table.rows:
# #                 cells = [c.text.strip() for c in row.cells]
# #                 rows.append(cells)
# #             lines.append({"type": "table", "cells": rows})
# #     return lines

# # # ------------------- PARSE QUESTIONS (TABLES + CASE STUDY) -------------------
# # def parse_docx_questions(path, image_output_dir=None):
# #     doc = load_docx(path)
# #     entries = flatten_doc(doc)

# #     questions = []
# #     current = None
# #     q_index = 0
# #     current_case_study = ""

# #     for entry in entries:
# #         if entry["type"] == "text":
# #             line = entry["content"]

# #             if is_question_line(line):
# #                 if current:
# #                     questions.append(current)
# #                 q_index += 1

# #                 raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

# #                 # Detect inline case study inside the same line
# #                 embedded_case = ""
# #                 for k in [
# #                     "use the following information",
# #                     "study the information",
# #                     "refer to the following",
# #                     "case study",
# #                     "use the data below"
# #                 ]:
# #                     if k in raw_question.lower():
# #                         parts = re.split(k, raw_question, flags=re.IGNORECASE)
# #                         raw_question = parts[0].strip()
# #                         embedded_case = k + " " + parts[1].strip()
# #                         break

# #                 mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
# #                 marks = int(mk.group(1)) if mk else 1
# #                 raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

# #                 current = {
# #                     "question": raw_question,
# #                     "instructions": current_case_study.strip(),
# #                     "a": "",
# #                     "b": "",
# #                     "c": "",
# #                     "d": "",
# #                     "answer_internal": "",  # for scoring
# #                     "marks": marks,
# #                     "image": None
# #                 }

# #                 if image_output_dir:
# #                     imgs = extract_images(doc, image_output_dir, q_index)
# #                     if imgs:
# #                         current["image"] = imgs[0]

# #                 current_case_study = embedded_case

# #             elif current and is_option_line(line):
# #                 letter = line[0].lower()
# #                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
# #                 current[letter] = text

# #             else:
# #                 if current and is_answer_line(line):
# #                     raw = line.split(":")[-1].strip().lower()
# #                     clean = re.sub(r"[^a-d]", "", raw)
# #                     current["answer_internal"] = clean
# #                 elif is_case_study_line(line):
# #                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
# #                 else:
# #                     if current:
# #                         current["question"] += " " + line.strip()
# #                     else:
# #                         current_case_study += ("<br>" if current_case_study else "") + line.strip()

# #         elif entry["type"] == "table":
# #             html_table = make_html_table(entry["cells"])
# #             if not current:
# #                 current_case_study += ("<br>" if current_case_study else "") + html_table
# #             else:
# #                 current["question"] += "<br>" + html_table

# #     if current:
# #         questions.append(current)

# #     return questions

# # # ------------------- PREPARE QUESTIONS FOR STUDENT -------------------
# # def prepare_questions_for_student(questions):
# #     student_questions = []
# #     for q in questions:
# #         student_questions.append({
# #             "question": q["question"],
# #             "instructions": q["instructions"],
# #             "a": q["a"],
# #             "b": q["b"],
# #             "c": q["c"],
# #             "d": q["d"],
# #             "marks": q["marks"],
# #             "image": q["image"]
# #         })
# #     return student_questions

# # # ------------------- SCORING ENGINE -------------------
# # def compute_score(questions, student_answers):
# #     score = 0
# #     total_marks = 0
# #     details = []

# #     for index, q in enumerate(questions, start=1):
# #         correct = q.get("answer_internal", "").strip().lower()
# #         total_marks += q["marks"]

# #         student_answer = ""
# #         for key in (index, f"q{index}"):
# #             if key in student_answers:
# #                 student_answer = student_answers[key].strip().lower()
# #                 break

# #         got_it = student_answer == correct
# #         if got_it:
# #             score += q["marks"]

# #         details.append({
# #             "question": q["question"],
# #             "correct": correct,
# #             "student_answer": student_answer,
# #             "marks": q["marks"],
# #             "earned": q["marks"] if got_it else 0
# #         })

# #     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

# #     return {
# #         "score": score,
# #         "total": total_marks,
# #         "percentage": percentage,
# #         "details": details
# #     }

# # # ------------------- QUIZ STATUS -------------------
# # def get_quiz_status(questions, student_answers):
# #     status_list = []

# #     for index, q in enumerate(questions, start=1):
# #         correct = q.get("answer_internal", "").strip().lower()
# #         student_answer = ""
# #         for key in (index, f"q{index}"):
# #             if key in student_answers:
# #                 student_answer = student_answers[key].strip().lower()
# #                 break

# #         if not student_answer:
# #             status = "unanswered"
# #         elif student_answer == correct:
# #             status = "correct"
# #         else:
# #             status = "incorrect"

# #         status_list.append({
# #             "question_index": index,
# #             "status": status,
# #             "student_answer": student_answer,
# #             "correct_answer": correct
# #         })

# #     return status_list

# import docx
# import re
# import os
# from docx.opc.constants import RELATIONSHIP_TYPE as RT

# # ---------------------------------------------------------
# # LOAD DOCX
# # ---------------------------------------------------------
# def load_docx(path):
#     return docx.Document(path)

# # ---------------------------------------------------------
# # HELPER MATCHERS
# # ---------------------------------------------------------
# def is_question_line(text):
#     return bool(re.match(r"^\d+[\.\)]\s*", text))

# def is_option_line(text):
#     return bool(re.match(r"^[A-D][\.\):]\s*", text.strip(), re.IGNORECASE))

# def is_answer_line(text):
#     return text.lower().startswith(("answer", "ans", "correct"))

# def is_case_study_line(text):
#     keywords = [
#         "use the following information",
#         "study the information",
#         "refer to the following",
#         "case study",
#         "use the data below"
#     ]
#     return any(k in text.lower() for k in keywords)

# # ---------------------------------------------------------
# # IMAGE EXTRACTION
# # ---------------------------------------------------------
# def extract_images(document, output_dir, q_index):
#     os.makedirs(output_dir, exist_ok=True)
#     images = []
#     count = 0
#     for rel in document.part.rels.values():
#         if rel.reltype == RT.IMAGE:
#             count += 1
#             ext = rel.target_ref.split('.')[-1]
#             filename = f"q{q_index}_img{count}.{ext}"
#             filepath = os.path.join(output_dir, filename)
#             with open(filepath, "wb") as f:
#                 f.write(rel.target_part.blob)
#             images.append(filename)
#     return images

# # ---------------------------------------------------------
# # HTML TABLE BUILDER
# # ---------------------------------------------------------
# def make_html_table(cells):
#     html = "<table class='table table-bordered'>"
#     for row in cells:
#         html += "<tr>" + "".join(f"<td>{c}</td>" for c in row) + "</tr>"
#     html += "</table>"
#     return html

# # ---------------------------------------------------------
# # FLATTEN DOCX (PARAGRAPHS + TABLES)
# # ---------------------------------------------------------
# def flatten_doc(document):
#     lines = []
#     for block in document.element.body:
#         # Paragraph
#         if block.tag.endswith('p'):
#             para = docx.text.paragraph.Paragraph(block, document)
#             text = para.text.strip()
#             if text:
#                 lines.append({"type": "text", "content": text})
#         # Table
#         elif block.tag.endswith('tbl'):
#             table = docx.table.Table(block, document)
#             rows = []
#             for row in table.rows:
#                 cells = [c.text.strip() for c in row.cells]
#                 rows.append(cells)
#             lines.append({"type": "table", "cells": rows})
#     return lines

# # ---------------------------------------------------------
# # PARSE DOCX QUESTIONS (WITH CASE STUDY, TABLES, IMAGES)
# # ---------------------------------------------------------
# def parse_docx_questions(path, image_output_dir=None):
#     doc = load_docx(path)
#     entries = flatten_doc(doc)

#     questions = []
#     current = None
#     q_index = 0
#     current_case_study = ""

#     for entry in entries:
#         if entry["type"] == "text":
#             line = entry["content"]

#             # ---------- NEW QUESTION ----------
#             if is_question_line(line):
#                 if current:
#                     questions.append(current)
#                 q_index += 1
#                 raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

#                 # Extract embedded case study from same line
#                 embedded_case = ""
#                 for k in [
#                     "use the following information",
#                     "study the information",
#                     "refer to the following",
#                     "case study",
#                     "use the data below"
#                 ]:
#                     if k in raw_question.lower():
#                         parts = re.split(k, raw_question, flags=re.IGNORECASE)
#                         raw_question = parts[0].strip()
#                         embedded_case = k + " " + parts[1].strip()
#                         break

#                 # Extract marks
#                 mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
#                 marks = int(mk.group(1)) if mk else 1
#                 raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

#                 current = {
#                     "question": raw_question,
#                     "instructions": current_case_study.strip(),
#                     "a": "",
#                     "b": "",
#                     "c": "",
#                     "d": "",
#                     "answer_internal": "",  # for scoring
#                     "marks": marks,
#                     "image": None
#                 }

#                 if image_output_dir:
#                     imgs = extract_images(doc, image_output_dir, q_index)
#                     if imgs:
#                         current["image"] = imgs[0]

#                 current_case_study = embedded_case

#             # ---------- OPTIONS ----------
#             elif current and is_option_line(line):
#                 letter = line[0].lower()
#                 text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
#                 current[letter] = text

#             # ---------- ANSWER / CASE STUDY ----------
#             else:
#                 if current and is_answer_line(line):
#                     raw = line.split(":")[-1].strip().lower()
#                     clean = re.sub(r"[^a-d]", "", raw)
#                     current["answer_internal"] = clean
#                 elif is_case_study_line(line) and not current:
#                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
#                 elif is_case_study_line(line) and current:
#                     current_case_study += ("<br>" if current_case_study else "") + line.strip()
#                 else:
#                     if current:
#                         current["question"] += " " + line.strip()
#                     else:
#                         current_case_study += ("<br>" if current_case_study else "") + line.strip()

#         # ---------- TABLE ENTRY ----------
#         elif entry["type"] == "table":
#             html_table = make_html_table(entry["cells"])
#             if not current:
#                 current_case_study += ("<br>" if current_case_study else "") + html_table
#             else:
#                 current["question"] += "<br>" + html_table

#     if current:
#         questions.append(current)

#     return questions

# # ---------------------------------------------------------
# # PREPARE STUDENT-FACING QUESTIONS (ANSWERS HIDDEN)
# # ---------------------------------------------------------
# def prepare_questions_for_student(questions):
#     student_questions = []
#     for q in questions:
#         student_questions.append({
#             "question": q["question"],
#             "instructions": q["instructions"],  # case study / instructions at top
#             "a": q["a"],
#             "b": q["b"],
#             "c": q["c"],
#             "d": q["d"],
#             "marks": q["marks"],
#             "image": q["image"]
#         })
#     return student_questions

# # ---------------------------------------------------------
# # SCORING ENGINE
# # ---------------------------------------------------------
# def compute_score(questions, student_answers):
#     score = 0
#     total_marks = 0
#     details = []

#     for index, q in enumerate(questions, start=1):
#         correct = q.get("answer_internal", "").strip().lower()
#         total_marks += q["marks"]

#         student_answer = ""
#         for key in (index, f"q{index}"):
#             if key in student_answers:
#                 val = student_answers[key].strip().lower()
#                 # Normalize to a,b,c,d
#                 student_answer = re.sub(r"[^a-d]", "", val)
#                 break

#         got_it = student_answer == correct

#         if got_it:
#             score += q["marks"]

#         details.append({
#             "question": q["question"],
#             "correct": correct,
#             "student_answer": student_answer,
#             "marks": q["marks"],
#             "earned": q["marks"] if got_it else 0
#         })

#     percentage = round((score / total_marks) * 100, 2) if total_marks else 0

#     return {
#         "score": score,
#         "total": total_marks,
#         "percentage": percentage,
#         "details": details
#     }


# # ---------------------------------------------------------
# # QUIZ STATUS
# # ---------------------------------------------------------
# def get_quiz_status(questions, student_answers):
#     status_list = []
#     for index, q in enumerate(questions, start=1):
#         correct = q.get("answer_internal", "").strip().lower()
#         student_answer = ""
#         for key in (index, f"q{index}"):
#             if key in student_answers:
#                 student_answer = student_answers[key].strip().lower()
#                 break

#         if not student_answer:
#             status = "unanswered"
#         elif student_answer == correct:
#             status = "correct"
#         else:
#             status = "incorrect"

#         status_list.append({
#             "question_index": index,
#             "status": status,
#             "student_answer": student_answer,
#             "correct_answer": correct
#         })
#     return status_list

# # ---------------------------------------------------------
# # GOOGLE DRIVE HELPERS (now included safely)
# # ---------------------------------------------------------
# def extract_drive_id(url):
#     patterns = [
#         r"https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)",
#         r"https://drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)"
#     ]
#     for pattern in patterns:
#         m = re.search(pattern, url)
#         if m:
#             return m.group(1)
#     return url

# def get_drive_embed_url(drive_url_or_id):
#     file_id = extract_drive_id(drive_url_or_id)
#     return f"https://drive.google.com/file/d/{file_id}/preview"

import docx
import re
import os
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# --------------------- LOAD DOCX ---------------------
def load_docx(path):
    return docx.Document(path)

# ------------------- IMAGE EXTRACTION PER PARAGRAPH -------------------
def extract_images_from_paragraph(paragraph, output_dir, q_index):
    """
    Extract images from a single paragraph only.
    Returns a list of filenames extracted from that paragraph.
    """
    os.makedirs(output_dir, exist_ok=True)
    images = []
    count = 0
    for run in paragraph.runs:
        # XPath to find blip (image) elements
        drawing_elements = run.element.xpath('.//a:blip', 
                                            namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
        for blip in drawing_elements:
            rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if rId:
                image_part = paragraph.part.related_parts[rId]
                count += 1
                ext = image_part.partname.split('.')[-1]
                filename = f"q{q_index}_img{count}.{ext}"
                filepath = os.path.join(output_dir, filename)
                with open(filepath, "wb") as f:
                    f.write(image_part.blob)
                images.append(filename)
    return images

# ------------------- FLATTEN DOCX --------------------
def flatten_doc(document):
    """
    Returns a list of entries preserving paragraphs and tables in document order.
    Each entry is {"type":"text","content":para} or {"type":"table","cells": [...]}
    """
    lines = []
    for block in document.element.body:
        if block.tag.endswith('p'):
            para = docx.text.paragraph.Paragraph(block, document)
            text = para.text.strip()
            if text or para.runs:
                lines.append({"type": "text", "content": text, "para": para})
        elif block.tag.endswith('tbl'):
            table = docx.table.Table(block, document)
            rows = []
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                rows.append(cells)
            lines.append({"type": "table", "cells": rows})
    return lines

# ------------------- PARSE DOCX QUESTIONS --------------------
def parse_docx_questions(path, image_output_dir=None):
    """
    Parse docx file and return list of questions.
    Each question dict has:
      - question (text with any inline tables as HTML)
      - instructions (case study/instructions appearing BEFORE the question)
      - a, b, c, d (option texts)
      - answer (single letter 'a'..'d', stored lower-case)
      - marks (int)
      - image (filename or None)
    """
    doc = load_docx(path)
    entries = flatten_doc(doc)

    questions = []
    current = None
    q_index = 0
    pending_instructions = ""

    for entry in entries:
        if entry["type"] == "text":
            line = entry["content"].strip()
            para = entry.get("para")

            # ---------- NEW QUESTION ----------
            if re.match(r"^\d+[\.\)]\s*", line):
                # push previous question
                if current:
                    questions.append(current)

                q_index += 1
                raw_question = re.sub(r"^\d+[\.\)]\s*", "", line).strip()

                # extract marks like (2 mks)
                mk = re.search(r"\((\d+)\s*mks?\)", raw_question, re.IGNORECASE)
                marks = int(mk.group(1)) if mk else 1
                raw_question = re.sub(r"\(\d+\s*mks?\)", "", raw_question).strip()

                current = {
                    "question": raw_question,
                    "instructions": pending_instructions.strip(),
                    "a": "",
                    "b": "",
                    "c": "",
                    "d": "",
                    "answer": "",
                    "marks": marks,
                    "image": None
                }

                # Extract images only from this paragraph
                if image_output_dir and para:
                    imgs = extract_images_from_paragraph(para, image_output_dir, q_index)
                    if imgs:
                        current["image"] = imgs[0]

                pending_instructions = ""  # reset pending for next question

            # ---------- OPTION (A-D) ----------
            elif current and re.match(r"^[A-D][\.\):]\s*", line.strip(), re.IGNORECASE):
                letter = line[0].lower()
                text = re.sub(r"^[A-D][\.\):]\s*", "", line).strip()
                current[letter] = text

            # ---------- ANSWER ----------
            elif current and line.lower().lstrip().startswith(("answer", "ans", "correct")):
                raw = line.split(":")[-1].strip().lower()
                clean = re.sub(r"[^a-d]", "", raw)
                current["answer"] = clean

            # ---------- CASE STUDY / INSTRUCTIONS ----------
            elif any(k in line.lower() for k in [
                "use the following information", "study the information", 
                "refer to the following", "case study", "use the data below"
            ]):
                if current:
                    pending_instructions += ("<br>" if pending_instructions else "") + line
                else:
                    pending_instructions += ("<br>" if pending_instructions else "") + line

            else:
                if current:
                    current["question"] += " " + line
                else:
                    pending_instructions += ("<br>" if pending_instructions else "") + line

        # ---------- TABLE ----------
        elif entry["type"] == "table":
            html_table = "<table class='table table-bordered'>"
            for row in entry["cells"]:
                html_table += "<tr>" + "".join([f"<td>{c}</td>" for c in row]) + "</tr>"
            html_table += "</table>"

            if current:
                current["question"] += "<br>" + html_table
            else:
                pending_instructions += ("<br>" if pending_instructions else "") + html_table

    # push last question
    if current:
        questions.append(current)

    return questions
