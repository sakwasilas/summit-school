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

# ------------------- IMAGE EXTRACTION -----------------
def extract_images(document, output_dir, q_index):
    os.makedirs(output_dir, exist_ok=True)
    images = []
    count = 0
    # iterate over relationships to find images
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

                # extract image if requested (note: this extracts doc-level images; names may collide
                # across questions if multiple images exist - kept for backward compatibility)
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
