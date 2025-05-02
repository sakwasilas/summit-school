def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
    """Parse case study + MCQ questions without exposing answers to the user."""

    document = Document(file_stream)
    os.makedirs(image_output_dir, exist_ok=True)

    case_study_html = ""
    questions = []
    current_question = None
    extra_html_parts = []
    image_counter = 0
    skipped = 0

    for block in iter_block_items(document):

        # -----------------------------
        # PARAGRAPHS
        # -----------------------------
        if isinstance(block, Paragraph):
            para = block
            text = para.text.strip()

            # Extract inline images
            for run in para.runs:
                image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
                if image_name:
                    image_counter += 1
                    if current_question is None:
                        case_study_html += f"<img src='/static/question_images/{image_name}' /><br>"
                    else:
                        current_question["image"] = image_name

            if not text:
                continue

            # -------------------------------------
            # BEFORE FIRST QUESTION — CASE STUDY
            # -------------------------------------
            if current_question is None and not re.match(r"^\d+[\.\)]", text):
                case_study_html += f"<p>{text}</p>"
                continue

            # -------------------------------------
            # NEW QUESTION DETECTED (1. , 1) , 1 )
            # -------------------------------------
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

                # Remove number + marks
                clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
                question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)

                current_question = {
                    "question": question_text,
                    "a": "", "b": "", "c": "", "d": "",
                    "correct_answer": "",   # stored internally but HIDDEN
                    "extra_content": "",
                    "image": None,
                    "marks": marks
                }

                continue

            # -------------------------------------
            # OPTIONS A–D
            # -------------------------------------
            if re.match(r"^\(?[a-dA-D][\.\)]", text):
                match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
                if match and current_question:
                    label = match.group(1).lower()
                    current_question[label] = match.group(2).strip()
                continue

            # -------------------------------------
            # Correct Answer line — stored but NOT returned
            # -------------------------------------
            if re.match(r"^(answer|correct answer):", text, re.IGNORECASE):
                match = re.search(r":\s*([a-dA-D])", text, re.IGNORECASE)
                if match and current_question:
                    current_question["correct_answer"] = match.group(1).lower()
                continue

            # -------------------------------------
            # Extra info inside question
            # -------------------------------------
            if current_question:
                extra_html_parts.append(f"<p>{text}</p>")
                continue

        # -----------------------------
        # TABLES
        # -----------------------------
        elif isinstance(block, Table):
            table_html = extract_table_html(block)
            if current_question is None:
                case_study_html += table_html
            else:
                current_question["extra_content"] += table_html
            continue

    # -----------------------------
    # SAVE LAST QUESTION
    # -----------------------------
    if current_question:
        if extra_html_parts:
            current_question["extra_content"] += ''.join(extra_html_parts)
        questions.append(current_question)

    # -----------------------------
    # REMOVE ANSWERS FROM USER OUTPUT
    # -----------------------------
    safe_questions = []
    for q in questions:
        safe_q = {k: v for k, v in q.items() if k != "correct_answer"}
        safe_questions.append(safe_q)

    return safe_questions
