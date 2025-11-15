def parse_docx_questions(file_stream, image_output_dir=DEFAULT_IMAGE_DIR):
    """Parse case study, questions, images, tables from a .docx file."""

    document = Document(file_stream)

    os.makedirs(image_output_dir, exist_ok=True)

    case_study_html = ""      # ⬅ store ALL content before question 1
    questions = []
    current_question = None
    extra_html_parts = []
    image_counter = 0
    skipped = 0

    for block in iter_block_items(document):

        # --------------------------------------------------
        # PARAGRAPH HANDLING
        # --------------------------------------------------
        if isinstance(block, Paragraph):
            para = block
            text = para.text.strip()

            # Extract images inside paragraph
            for run in para.runs:
                image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
                if image_name:
                    image_counter += 1

                    # Image BEFORE any question → belongs to case study
                    if current_question is None:
                        case_study_html += f"<img src='/static/question_images/{image_name}' /><br>"
                    else:
                        current_question["image"] = image_name

            # Skip empty lines
            if not text:
                continue

            # --------------------------------------------------
            # BEFORE FIRST QUESTION → CASE STUDY
            # --------------------------------------------------
            if current_question is None and not re.match(r"^\d+[\.\)]", text):
                case_study_html += f"<p>{text}</p>"
                continue

            # --------------------------------------------------
            # QUESTION DETECTED (e.g., “1.” or “1)”)
            # --------------------------------------------------
            if re.match(r"^\d+[\.\)]", text):

                # Close previous question
                if current_question:
                    if extra_html_parts:
                        current_question["extra_content"] += ''.join(extra_html_parts)
                        extra_html_parts = []

                    if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
                        questions.append(current_question)
                    else:
                        skipped += 1

                # Extract marks
                marks_match = re.search(r"\((\d+)\s?(?:mks|marks?)\)", text, re.IGNORECASE)
                marks = int(marks_match.group(1)) if marks_match else 1

                # Clean question text
                clean_text = re.sub(r"\s*\(\d+\s?(?:mks|marks?)\)", "", text)
                question_text = re.sub(r"^\d+[\.\)]\s*", "", clean_text)

                # Create new question object
                current_question = {
                    "question": question_text,
                    "a": "", "b": "", "c": "", "d": "",
                    "answer": "",
                    "extra_content": "",
                    "image": None,
                    "marks": marks
                }

                continue

            # --------------------------------------------------
            # OPTIONS (A, B, C, D)
            # --------------------------------------------------
            if re.match(r"^\(?[a-dA-D][\.\)]", text):
                match = re.match(r"^\(?([a-dA-D])[\.\)]\s*(.+)", text)
                if match and current_question:
                    label = match.group(1).lower()
                    current_question[label] = match.group(2).strip()
                continue

            # --------------------------------------------------
            # ANSWER LINE
            # --------------------------------------------------
            if re.match(r"^(answer|correct answer):", text, re.IGNORECASE):
                match = re.search(r":\s*([a-dA-D])", text, re.IGNORECASE)
                if match and current_question:
                    current_question["answer"] = match.group(1).lower()
                continue

            # --------------------------------------------------
            # EXTRA CONTENT INSIDE QUESTION
            # --------------------------------------------------
            if current_question:
                extra_html_parts.append(f"<p>{text}</p>")
                continue

        # --------------------------------------------------
        # TABLE HANDLING
        # --------------------------------------------------
        elif isinstance(block, Table):
            table_html = extract_table_html(block)

            # Case study table
            if current_question is None:
                case_study_html += table_html
            else:
                current_question["extra_content"] += table_html

            continue

    # --------------------------------------------------
    # SAVE LAST QUESTION
    # --------------------------------------------------
    if current_question:
        if extra_html_parts:
            current_question["extra_content"] += ''.join(extra_html_parts)

        if current_question.get("question") and current_question.get("answer") in ["a", "b", "c", "d"]:
            questions.append(current_question)
        else:
            skipped += 1

    # Return structured result
    return {
        "case_study": case_study_html,
        "questions": questions,
        "skipped": skipped
    }
