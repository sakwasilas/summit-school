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