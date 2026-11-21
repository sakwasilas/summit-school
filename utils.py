import os
import re
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

DEFAULT_IMAGE_DIR = "static/question_images"


def normalize_text(text):
    """Clean up text from Word documents - preserving numbers and currency"""
    if text is None:
        return ""
    # Remove non-breaking spaces
    text = text.replace("\xa0", " ").replace("\u2022", "-")
    # Remove multiple spaces
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def extract_table_text(table):
    """Extract text from a Word table and convert to HTML table"""
    html = "<table class='data-table'>"
    for row in table.rows:
        html += "<tr>"
        for cell in row.cells:
            # Get text from paragraphs in cell
            cell_text = "<br>".join([normalize_text(p.text) for p in cell.paragraphs if p.text.strip()])
            html += f"<td>{cell_text}</td>"
        html += "</tr>"
    html += "</table>"
    return html


def save_image_from_run(run, output_dir, image_counter):
    """Save an image from a Word run to disk"""
    try:
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

        # Determine image extension from content type
        ext = 'png'
        if hasattr(image_part, 'content_type'):
            if 'jpeg' in image_part.content_type or 'jpg' in image_part.content_type:
                ext = 'jpg'
            elif 'gif' in image_part.content_type:
                ext = 'gif'
            elif 'png' in image_part.content_type:
                ext = 'png'

        image_filename = f"question_image_{image_counter}.{ext}"
        image_path = os.path.join(output_dir, image_filename)

        with open(image_path, 'wb') as f:
            f.write(image_data)

        return image_filename
    except Exception as e:
        print(f"⚠️ Error saving image: {e}")
        return None


def iter_block_items(document):
    """Iterate through paragraphs and tables in a document"""
    body = document.element.body
    for child in body.iterchildren():
        if child.tag.endswith('p'):
            yield Paragraph(child, document)
        elif child.tag.endswith('tbl'):
            yield Table(child, document)


def append_with_newline(existing, new_text):
    """Append text with proper newline formatting"""
    new_text = normalize_text(new_text)
    if not new_text:
        return existing
    if not existing:
        return new_text
    return existing + "\n" + new_text


def is_question_start(text):
    """Check if text starts a new question - handles numbered questions"""
    text = text.strip()
    return bool(re.match(r"^(\d+)[\.\)]\s*", text))


def has_word_numbering(paragraph):
    """Detect if a paragraph was given an automatic number by MS Word."""
    try:
        pPr = paragraph._element.pPr
        if pPr is None:
            return False
        numPr = pPr.numPr
        if numPr is None:
            return False
        return numPr.numId is not None
    except Exception:
        return False


def is_option_line(text):
    """Check if text is an option (A, B, C, D)"""
    text = text.strip()
    return bool(re.match(r"^([A-Da-d])[\.\)]\s*", text))


def is_answer_line(text):
    """Check if text is an answer line"""
    text = text.strip().lower()
    return bool(re.match(r"^(answer|correct answer|ans)\s*[:.]?\s*([A-Da-d])", text))


def extract_answer(text):
    """Extract the answer letter from text"""
    text = text.strip().lower()
    match = re.search(r"(answer|correct answer|ans)\s*[:.]?\s*([A-Da-d])", text)
    if match:
        return match.group(2).lower()
    return ""


def get_marks_from_text(text):
    """Extract marks from question text"""
    text = text.strip()
    match = re.search(r"\((\d+)\s*(?:mks|marks?)?\)", text, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return 1


def is_financial_table(table_html):
    """Check if table contains financial data"""
    table_text = table_html.lower()
    financial_terms = [
        'shs', 'debit', 'credit', 'balance', 'account', 'amount',
        'revenue', 'expense', 'asset', 'liability', 'equity',
        'income', 'profit', 'loss', 'cash', 'bank', 'capital',
        'opening', 'closing', 'total', 'subtotal', 'cost', 'selling',
        'sales', 'purchases', 'inventory', 'stock', 'receivable',
        'payable', 'loan', 'interest', 'tax', 'dividend', 'depreciation'
    ]
    count = sum(1 for term in financial_terms if term in table_text)
    return count >= 3


def is_financial_data(text):
    """Check if text contains financial data"""
    text = text.lower().strip()
    financial_patterns = [
        r'(?:shs|kshs|kes)\s*[\d,]+',
        r'[\$\€\£\₦\₹]',
        r'\d{1,3}(?:,\d{3})*(?:\.\d{2})?',
        r'debit|credit|dr|cr',
        r'balance sheet|income statement|trial balance',
        r'profit|loss|revenue|expense|asset|liability|equity',
        r'cost|selling|inventory|stock|receivable|payable|loan',
        r'interest|tax|dividend|depreciation'
    ]
    return any(re.search(pattern, text, re.IGNORECASE) for pattern in financial_patterns)


def is_instruction_line(text):
    """Check if text contains exam instructions"""
    text = text.lower().strip()
    instruction_keywords = [
        "use the following", "use matrices", "use the information",
        "refer to the following", "answer question", "answer questions",
        "read the following", "study the following", "consider the following",
        "based on the following", "using the following", "given the following",
        "the following table", "the following information", "the following data",
        "the following scenario", "the following passage", "the following text",
        "the following case", "the following diagram", "the following figure",
        "the following chart", "the following graph", "the following trial balance",
        "the following balance sheet", "the following income statement",
        "the following financial statements", "the following accounts",
        "the following transactions", "the following ledger", "the following journal",
        "the following inventory", "the following vehicles", "the following assets",
        "the following liabilities", "the following statements",
        "statement of profit or loss", "statement of financial position",
        "use the following information to answer", "prepare the following",
        "calculate the following", "compute the following", "determine the following",
        "analyze the following", "the following additional information",
        "the following further information"
    ]
    return any(keyword in text for keyword in instruction_keywords)


def parse_docx_questions(file_path, image_output_dir=DEFAULT_IMAGE_DIR):
    """
    Parse questions from a DOCX file with enhanced support for accounting questions
    """
    try:
        document = Document(file_path)
    except Exception as e:
        print(f"❌ Error opening document: {e}")
        return []

    questions = []
    current_question = None
    current_option = None
    shared_instructions = []
    image_counter = 0
    skipped = 0
    in_instruction_block = False
    instruction_counter = 0

    os.makedirs(image_output_dir, exist_ok=True)

    blocks = list(iter_block_items(document))
    print(f"📄 Found {len(blocks)} blocks to process")

    def save_current_question():
        nonlocal current_question, skipped
        if not current_question:
            return
        
        question_text = normalize_text(current_question.get("question", ""))
        if question_text:
            questions.append(current_question)
            print(f"✅ Saved question {len(questions)}")
        else:
            skipped += 1

    def clean_shared_instructions():
        nonlocal shared_instructions
        if shared_instructions:
            result = "\n\n".join(shared_instructions)
            shared_instructions = []
            return result
        return ""

    for idx, block in enumerate(blocks):
        if isinstance(block, Paragraph):
            text = normalize_text(block.text)
            if not text:
                continue

            print(f"\n📝 Block {idx}: {text[:100]}...")

            for run in block.runs:
                if current_question and not current_question.get("image"):
                    image_name = save_image_from_run(run, image_output_dir, image_counter + 1)
                    if image_name:
                        image_counter += 1
                        current_question["image"] = image_name
                        print(f"📸 Saved image: {image_name}")

            if not current_question and is_instruction_line(text):
                shared_instructions.append(text)
                in_instruction_block = True
                instruction_counter += 1
                print(f"📘 Added to shared instructions ({instruction_counter}): {text[:50]}...")
                continue

            is_auto_numbered = has_word_numbering(block) and not is_option_line(text)
            starts_new_question = is_question_start(text) or is_auto_numbered

            if starts_new_question:
                save_current_question()

                display_text = text
                if is_auto_numbered and not is_question_start(text):
                    display_text = f"{len(questions) + 1}. {text}"
                    print(f"🔢 Auto-numbered by Word → injecting '{len(questions) + 1}.'")

                instructions = clean_shared_instructions()
                
                full_text = display_text
                if instructions:
                    full_text = instructions + "\n\n" + display_text
                    in_instruction_block = False

                current_question = {
                    "question": full_text,
                    "a": "",
                    "b": "",
                    "c": "",
                    "d": "",
                    "answer": "",
                    "image": None,
                    "marks": get_marks_from_text(display_text),
                    "extra_content": ""
                }
                current_option = None
                print(f"🆕 New question started: {display_text[:50]}...")
                continue

            option_match = re.match(r"^([A-Da-d])[\.\)]\s*(.*)", text)
            if option_match and current_question:
                label = option_match.group(1).lower()
                option_text = option_match.group(2).strip()
                current_question[label] = option_text
                current_option = label
                print(f"🔠 Option {label.upper()}: {option_text[:50]}...")
                continue

            if is_answer_line(text) and current_question:
                current_question["answer"] = extract_answer(text)
                print(f"✅ Answer: {current_question['answer'] or 'Not found'}")
                continue

            if current_question:
                if current_option:
                    current_question[current_option] = append_with_newline(
                        current_question[current_option], text
                    )
                    print(f"↪ Added to option {current_option.upper()}")
                else:
                    if is_financial_data(text):
                        text = "📊 " + text
                    
                    current_question["question"] = append_with_newline(
                        current_question["question"], text
                    )
                    print(f"↪ Added to question")
            else:
                shared_instructions.append(text)
                print(f"↪ Added to shared instructions")

        elif isinstance(block, Table):
            table_html = extract_table_text(block)
            if not table_html:
                continue

            print(f"📊 Table found")

            # FIX: If the previous question already has an answer, we assume the table
            # is a case study for upcoming questions and detach it from the previous question.
            if current_question and current_question.get("answer"):
                save_current_question()
                current_question = None
                current_option = None

            # If we are currently in the middle of choosing an option, reset it
            # so the table goes to the main question text, not Option D.
            if current_option:
                current_option = None

            if current_question:
                current_question["question"] = append_with_newline(
                    current_question["question"], table_html
                )
                print(f"↪ Table added to question")
            else:
                shared_instructions.append(table_html)
                print(f"↪ Table added to shared instructions")

    save_current_question()

    print(f"\n✅ Parsed {len(questions)} questions successfully")
    if skipped:
        print(f"⚠️ Skipped {skipped} invalid questions")

    for i, q in enumerate(questions[:3], 1):
        print(f"\n📋 Question {i}:")
        print(f"   Text: {q['question'][:100]}...")
        if q.get('a'):
            print(f"   A: {q['a'][:50]}")
        if q.get('b'):
            print(f"   B: {q['b'][:50]}")
        if q.get('c'):
            print(f"   C: {q['c'][:50]}")
        if q.get('d'):
            print(f"   D: {q['d'][:50]}")
        print(f"   Answer: {q['answer']}")
        if q.get('image'):
            print(f"   Image: {q['image']}")

    return questions


def get_quiz_status(user_id):
    return "active"


def extract_drive_id(url):
    patterns = [
        r"https://drive\.google\.com/file/d/([A-Za-z0-9_-]+)",
        r"https://drive\.google\.com/open\?id=([A-Za-z0-9_-]+)",
        r"https://drive\.google\.com/uc\?id=([A-Za-z0-9_-]+)",
    ]

    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)

    return url


def get_drive_embed_url(drive_url_or_id):
    file_id = extract_drive_id(drive_url_or_id)
    return f"https://drive.google.com/file/d/{file_id}/preview"