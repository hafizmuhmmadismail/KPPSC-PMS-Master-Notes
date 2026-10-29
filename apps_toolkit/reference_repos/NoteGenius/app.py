from flask import Flask, render_template, request, redirect, url_for, session
from utils.flashcards import generate_flashcards_from_keywords
from utils.mindmap import generate_mindmap
from utils.mnemonics import create_mnemonic
from utils.pptgen import generate_ppt
from utils.tts import text_to_speech
from utils.qna import get_answers
from utils.quiz_generator import generate_quiz
from utils.suggestions import suggest_links
from utils.summary import summarize_text
from utils.mcq_generator import MCQGenerator
from utils.ques_gen import process_text_by_sentence
import pdfplumber 
import re 
import os
import pytesseract
from PIL import Image
import io




app = Flask(__name__, instance_relative_config=True)
app.secret_key = '95ecb3644cd380a0551436b7eed83c00'
app.config['UPLOAD_FOLDER'] = 'instance'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        choice = request.form.get('input_choice')
        if choice == 'text':
            return redirect(url_for('text_input'))
        elif choice == 'pdf':
            return redirect(url_for('upload_pdf'))
    return render_template('input_choice.html')

# ---- Step 2a: Enter Text ----
@app.route('/text_input', methods=['GET', 'POST'])
def text_input():
    if request.method == 'POST':
        input_text = request.form.get('input_text', '').strip()
        if input_text:
            print("📥 Text received from user:\n", input_text[:500])
            session['input_text'] = input_text
            print("✅ Session set for text input.")
            return redirect(url_for('operations'))
        else:
            print("⚠️ No text entered.")
    return render_template('text_input.html')

# ---- Step 2b: Upload PDF ----
import pdfplumber
import pytesseract
from PIL import Image
from docx import Document
from fpdf import FPDF
import os

def extract_text_from_pdf(pdf_path):
    extracted_text = ""

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"
                else:
                    print(f"⚠️ Page {page_num} has no text, using OCR...")
                    pil_image = page.to_image(resolution=300).original
                    ocr_text = pytesseract.image_to_string(pil_image)
                    extracted_text += ocr_text + "\n"

        cleaned_text = clean_text(extracted_text)
        if not cleaned_text.strip():
            print("⚠️ Extracted text is empty after cleaning.")
            return None

        print("✅ PDF text extracted successfully.")
        return cleaned_text

    except Exception as e:
        print(f"❌ Error during PDF extraction: {e}")
        return None

def convert_to_pdf(input_path, output_folder):
    """
    Converts .doc/.docx/.jpg/.jpeg/.png to .pdf
    Returns the path to the generated PDF.
    """
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    pdf_path = os.path.join(output_folder, f"{base_name}.pdf")

    try:
        if input_path.lower().endswith(('.doc', '.docx')):
            doc = Document(input_path)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            for para in doc.paragraphs:
                text = para.text.encode('latin-1', 'replace').decode('latin-1')
                pdf.multi_cell(0, 10, text)
            pdf.output(pdf_path)

        elif input_path.lower().endswith(('.jpg', '.jpeg', '.png')):
            image = Image.open(input_path)
            rgb_image = image.convert('RGB')
            rgb_image.save(pdf_path)

        else:
            raise ValueError("Unsupported file type")

        print(f"✅ Converted to PDF: {pdf_path}")
        return pdf_path

    except Exception as e:
        print(f"❌ Error converting to PDF: {e}")
        return None

@app.route('/upload_file', methods=['GET', 'POST'])
def upload_file():
    message = None

    if request.method == 'POST':
        uploaded_file = request.files.get('uploaded_file')
        if not uploaded_file:
            return render_template('upload_file.html', message="⚠ Please upload a file.")

        filename = uploaded_file.filename.lower()
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename)
        uploaded_file.save(file_path)

        # --- Step 1: If it's a PDF, extract directly ---
        if filename.endswith('.pdf'):
            cleaned_text = extract_text_from_pdf(file_path)

        # --- Step 2: If it's a DOC/DOCX/JPG, convert then extract ---
        elif filename.endswith(('.doc', '.docx', '.jpg', '.jpeg', '.png')):
            pdf_path = convert_to_pdf(file_path, app.config['UPLOAD_FOLDER'])
            if not pdf_path:
                return render_template('upload_file.html', message="❌ Conversion to PDF failed.")
            cleaned_text = extract_text_from_pdf(pdf_path)

        else:
            return render_template('upload_file.html', message="❌ Unsupported file type.")

        # --- Step 3: Handle extracted text ---
        if cleaned_text:
            session['input_text'] = cleaned_text
            print("✅ File processed successfully.")
            print("📄 Preview (first 500 chars):\n", cleaned_text[:500])
            return redirect(url_for('operations'))
        else:
            return render_template('upload_file.html', message="❌ Text extraction failed.")

    return render_template('upload_file.html', message=message)


# ---- Step 3: Operations ----
@app.route('/operations', methods=['GET', 'POST'])
def operations():
    if request.method == 'POST':
        # Safely get form values
        input_text = request.form.get('input_text')
        operation = request.form.get('operations')  # now matches the <select name="operations">

        if input_text:
            session['input_text'] = input_text  # store the text in session

        if not operation:
            message = "Please select an operation."
            return render_template('operations.html', message=message)

        # Redirect based on selected operation
        if operation == 'flashcards':
            return redirect(url_for('flashcards'))
        elif operation == 'mindmap':
            return redirect(url_for('mindmap'))
        elif operation == 'mnemonic':
            return redirect(url_for('mnemonic'))
        elif operation == 'ppt':
            return redirect(url_for('ppt'))
        elif operation == 'summary':
            return redirect(url_for('summary'))
        elif operation == 'tts':
            return redirect(url_for('tts'))
        elif operation == 'qna':
            return redirect(url_for('qna'))
        elif operation == 'mcq':
            return redirect(url_for('mcq'))
        elif operation == 'suggestions':
            return redirect(url_for('suggestions'))
        elif operation == 'qgen':
            return redirect(url_for('qgen'))
        else:
            message = "Invalid operation selected."
            return render_template('operations.html', message=message)

    # GET request
    return render_template('operations.html')

def clean_text(text):
    # Step 1: Remove common headers/footers and unwanted patterns
    patterns_to_remove = [
        r"(?i)\bunit\s*\d+\b",
        r"(?i)\bchapter\s*\d+\b",
        r"(?i)\bpage\s*\d+\b",
        r"(?i)\bsection\s*\d+\b",
        r"^\s*\d+\s*$",                       # page numbers
        r"^\s*[A-Z]\s*$",                     # single capital letters
        r"(?i)copyright.*",
        r"(?i)all rights reserved.*",
        r"(?i)www\.\S+",
        r"https?://\S+",
        r"^\s*[\-–—]+\s*$",                   # separator lines
    ]

    for pat in patterns_to_remove:
        text = re.sub(pat, "", text, flags=re.MULTILINE)

    # Step 2: Fix hyphenated words split across lines (e.g., "devel-\nopment" → "development")
    text = re.sub(r"(\w+)-\n(\w+)", r"\1\2", text)

    # Step 3: Merge broken lines intelligently
    # Merge lines that do not end with punctuation (., ?, !) into the next line
    text = re.sub(r"(?<![.\?!])\n(?=[a-z])", " ", text)

    # Step 4: Replace multiple newlines with paragraph breaks
    text = re.sub(r"\n\s*\n+", "\n\n", text)

    # Step 5: Remove excessive spaces
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s+", "\n", text)
    text = text.strip()

    # Step 6: Fix encoding issues or extra punctuation
    text = text.replace("•", "-").replace("", "-")
    text = re.sub(r"[“”]", '"', text)
    text = re.sub(r"[‘’]", "'", text)

    # Step 7: Ensure paragraphs start with a capital letter if possible
    lines = []
    for para in text.split("\n"):
        para = para.strip()
        if not para:
            continue
        if len(para) > 2:
            para = para[0].upper() + para[1:]
        lines.append(para)

    cleaned_text = "\n\n".join(lines)
    return cleaned_text



@app.route('/summary', methods=['GET', 'POST'])
def summary():
    text = session.get('input_text', '')
    result = ""
    if request.method == 'POST':
        text = request.form.get('text', '').strip()
        if text:
            result = summarize_text(text)
        else:
            result = "⚠ Please enter some text."
    return render_template('summary.html', text=text,result=result)

@app.route('/flashcards')
def flashcards():
    text = session.get('input_text', '')  # Get the input text from session
    result = generate_flashcards_from_keywords(text)
    return render_template('flashcards.html', result=result)

@app.route('/qgen', methods=['GET', 'POST'])
def qgen():
    questions = []
    text = session.get('input_text', '')
    if request.method == 'POST':
        text = request.form.get('text', '').strip()
        num_q = int(request.form.get('num_questions', 1))
        if text:
            session['input_text'] = text
            questions = process_text_by_sentence(text, num_questions=num_q)
    return render_template('qgen.html', questions=questions, text=text)


@app.route('/mindmap')
def mindmap():
    text = session.get('input_text', '')  # Get the text stored from user
    if not text:
        return render_template('mindmap.html', error="No text provided for mindmap.")

    output_file = generate_mindmap(text, filename="static/mindmap")
    return render_template('mindmap.html', mindmap_image=output_file)



@app.route('/mnemonics')
def mnemonic():
    text = session.get('input_text', '')  # Get stored text
    if not text:
        return render_template('mnemonics.html', error="No text provided.")

    keywords, initials, sentence = create_mnemonic(text)
    return render_template('mnemonics.html', keywords=keywords, initials=initials, sentence=sentence)

@app.route('/ppt')
def ppt():
    text = session.get('input_text', '')  # Get stored text
    if not text:
        return render_template('ppt.html', error="No text provided.")

    ppt_file = generate_ppt(text)
    return render_template('ppt.html', ppt_file=ppt_file)
from flask import send_from_directory

@app.route('/instance/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


@app.route('/tts', methods=['GET', 'POST'])
def tts():
    text = session.get('input_text', '')
    audio_file = None
    if request.method == 'POST':
        text = request.form.get('text')
        if text:
            audio_file = text_to_speech(text)  # Saved in static/output.mp3
    return render_template('tts.html', text=text , audio_file=audio_file)



@app.route('/qna', methods=['GET', 'POST'])
def qna():
    error = None
    results = []
    if request.method == 'POST':
        context = session.get('input_text', '')
        if not context:
            error = "No context text provided."
        else:
            questions = request.form.get('questions', '').splitlines()
            questions = [q.strip() for q in questions if q.strip()]
            if not questions:
                error = "Please enter at least one question."
            else:
                results = get_answers(context, questions)
    return render_template('qna.html', error=error, results=results)


@app.route('/mcq', methods=['GET', 'POST'])
def mcq():
    error = None
    context = session.get('input_text', '').strip()

    if not context:
        error = "No context text available. Please input text first."
        return render_template('mcq.html', error=error)

    if request.method == 'POST':
        action = request.form.get('action')
        quiz = session.get('quiz')

        # --- Preview Mode ---
        if action == 'preview':
            mcq_gen = MCQGenerator()
            questions = mcq_gen.generate_mcqs(context)
            if not questions:
                error = "No questions could be generated from the text."
                return render_template('mcq.html', error=error)
            return render_template('mcq.html', questions=questions, mode="preview")

        # --- Start Quiz ---
        elif action == 'start_quiz':
            mcq_gen = MCQGenerator()
            questions = mcq_gen.generate_mcqs(context)
            if not questions:
                error = "No questions could be generated from the text."
                return render_template('mcq.html', error=error)

            session['quiz'] = {
                'questions': questions,
                'current': 0,
                'score': 0,
                'feedback_mode': False,
                'last_feedback': None
            }
            return redirect(url_for('mcq'))

        # --- Submit Answer ---
        elif action == 'submit_answer':
            selected_option = request.form.get('option')
            quiz = session.get('quiz')

            if not quiz:
                error = "No active quiz found."
                return render_template('mcq.html', error=error)

            current_q = quiz['questions'][quiz['current']]
            correct_answer = current_q['answer']
            is_correct = selected_option == correct_answer
            feedback = "✅ Correct!" if is_correct else "❌ Incorrect!"

            if is_correct:
                quiz['score'] += 1

            quiz['feedback_mode'] = True
            quiz['last_feedback'] = {
                'feedback': feedback,
                'selected_option': selected_option,
                'correct_answer': correct_answer,
                'context': current_q.get('context', ''),
                'is_correct': is_correct
            }

            session['quiz'] = quiz
            return redirect(url_for('mcq'))

        # --- Next Question ---
        elif action == 'next_question':
            quiz = session.get('quiz')
            quiz['feedback_mode'] = False
            quiz['last_feedback'] = None
            quiz['current'] += 1

            if quiz['current'] >= len(quiz['questions']):
                final_score = quiz['score']
                total = len(quiz['questions'])
                session.pop('quiz', None)
                return render_template('mcq.html', mode="result", final_score=final_score, total=total)

            session['quiz'] = quiz
            return redirect(url_for('mcq'))

    # --- GET Request ---
    quiz = session.get('quiz')
    if quiz:
        question = quiz['questions'][quiz['current']]
        if quiz.get('feedback_mode') and quiz.get('last_feedback'):
            fb = quiz['last_feedback']
            return render_template(
                'mcq.html', 
                mode="quiz", 
                question=question, 
                quiz=quiz,
                feedback=fb['feedback'],
                selected_option=fb['selected_option'],
                correct_answer=fb['correct_answer'],
                context=fb['context'],
                is_correct=fb['is_correct']
            )
        else:
            return render_template('mcq.html', mode="quiz", question=question, quiz=quiz)

    return render_template('mcq.html', mode="start")


@app.route('/suggestions', methods=['GET', 'POST'])
def suggestions():
    text = session.get('input_text', '')
    result = None
    if request.method == 'POST':
        text = request.form.get('text', '')
        if text.strip():
            result = suggest_links(text)
            print("\n=== Keywords ===")
            print(result["keywords"])
            print("Query:", result["query"])

            print("\n=== Google Links ===")
            if result["google"]["error"]:
                print(result["google"]["error"])
            elif result["google"]["links"]:
                for i, link in enumerate(result["google"]["links"], 1):
                    print(f"{i}. {link}")
            else:
                print("❌ No Google results")

            print("\n=== YouTube Videos ===")
            for i, (title, url) in enumerate(result["youtube"], 1):
                print(f"{i}. {title}\n   {url}")
    return render_template('suggestions.html', result=result, text=text)


if __name__ == "__main__":
    app.run(debug=True)
