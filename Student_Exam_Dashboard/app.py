from flask import Flask, render_template, request, redirect, url_for, session, flash
from werkzeug.utils import secure_filename
from docx import Document
from openpyxl import Workbook, load_workbook
from pathlib import Path
import json, random, re, os, uuid
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get("EXAM_SECRET_KEY", "change-this-local-secret")
ROOT = Path(__file__).resolve().parent
UPLOADS = ROOT / "uploads"
RESULTS = ROOT / "results"
EXAMS_DIR = ROOT / "exam_sessions"
BANK_FILE = ROOT / "question_bank.json"
ALLOWED_EXTENSIONS = {".docx"}

UPLOADS.mkdir(exist_ok=True)
RESULTS.mkdir(exist_ok=True)
EXAMS_DIR.mkdir(exist_ok=True)

def exam_state_path(exam_id):
    # The full exam is stored on the server, not in Flask's small browser cookie.
    return EXAMS_DIR / f"{exam_id}.json"

def load_exam_state():
    exam_id = session.get("exam_id")
    if not exam_id:
        return None
    path = exam_state_path(exam_id)
    if not path.exists():
        session.pop("exam_id", None)
        session.pop("duration", None)
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None

def save_exam_state(data):
    exam_id = session.get("exam_id")
    if not exam_id:
        return False
    exam_state_path(exam_id).write_text(
        json.dumps(data, ensure_ascii=False), encoding="utf-8"
    )
    return True

def read_bank():
    if BANK_FILE.exists():
        try:
            return json.loads(BANK_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return []
    return []

def write_bank(questions):
    BANK_FILE.write_text(json.dumps(questions, ensure_ascii=False, indent=2), encoding="utf-8")

def clean_text(text):
    return re.sub(r"\s+", " ", (text or "")).strip()

def paragraph_is_bold(paragraph):
    # Treat a paragraph as bold when at least one non-empty run is explicitly bold.
    return any(run.bold is True and run.text.strip() for run in paragraph.runs)

def parse_docx(path):
    """Read paragraphs separated by blank lines. Expected groups: question + 2 or 4 answers.
    Bold answer text is used as the key; ambiguous groups are imported for review.
    """
    doc = Document(path)
    groups, current = [], []
    for p in doc.paragraphs:
        text = clean_text(p.text)
        if text:
            current.append({"text": text, "bold": paragraph_is_bold(p)})
        elif current:
            groups.append(current)
            current = []
    if current:
        groups.append(current)

    result = []
    for group in groups:
        # Skip groups with too few lines; preserve the rest so an admin can review them.
        if len(group) < 3:
            continue
        question_text = group[0]["text"]
        options = [{"text": item["text"], "is_bold": item["bold"]} for item in group[1:]]
        bold_indexes = [i for i, opt in enumerate(options) if opt["is_bold"]]
        status = "ready" if len(bold_indexes) == 1 else "review"
        correct_index = bold_indexes[0] if len(bold_indexes) == 1 else None
        result.append({
            "id": str(uuid.uuid4()),
            "question": question_text,
            "options": [o["text"] for o in options],
            "correct_index": correct_index,
            "review_status": status,
            "source_option_count": len(options)
        })
    return result

def batch_path(batch):
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", batch).strip("_")
    return RESULTS / f"{safe or 'batch'}.xlsx"

def ensure_batch_file(batch):
    path = batch_path(batch)
    if not path.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(["Timestamp", "Student Name", "Batch", "User ID", "Marks", "Total Questions", "Duration (minutes)"])
        wb.save(path)
    return path

def list_batches():
    batches = []
    for path in sorted(RESULTS.glob("*.xlsx")):
        batches.append(path.stem.replace("_", " "))
    return batches

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "create_batch":
            batch = clean_text(request.form.get("batch_name"))
            try:
                duration = int(request.form.get("duration", "30"))
            except ValueError:
                duration = 30
            if not batch:
                flash("Please enter a batch name.", "error")
            elif not 1 <= duration <= 300:
                flash("Duration must be between 1 and 300 minutes.", "error")
            else:
                settings_path = ROOT / "batch_settings.json"
                settings = json.loads(settings_path.read_text(encoding="utf-8")) if settings_path.exists() else {}
                settings[batch] = {"duration": duration}
                settings_path.write_text(json.dumps(settings, indent=2), encoding="utf-8")
                ensure_batch_file(batch)
                flash(f"Batch '{batch}' is ready.", "success")
            return redirect(url_for("home"))
        if action == "upload_bank":
            file = request.files.get("question_file")
            if not file or not file.filename:
                flash("Choose a Word (.docx) file first.", "error")
            elif Path(file.filename).suffix.lower() not in ALLOWED_EXTENSIONS:
                flash("Please upload a .docx Word document.", "error")
            else:
                name = secure_filename(file.filename)
                path = UPLOADS / name
                file.save(path)
                questions = parse_docx(path)
                if not questions:
                    flash("No question groups were found. Make sure each question and its options are separated from the next question by a blank line.", "error")
                else:
                    write_bank(questions)
                    ready = sum(q["review_status"] == "ready" for q in questions)
                    review = len(questions) - ready
                    flash(f"Imported {len(questions)} question groups. {ready} have one bold answer; {review} need review.", "success")
            return redirect(url_for("home"))
    bank = read_bank()
    settings_path = ROOT / "batch_settings.json"
    settings = json.loads(settings_path.read_text(encoding="utf-8")) if settings_path.exists() else {}
    return render_template("admin_dashboard.html", batches=list_batches(), bank_count=len(bank),
                           review_count=sum(q.get("review_status") != "ready" for q in bank),
                           settings=settings)

@app.route("/review", methods=["GET", "POST"])
def review_questions():
    questions = read_bank()
    if request.method == "POST":
        for q in questions:
            value = request.form.get("correct_" + q["id"])
            if value is not None and value != "":
                try:
                    idx = int(value)
                    if 0 <= idx < len(q["options"]):
                        q["correct_index"] = idx
                        q["review_status"] = "ready"
                except ValueError:
                    pass
        write_bank(questions)
        flash("Question answers saved.", "success")
        return redirect(url_for("review_questions"))
    return render_template("review.html", questions=questions)

@app.route("/student", methods=["GET", "POST"])
def student_entry():
    batches = list_batches()
    if request.method == "POST":
        name = clean_text(request.form.get("student_name"))
        batch = clean_text(request.form.get("batch"))
        user_id = clean_text(request.form.get("user_id"))
        if not name or not batch or not user_id:
            flash("Please fill in all three fields.", "error")
        elif batch not in batches:
            flash("Choose a batch created on the start page.", "error")
        else:
            bank = [q for q in read_bank() if q.get("review_status") == "ready" and q.get("correct_index") is not None and len(q.get("options", [])) >= 2]
            if len(bank) < 50:
                flash(f"At least 50 reviewed questions are needed. Currently available: {len(bank)}.", "error")
            else:
                random.shuffle(bank)
                exam_data = {
                    "name": name, "batch": batch, "user_id": user_id,
                    "questions": bank[:50], "answers": {}, "started_at": datetime.now().isoformat()
                }
                # Store the full question set on the server; keep only a short ID in the cookie.
                exam_id = str(uuid.uuid4())
                exam_state_path(exam_id).write_text(
                    json.dumps(exam_data, ensure_ascii=False), encoding="utf-8"
                )
                session["exam_id"] = exam_id
                settings_path = ROOT / "batch_settings.json"
                settings = json.loads(settings_path.read_text(encoding="utf-8")) if settings_path.exists() else {}
                session["duration"] = int(settings.get(batch, {}).get("duration", 30))
                return redirect(url_for("exam"))
    return render_template("student_entry.html", batches=batches)

@app.route("/exam", methods=["GET", "POST"])
def exam():
    data = load_exam_state()
    if not data:
        flash("Your exam session expired or could not be found. Please enter again.", "error")
        return redirect(url_for("student_entry"))
    if request.method == "POST":
        answers = {}
        for q in data["questions"]:
            value = request.form.get("answer_" + q["id"])
            if value is not None:
                try:
                    idx = int(value)
                    if 0 <= idx < len(q.get("options", [])):
                        answers[q["id"]] = idx
                except ValueError:
                    continue
        data["answers"] = answers
        save_exam_state(data)
        return redirect(url_for("submit_exam"))
    return render_template("exam.html", exam=data, duration=session.get("duration", 30))

@app.route("/submit", methods=["GET", "POST"])
def submit_exam():
    data = load_exam_state()
    if not data:
        flash("Your exam session expired or could not be found. Please enter again.", "error")
        return redirect(url_for("student_entry"))
    if request.method == "POST":
        # POST is used by the exam page's form; the exam route has already saved selected answers.
        pass
    marks = 0
    for q in data["questions"]:
        chosen = data.get("answers", {}).get(q["id"])
        if chosen is not None and chosen == q.get("correct_index"):
            marks += 1
    path = ensure_batch_file(data["batch"])
    wb = load_workbook(path)
    ws = wb["Results"]
    ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), data["name"], data["batch"],
               data["user_id"], marks, len(data["questions"]), session.get("duration", 30)])
    wb.save(path)
    result = {"name": data["name"], "batch": data["batch"], "user_id": data["user_id"],
              "marks": marks, "total": len(data["questions"])}
    exam_id = session.pop("exam_id", None)
    session.pop("duration", None)
    if exam_id:
        try:
            exam_state_path(exam_id).unlink(missing_ok=True)
        except OSError:
            pass
    return render_template("result.html", result=result)

if __name__ == "__main__":
    # Local-use starter. Do not expose this development server to the public internet.
    app.run(debug=True)
