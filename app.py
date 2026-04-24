# app.py — SpecSync Flask Web App
import os
import sys
import json
import time
import zipfile
import tempfile
import threading
import queue
from flask import Flask, render_template, request, jsonify, send_file, Response
from dotenv import load_dotenv

load_dotenv()

from dvp_reader import generate_dvp

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB max upload

STANDARDS_LIBRARY = "standards_library"
OUTPUT_DIR        = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Progress streaming ─────────────────────────────────────────────────────────

progress_queues = {}  # job_id → queue

class QueueLogger:
    """Redirects print() output into a queue for SSE streaming."""
    def __init__(self, q):
        self.q = q
        self.original = sys.stdout

    def write(self, msg):
        if msg.strip():
            self.q.put({"type": "log", "msg": msg.rstrip()})
        self.original.write(msg)

    def flush(self):
        self.original.flush()


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.endswith(".zip"):
        return jsonify({"error": "Please upload a .zip file"}), 400

    job_id = str(int(time.time() * 1000))
    q = queue.Queue()
    progress_queues[job_id] = q

    # Save zip to temp dir
    tmp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(tmp_dir, "rfq.zip")
    f.save(zip_path)

    # Save standards zip if provided
    std_library_path = STANDARDS_LIBRARY  # default to server library
    if 'standards' in request.files:
        std_file = request.files['standards']
        if std_file.filename.endswith('.zip'):
            import shutil
            std_zip_path = os.path.join(tmp_dir, "standards.zip")
            std_file.save(std_zip_path)
            std_tmp = os.path.join(tmp_dir, "standards_lib")
            os.makedirs(std_tmp, exist_ok=True)
            with zipfile.ZipFile(std_zip_path, "r") as z:
                z.extractall(std_tmp)
            for fn in os.listdir(STANDARDS_LIBRARY):
                src = os.path.join(STANDARDS_LIBRARY, fn)
                dst = os.path.join(std_tmp, fn)
                if not os.path.exists(dst):
                    shutil.copy2(src, dst)
            std_library_path = std_tmp

    company_name = request.form.get('company_name', '')
    print(f"   🏢 Company name received: '{company_name}'")

    def run_job():
        try:
            # Extract zip
            q.put({"type": "log", "msg": "📦 Extracting zip..."})
            with zipfile.ZipFile(zip_path, "r") as z:
                z.extractall(tmp_dir)

            # Find the RFQ folder (first non-hidden dir or tmp_dir itself)
            rfq_folder = tmp_dir
            for item in os.listdir(tmp_dir):
                item_path = os.path.join(tmp_dir, item)
                if os.path.isdir(item_path) and not item.startswith(".") and item != "__MACOSX":
                    rfq_folder = item_path
                    break

            # Redirect stdout to queue
            old_stdout = sys.stdout
            sys.stdout = QueueLogger(q)

            output_path = os.path.join(OUTPUT_DIR, f"DVP_{job_id}.xlsx")

            tests = generate_dvp(
                folder_path       = rfq_folder,
                standards_library = std_library_path,
                output_path       = output_path,
                company_name      = company_name,
            )

            sys.stdout = old_stdout

            if tests:
                available   = sum(1 for t in tests if t["available"])
                unavailable = len(tests) - available
                q.put({
                    "type":        "done",
                    "job_id":      job_id,
                    "total":       len(tests),
                    "available":   available,
                    "unavailable": unavailable,
                    "output":      output_path,
                })
            else:
                q.put({"type": "error", "msg": "No tests found in the uploaded folder."})

        except Exception as e:
            sys.stdout = old_stdout if "old_stdout" in dir() else sys.stdout
            q.put({"type": "error", "msg": str(e)})

    threading.Thread(target=run_job, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/stream/<job_id>")
def stream(job_id):
    def generate():
        q = progress_queues.get(job_id)
        if not q:
            yield f"data: {json.dumps({'type': 'error', 'msg': 'Job not found'})}\n\n"
            return

        while True:
            try:
                msg = q.get(timeout=60)
                yield f"data: {json.dumps(msg)}\n\n"
                if msg["type"] in ("done", "error"):
                    break
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'ping'})}\n\n"

        progress_queues.pop(job_id, None)

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/download/<job_id>")
def download(job_id):
    path = os.path.join(OUTPUT_DIR, f"DVP_{job_id}.xlsx")
    if not os.path.exists(path):
        return "File not found", 404
    return send_file(path, as_attachment=True,
                     download_name="DVP_Test_Plan.xlsx")


if __name__ == "__main__":
    app.run(debug=True, threaded=True)