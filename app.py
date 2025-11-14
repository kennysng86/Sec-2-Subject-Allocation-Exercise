from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import tempfile

from matcher_core import run_matching_core

app = Flask(__name__)
CORS(app)


@app.route("/", methods=["GET"])
def health_check():
    return jsonify({"status": "ok", "message": "Sec 2 Subject Allocation backend running"})


@app.route("/api/run-matching", methods=["POST"])
def run_matching():
    try:
        # --- Validate inputs ---
        if "student_file" not in request.files or "course_file" not in request.files:
            return jsonify({"error": "Missing student_file or course_file"}), 400

        student_file = request.files["student_file"]
        course_file = request.files["course_file"]
        run_name = request.form.get("run_name", "run")

        # --- Work in a temporary directory ---
        with tempfile.TemporaryDirectory() as tmpdir:
            student_path = os.path.join(tmpdir, "studentdata.xlsx")
            course_path = os.path.join(tmpdir, "coursedata.xlsx")

            student_file.save(student_path)
            course_file.save(course_path)

            output_folder = os.path.join(tmpdir, run_name)
            os.makedirs(output_folder, exist_ok=True)

            # Call your core algorithm
            result = run_matching_core(student_path, course_path, output_folder)

        # result already has JSON-safe lists/dicts
        return jsonify(result)

    except Exception as e:
        # Log full error to server logs
        print("Error in /api/run-matching:", repr(e))
        return jsonify({"error": "Internal server error", "details": str(e)}), 500


if __name__ == "__main__":
    # For local debugging only; Render uses gunicorn
    app.run(host="0.0.0.0", port=5000, debug=True)
