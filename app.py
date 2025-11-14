from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import tempfile
import base64
from io import BytesIO

import pandas as pd

from matcher_core import run_matching_core

app = Flask(__name__)
CORS(app)


def df_to_excel_base64(rows):
  """
  Take a list of dicts (rows) and return an Excel file as base64 string.
  If rows is empty, return None.
  """
  if not rows:
      return None
  df = pd.DataFrame(rows)
  buffer = BytesIO()
  with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
      df.to_excel(writer, index=False)
  buffer.seek(0)
  return base64.b64encode(buffer.read()).decode("utf-8")


@app.route("/", methods=["GET"])
def health_check():
  return jsonify(
      {"status": "ok", "message": "Sec 2 Subject Allocation backend running"}
  )


@app.route("/api/run-matching", methods=["POST"])
def run_matching():
    try:
        # --- Validate input files ---
        if "student_file" not in request.files or "course_file" not in request.files:
            return jsonify({"error": "Missing student_file or course_file"}), 400

        student_file = request.files["student_file"]
        course_file = request.files["course_file"]
        run_name = request.form.get("run_name", "run")

        # --- Use temporary directory for processing ---
        with tempfile.TemporaryDirectory() as tmpdir:
            student_path = os.path.join(tmpdir, "studentdata.xlsx")
            course_path = os.path.join(tmpdir, "coursedata.xlsx")

            student_file.save(student_path)
            course_file.save(course_path)

            output_folder = os.path.join(tmpdir, run_name)
            os.makedirs(output_folder, exist_ok=True)

            # Run your core matching logic
            result = run_matching_core(student_path, course_path, output_folder)

        # result is expected to contain JSON-safe lists:
        students = result.get("students", [])
        course_report = result.get("course_report", [])
        unplaced = result.get("unplaced", [])

        # ---------- Build Excel files in-memory (base64) ----------
        excel_files = {
            "students_xlsx": df_to_excel_base64(students),
            "course_report_xlsx": df_to_excel_base64(course_report),
            "unplaced_xlsx": df_to_excel_base64(unplaced),
        }

        # ---------- Build a simple text log ----------
        log_lines = []
        log_lines.append(f"Run name: {run_name}")
        log_lines.append(f"Total students placed: {len(students)}")
        log_lines.append(f"Total courses: {len(course_report)}")
        log_lines.append(f"Total unplaced students: {len(unplaced)}")
        log_lines.append("")
        log_lines.append("Course summary:")
        for course in course_report:
            cname = course.get("Course Name", "Unknown")
            num_posted = course.get("Number of students posted", "N/A")
            remaining = course.get("Remaining Vacancies", "N/A")
            log_lines.append(
                f" - {cname}: posted={num_posted}, remaining_vacancies={remaining}"
            )

        log_text = "\n".join(log_lines)

        payload = {
            "students": students,
            "course_report": course_report,
            "unplaced": unplaced,
            "excel_files": excel_files,
            "log_text": log_text,
        }

        return jsonify(payload)

    except Exception as e:
        print("Error in /api/run-matching:", repr(e))
        return jsonify({"error": "Internal server error", "details": str(e)}), 500



if __name__ == "__main__":
  # For local debugging only; Render uses gunicorn
  app.run(host="0.0.0.0", port=5000, debug=True)
