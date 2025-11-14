from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
from matcher_core import run_matching_core
import json

app = Flask(__name__)
CORS(app)  # Enable CORS for frontend communication

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/api/run-matching', methods=['POST'])
def run_matching():
    try:
        # Get uploaded files and form data
        student_file = request.files['student_file']
        course_file = request.files['course_file']
        run_name = request.form['run_name']
        
        # Save uploaded files
        student_path = os.path.join(UPLOAD_FOLDER, student_file.filename)
        course_path = os.path.join(UPLOAD_FOLDER, course_file.filename)
        student_file.save(student_path)
        course_file.save(course_path)
        
        # Create a new output directory for this run
        output_path = os.path.join(OUTPUT_FOLDER, run_name)
        os.makedirs(output_path, exist_ok=True)
        
        # Run your core matching algorithm
        result = run_matching_core(student_path, course_path, output_path)
        
        # Send DataFrames back to the frontend
        response = {
            'students': result['results_df'].to_dict('records'),
            'course_report': result['course_report_df'].to_dict('records'),
            'unplaced': result['unplaced_students_df'].to_dict('records'),
            'output_files': result['output_files']
        }
        
        return jsonify(response)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/download/<path:filename>', methods=['GET'])
def download_file(filename):
    try:
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)
