import os
import pandas as pd
import numpy as np

from deferred_acceptance_with_displacement_final4 import (
    read_student_data,
    read_course_data,
    deferred_acceptance_with_displacement,
)

# ------------------------------------------------------------
# JSON CLEANING HELPER  (Option B)
# ------------------------------------------------------------
def _clean_df_for_json(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert DataFrame values into JSON-safe values:
    - NaN -> None
    - NA  -> None
    - +inf / -inf -> None   (frontend will interpret None as "No limit")
    """
    if df is None or df.empty:
        return df

    # Convert unlimited values (Infinity) into JSON null
    df = df.replace([np.inf, -np.inf], None)

    # Convert NaN / NA -> JSON null
    df = df.where(pd.notnull(df), None)

    return df


# ------------------------------------------------------------
# MAIN CORE FUNCTION
# ------------------------------------------------------------
def run_matching_core(
    student_data_path: str,
    course_data_path: str,
    output_folder_path: str,
):
    """
    This function:
    - loads student & course Excel files
    - runs the matching algorithm
    - builds 3 report DataFrames
    - saves Excel files to output folder
    - returns JSON-safe dict for frontend
    """

    # Ensure output folder exists
    os.makedirs(output_folder_path, exist_ok=True)

    # ----------------------------------------
    # LOAD INPUT FILES
    # ----------------------------------------
    student_marks, num_preferences = read_student_data(student_data_path)
    course_data = read_course_data(course_data_path)

    # ----------------------------------------
    # RUN MATCHING ALGORITHM
    # ----------------------------------------
    course_matches, unplaced_students, student_course_assignment = (
        deferred_acceptance_with_displacement(student_marks, course_data, num_preferences)
    )

    # ------------------------------------------------------------
    # BUILD STUDENT PLACEMENT DATAFRAME
    # ------------------------------------------------------------
    student_assignments = []
    placed_students = set()
    pref_cols = [f"Preference {i}" for i in range(1, num_preferences + 1)]

    for course, students in course_matches.items():
        for student in students:
            if student in placed_students:
                continue

            placed_students.add(student)

            sdata = student_marks.get(student)
            if not sdata:
                continue

            entry = sdata.copy()
            entry["Student Name"] = student
            entry["Assigned Course"] = course
            student_assignments.append(entry)

    column_order = ["Student Name", "Assigned Course"] + pref_cols + ["Total Score"]

    results_df = pd.DataFrame(student_assignments)

    # Ensure all required columns exist
    for col in column_order:
        if col not in results_df.columns:
            results_df[col] = pd.NA

    results_df = results_df[column_order]

    # ------------------------------------------------------------
    # COURSE REPORT
    # ------------------------------------------------------------
    report_rows = []

    for course_name, info in course_data.items():
        original_vacancies = info["capacity"]  # may be inf if unlimited
        assigned_students = course_matches.get(course_name, [])
        remaining_vacancies = original_vacancies - len(assigned_students)

        if len(assigned_students) > 0:
            # Lowest total score among those placed
            last_student = min(
                assigned_students,
                key=lambda stu: student_marks[stu]["Total Score"]
            )
            last_total = student_marks[last_student]["Total Score"]
            last_subject_scores = [
                student_marks[last_student].get(subject, "N/A")
                for subject in info["subject_criteria"].keys()
            ]
        else:
            last_student = "N/A"
            last_total = "N/A"
            last_subject_scores = ["N/A"] * len(info["subject_criteria"])

        row = {
            "Course Name": course_name,
            "Original Vacancies": original_vacancies,
            "Remaining Vacancies": remaining_vacancies,
            "Number of students posted": len(assigned_students),
            "Last Ranked Student Posted": last_student,
            "Last Ranked Student Overall Score": last_total,
        }

        # subject_criteria columns
        for subject, score in zip(info["subject_criteria"].keys(), last_subject_scores):
            row[f"Last Ranked Student {subject} Score"] = score

        report_rows.append(row)

    course_report_df = pd.DataFrame(report_rows)

    # ------------------------------------------------------------
    # UNPLACED STUDENTS
    # ------------------------------------------------------------
    unplaced_list = []

    for student in unplaced_students:
        if student not in student_marks:
            continue

        sdata = student_marks[student].copy()
        sdata["Student Name"] = student
        sdata["Reason for not being placed"] = "No available courses in preferences"

        unplaced_list.append(sdata)

    unplaced_students_df = pd.DataFrame(unplaced_list)

    # ------------------------------------------------------------
    # CLEAN FOR JSON (Option B)
    # ------------------------------------------------------------
    results_df = _clean_df_for_json(results_df)
    course_report_df = _clean_df_for_json(course_report_df)
    unplaced_students_df = _clean_df_for_json(unplaced_students_df)

    # ------------------------------------------------------------
    # SAVE OUTPUT EXCEL FILES
    # ------------------------------------------------------------
    students_xlsx = os.path.join(output_folder_path, "outputmatchingresults.xlsx")
    course_xlsx = os.path.join(output_folder_path, "course_report.xlsx")
    unplaced_xlsx = os.path.join(output_folder_path, "unplaced_students_report.xlsx")

    results_df.to_excel(students_xlsx, index=False)
    course_report_df.to_excel(course_xlsx, index=False)
    if not unplaced_students_df.empty:
        unplaced_students_df.to_excel(unplaced_xlsx, index=False)
    else:
        unplaced_xlsx = None

    # ------------------------------------------------------------
    # RETURN JSON DATA FOR FRONTEND
    # ------------------------------------------------------------
    return {
        "students": results_df.to_dict(orient="records"),
        "course_report": course_report_df.to_dict(orient="records"),
        "unplaced": unplaced_students_df.to_dict(orient="records"),
        "output_files": {
            "students": students_xlsx,
            "course_report": course_xlsx,
            "unplaced": unplaced_xlsx,
        },
    }
