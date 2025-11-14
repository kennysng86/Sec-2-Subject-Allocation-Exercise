import os
import pandas as pd

from deferred_acceptance_with_displacement_final4 import (
    read_student_data,
    read_course_data,
    deferred_acceptance_with_displacement,
)


def _clean_df_for_json(df: pd.DataFrame) -> pd.DataFrame:
    """
    Replace all NaN / NA values with None so that when Flask/jsonify
    converts the DataFrame (via .to_dict), the resulting JSON is valid.
    JSON does not allow NaN, but it allows null (from Python None).
    """
    if df is None or df.empty:
        return df
    # pd.notnull works for both NaN and pd.NA
    return df.where(pd.notnull(df), None)


def run_matching_core(
    student_data_path: str,
    course_data_path: str,
    output_folder_path: str,
):
    """
    Core function for web / API use.

    - No Tkinter
    - No messageboxes
    - Pure: paths in  -> dataframes + output file paths out
    """

    # --- Ensure output folder exists ---
    os.makedirs(output_folder_path, exist_ok=True)

    # --- Read input files using your existing helpers ---
    student_marks, num_preferences = read_student_data(student_data_path)
    course_data = read_course_data(course_data_path)

    # --- Run the existing matching algorithm ---
    course_matches, unplaced_students, student_course_assignment = (
        deferred_acceptance_with_displacement(student_marks, course_data, num_preferences)
    )

    # ------------------------------------------------------------------
    # Build per-student placement DataFrame (similar to your GUI version)
    # ------------------------------------------------------------------
    student_assignments = []
    placed_students = set()

    for course, students in course_matches.items():
        for student in students:
            if student in placed_students:
                continue
            placed_students.add(student)

            sdata = student_marks.get(student)
            if not sdata:
                continue

            record = sdata.copy()
            record["Student Name"] = student
            record["Assigned Course"] = course
            student_assignments.append(record)

    pref_cols = [f"Preference {i}" for i in range(1, num_preferences + 1)]
    column_order = ["Student Name", "Assigned Course"] + pref_cols + ["Total Score"]

    results_df = pd.DataFrame(student_assignments)

    # Ensure required columns exist
    for col in column_order:
        if col not in results_df.columns:
            results_df[col] = pd.NA

    # Reorder columns
    results_df = results_df[column_order]

    # ------------------------------------------------------------------
    # Build course-level summary DataFrame
    # ------------------------------------------------------------------
    course_report_rows = []

    for course_name, course_info in course_data.items():
        original_vacancies = course_info["capacity"]
        assigned_students = course_matches.get(course_name, [])
        remaining_vacancies = original_vacancies - len(assigned_students)
        num_students_posted = len(assigned_students)

        if num_students_posted > 0:
            # Lowest Total Score among those placed in that course
            last_ranked_student = min(
                assigned_students,
                key=lambda stu: student_marks[stu]["Total Score"],
            )
            last_ranked_student_scores = [
                student_marks[last_ranked_student].get(subject, "N/A")
                for subject in course_info["subject_criteria"].keys()
            ]
            last_ranked_student_total_score = student_marks[last_ranked_student]["Total Score"]
        else:
            last_ranked_student = "N/A"
            last_ranked_student_scores = ["N/A"] * len(course_info["subject_criteria"])
            last_ranked_student_total_score = "N/A"

        row = {
            "Course Name": course_name,
            "Original Vacancies": original_vacancies,
            "Remaining Vacancies": remaining_vacancies,
            "Number of students posted": num_students_posted,
            "Last Ranked Student Posted": last_ranked_student,
            "Last Ranked Student Overall Score": last_ranked_student_total_score,
        }

        # Add per-subject scores for last ranked student
        for subject, score in zip(
            course_info["subject_criteria"].keys(), last_ranked_student_scores
        ):
            row[f"Last Ranked Student {subject} Score"] = score

        course_report_rows.append(row)

    course_report_df = pd.DataFrame(course_report_rows)

    # ------------------------------------------------------------------
    # Build unplaced students DataFrame
    # ------------------------------------------------------------------
    unplaced_students_list = []

    for student in unplaced_students:
        # Skip if somehow assigned anyway
        if student in student_course_assignment:
            continue
        if student not in student_marks:
            continue

        student_data = student_marks[student].copy()
        student_data["Student Name"] = student
        student_data["Reason for not being placed"] = (
            "No available courses in preferences"
        )
        unplaced_students_list.append(student_data)

    if unplaced_students_list:
        unplaced_students_df = pd.DataFrame(unplaced_students_list)
    else:
        unplaced_students_df = pd.DataFrame()

    # ------------------------------------------------------------------
    # CLEAN DATAFRAMES FOR JSON (NaN -> None)
    # ------------------------------------------------------------------
    results_df = _clean_df_for_json(results_df)
    course_report_df = _clean_df_for_json(course_report_df)
    unplaced_students_df = _clean_df_for_json(unplaced_students_df)

    # ------------------------------------------------------------------
    # Save Excel outputs (same filenames as your desktop version)
    # ------------------------------------------------------------------
    students_xlsx = os.path.join(output_folder_path, "outputmatchingresults.xlsx")
    course_report_xlsx = os.path.join(output_folder_path, "course_report.xlsx")
    unplaced_xlsx = os.path.join(output_folder_path, "unplaced_students_report.xlsx")
    log_txt = os.path.join(output_folder_path, "matcher_log.txt")

    results_df.to_excel(students_xlsx, index=False)
    course_report_df.to_excel(course_report_xlsx, index=False)

    if not unplaced_students_df.empty:
        unplaced_students_df.to_excel(unplaced_xlsx, index=False)
    else:
        # If you prefer, you can still create an empty file; for now we just skip
        unplaced_xlsx = None

    # ------------------------------------------------------------------
    # Return everything the web app / Flask API needs
    # ------------------------------------------------------------------
    return {
        "results_df": results_df,
        "course_report_df": course_report_df,
        "unplaced_students_df": unplaced_students_df,
        "output_files": {
            "students": students_xlsx,
            "course_report": course_report_xlsx,
            "unplaced": unplaced_xlsx,
            "log": log_txt,
        },
    }
