import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import re
import math
from collections import deque  # Import deque for managing displaced students

def read_student_data(file_path):
    df = pd.read_excel(file_path)
    student_preferences = {}
    student_marks = {}

    num_preferences = sum(col.startswith('Preference') for col in df.columns[1:-1])

    for index, row in df.iterrows():
        student_name = row['Student Name']
        preferences = [course.strip() for course in row.iloc[1:1 + num_preferences]]
        marks = {col: row[col] for col in row.index[1 + num_preferences:-1]}

        total_score = row['Total Score']
        marks['Total Score'] = total_score
        marks['Overall Score'] = total_score

        for i, preference in enumerate(preferences):
            marks[f'Preference {i + 1}'] = preference

        student_preferences[student_name] = preferences
        student_marks[student_name] = marks

    return student_marks, num_preferences


def read_course_data(file_path):
    df = pd.read_excel(file_path)
    course_data = {}

    for index, row in df.iterrows():
        course_name = row['Course Name']
        group = row.get('Group')
        group_constraint = row.get('Group Constraint')

        subject_criteria = {col: row[col] for col in row.index[2:]}

        # Treat capacity as infinite if group constraint is NaN or empty
        capacity = row['Capacity'] if pd.notna(row['Capacity']) else None

        if capacity is None:
            if pd.isna(group_constraint):
                capacity = float('inf')  # Treat empty group constraint as infinite
            else:
                capacity = group_constraint

        subject_criteria_dict = {}
        for subject, criteria in subject_criteria.items():
            if pd.notna(criteria):
                match = re.match(r'([<>]=?)\s*(\d+)', str(criteria))
                if match:
                    inequality, value = match.groups()
                    subject_criteria_dict[subject] = (inequality, int(value))

        # Handle tiebreaker subjects
        tiebreaker_subjects = row.get('Tiebreaker Subjects')
        if pd.notna(tiebreaker_subjects) and tiebreaker_subjects.strip():
            tiebreaker_subjects = [subject.strip() for subject in tiebreaker_subjects.split(",")]
        else:
            tiebreaker_subjects = []  # No tiebreakers specified, default to an empty list

        course_data[course_name] = {
            'capacity': capacity,
            'subject_criteria': subject_criteria_dict,
            'group': group,
            'group_constraint': group_constraint,
            'tiebreaker_subjects': tiebreaker_subjects  # Add tiebreaker subjects
        }

    return course_data



def student_meets_course_criteria(student_name, student_marks, preferred_course, course_data):
    """
    Check if a student meets the criteria for a specific course.
    """
    course_info = course_data.get(preferred_course)
    if not course_info:
        return False

    # Ensure the student meets the subject criteria
    return all(
        compare_subject_score(student_marks[student_name].get(subject, 0), inequality, value)
        for subject, (inequality, value) in course_info['subject_criteria'].items()
    )


def check_group_vacancies(course_data, course_matches, preferred_course, student_marks):
    """
    Check if there is still capacity within the group for the preferred course.
    Returns False if the group is full, otherwise True.
    """
    group_constraint = course_data[preferred_course].get('group_constraint')
    if group_constraint:
        group_courses = [course_name for course_name, course_info in course_data.items()
                         if course_info.get('group') == course_data[preferred_course]['group']]
        total_assigned_students = sum(len(course_matches.get(course_name, [])) for course_name in group_courses)
        if total_assigned_students >= group_constraint:
            return False
    return True

from collections import deque
import pandas as pd

# --- Hardened numeric comparisons (treat non-numeric as failing the criterion) ---
import math

def _to_num(x):
    """
    Convert x to float if possible.
    - Returns NaN for blanks/None/non-numeric so comparisons can 'fail closed'.
    """
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return float('nan')
        # strings like '  75 ' are fine; 'VR' will go to except and return NaN
        return float(x)
    except (ValueError, TypeError):
        return float('nan')

def compare_subject_score(student_score, inequality, value):
    """
    Compare a student's score against a criterion like '>= 70' or '<= 60'.
    - Non-numeric student scores (e.g., 'VR', 'ABS', blanks) => treated as NOT meeting the criterion.
    """
    s = _to_num(student_score)
    v = _to_num(value)

    # If either side isn't numeric, treat as not meeting the criterion
    if math.isnan(s) or math.isnan(v):
        return False

    if inequality == '>=':
        return s >= v
    elif inequality == '<=':
        return s <= v
    return False


# Try to place a student in the course with displacement logic
def try_place_student_in_course(student_name, student_marks, preferred_course, course_data, course_matches, student_course_assignment, group_capacity_tracker, unplaced_students):
    print(f"\nConsidering student: {student_name} for course: {preferred_course}")

    # ---------------- helpers ----------------
    def _safe_tie_score(x):
        """Numeric-safe value for tiebreaker tuple. Non-numeric -> very low so it won't unfairly win ties."""
        val = _to_num(x)
        return val if not math.isnan(val) else float('-inf')

    def _total_score(name):
        """Numeric Total Score (NaN if invalid)."""
        return _to_num(student_marks[name].get('Total Score'))

    def _place_into(course_to_place):
        """Append student to course_to_place and record assignment (no double-append)."""
        if course_to_place not in course_matches:
            course_matches[course_to_place] = []
        if student_name not in course_matches[course_to_place]:
            course_matches[course_to_place].append(student_name)
        student_course_assignment[student_name] = course_to_place
        return True

    def _remove_from_course(student_to_remove):
        """Remove a student from whichever course currently holds them (if any)."""
        prev = student_course_assignment.get(student_to_remove)
        if prev and prev in course_matches and student_to_remove in course_matches[prev]:
            course_matches[prev].remove(student_to_remove)
        if student_to_remove in student_course_assignment:
            del student_course_assignment[student_to_remove]

    # --------------- course presence ---------------
    course_info = course_data.get(preferred_course)
    if not course_info:
        print(f"Course {preferred_course} not found in course data.")
        return False

    # --------------- subject criteria ---------------
    # Ensure the student meets all subject criteria (numeric-safe)
    if not all(
        compare_subject_score(student_marks[student_name].get(subject, None), inequality, value)
        for subject, (inequality, value) in course_info.get('subject_criteria', {}).items()
    ):
        print(f"Student {student_name} does not meet the criteria for {preferred_course}.\n")
        return False

    capacity = course_info.get('capacity')
    current_students = course_matches.get(preferred_course, [])

    group_name = course_info.get('group')
    group_constraint = course_info.get('group_constraint')

    # --------------- non-grouped capacity placement ---------------
    if not group_name or not pd.notna(group_name):
        # interpret capacity: None/NaN => infinite
        if capacity is not None and pd.notna(capacity):
            try:
                cap_val = int(capacity)
            except Exception:
                cap_val = None
        else:
            cap_val = None  # infinite

        if cap_val is not None and len(current_students) >= cap_val:
            print(f"Course {preferred_course} is full. Cannot place {student_name}.")
            return False

        # place (remove from old course first to avoid duplicates)
        if student_course_assignment.get(student_name) and student_course_assignment[student_name] != preferred_course:
            _remove_from_course(student_name)

        print(f"Course {preferred_course} has capacity. Placing student {student_name}.")
        placed = _place_into(preferred_course)
        if student_name in unplaced_students:
            unplaced_students.remove(student_name)
        return placed

    # --------------- group-constrained logic ---------------
    # Collect all courses in this group
    group_courses = [cname for cname, info in course_data.items() if info.get('group') == group_name]

    # Ensure course_matches entries exist for all group courses (debug visibility)
    for cname in group_courses:
        if cname not in course_matches:
            print(f"Initializing missing course: {cname}")
            course_matches[cname] = []
        print(f"Current students in course '{cname}': {course_matches[cname]}")

    # Compute numeric group limit (None => treat as no cap)
    if group_constraint is not None and pd.notna(group_constraint):
        try:
            group_limit = int(group_constraint)
        except Exception:
            group_limit = None
    else:
        group_limit = None

    all_group_students = []
    for cname in group_courses:
        all_group_students.extend(course_matches.get(cname, []))
    total_group_students = len(all_group_students)

    # If no valid cap or cap not reached -> place directly
    if (group_limit is None) or (total_group_students < group_limit):
        print(f"Group {group_name}: capacity available (limit={group_limit}, used={total_group_students}). Placing {student_name}.")
        if student_course_assignment.get(student_name) and student_course_assignment[student_name] != preferred_course:
            _remove_from_course(student_name)
        placed = _place_into(preferred_course)
        if group_limit is not None:
            group_capacity_tracker[group_name] = total_group_students + 1
        if student_name in unplaced_students:
            unplaced_students.remove(student_name)
        return placed

    # Group is full: consider displacement
    print(f"Group {group_name} has reached its constraint. Checking for displacement across all group courses.")

    # Find lowest-merit student by Total Score first (numeric-safe)
    if not all_group_students:
        return False  # guard

    lowest_merit_student = min(all_group_students, key=_total_score)

    candidate_total = _total_score(student_name)
    lowest_total = _total_score(lowest_merit_student)

    if math.isnan(candidate_total):
        print(f"Candidate {student_name} has non-numeric Total Score; cannot displace.")
        return False

    if candidate_total > lowest_total:
        # Displace on higher Total Score
        print(f"Student {student_name} will displace {lowest_merit_student} from group {group_name} (Total Score comparison).")
        # remove displaced from whichever course
        for cname in group_courses:
            if lowest_merit_student in course_matches.get(cname, []):
                course_matches[cname].remove(lowest_merit_student)
                if lowest_merit_student in student_course_assignment:
                    del student_course_assignment[lowest_merit_student]
                break

        # place candidate
        if student_course_assignment.get(student_name) and student_course_assignment[student_name] != preferred_course:
            _remove_from_course(student_name)
        course_matches.setdefault(preferred_course, []).append(student_name)
        student_course_assignment[student_name] = preferred_course

        # add displaced to unplaced list and return their name (your original behavior)
        unplaced_students.append(lowest_merit_student)
        return lowest_merit_student

    elif candidate_total == lowest_total:
        # Tie: use numeric-safe tiebreakers
        tiebreaker_subjects = course_info.get('tiebreaker_subjects', [])
        if tiebreaker_subjects:
            # Re-identify "lowest" under lexicographic tiebreaker tuple
            lowest_merit_student = min(
                all_group_students,
                key=lambda s: tuple(_safe_tie_score(student_marks[s].get(subj)) for subj in tiebreaker_subjects)
            )

            candidate_tuple = tuple(_safe_tie_score(student_marks[student_name].get(subj)) for subj in tiebreaker_subjects)
            lowest_tuple    = tuple(_safe_tie_score(student_marks[lowest_merit_student].get(subj)) for subj in tiebreaker_subjects)

            if candidate_tuple > lowest_tuple:
                print(f"Student {student_name} will displace {lowest_merit_student} from group {group_name} (Tiebreaker comparison).")
                for cname in group_courses:
                    if lowest_merit_student in course_matches.get(cname, []):
                        course_matches[cname].remove(lowest_merit_student)
                        if lowest_merit_student in student_course_assignment:
                            del student_course_assignment[lowest_merit_student]
                        break

                if student_course_assignment.get(student_name) and student_course_assignment[student_name] != preferred_course:
                    _remove_from_course(student_name)
                course_matches.setdefault(preferred_course, []).append(student_name)
                student_course_assignment[student_name] = preferred_course

                unplaced_students.append(lowest_merit_student)
                return lowest_merit_student
            else:
                print(f"Student {student_name} cannot displace {lowest_merit_student} (Tiebreaker comparison).")
                return False
        else:
            print(f"No tiebreaker subjects specified for group {group_name}, using Total Score only.")
            return False

    # No advantage -> cannot place
    print(f"Student {student_name} cannot displace any student in group {group_name} (Total Score comparison).")
    return False


from collections import deque


def deferred_acceptance_with_displacement(student_marks, course_data, num_preferences):
    course_matches = {}  # Course to students mapping
    student_course_assignment = {}  # Track which course each student is placed in
    group_capacity_tracker = {}  # Track how many students are placed in each group
    unplaced_students = []
    displaced_students_queue = deque()  # Queue for displaced students
    preference_tracker = {}  # Tracks each student's last attempted preference

    # Queue for students and their current preference being processed
    students_to_process = deque([(student, 1) for student in student_marks.keys()])

    while students_to_process or displaced_students_queue:
        if displaced_students_queue:
            # Prioritize displaced students
            current_student, current_preference = displaced_students_queue.popleft()
        else:
            # Process new students if no displaced students
            current_student, current_preference = students_to_process.popleft()

        # Update preference_tracker for accurate tracking
        preference_tracker[current_student] = current_preference

        # Check if the student has exhausted all preferences
        if current_preference > num_preferences:
            print(f"Student {current_student} has exhausted all preferences.")
            unplaced_students.append(current_student)
            continue

        # Process the student's current preference
        marks = student_marks[current_student]
        preferred_course = marks.get(f'Preference {current_preference}')

        if not preferred_course:
            # Move to the next preference if no preference is listed
            students_to_process.append((current_student, current_preference + 1))
            continue

        # Attempt to place the student
        print(f"\nProcessing {current_student} for preference {current_preference}: {preferred_course}")
        result = try_place_student_in_course(
            current_student, student_marks, preferred_course,
            course_data, course_matches, student_course_assignment,
            group_capacity_tracker, unplaced_students
        )

        if result is True:
            # Successful placement; remove any previous assignment if it exists
            if current_student in student_course_assignment:
                prior_course = student_course_assignment[current_student]
                print(f"Removing previous assignment of {current_student} from {prior_course}")
                course_matches[prior_course].remove(current_student)
                del student_course_assignment[current_student]

            student_course_assignment[current_student] = preferred_course
            course_matches.setdefault(preferred_course, []).append(current_student)
            print(f"{current_student} placed in {preferred_course}")

        elif isinstance(result, str):
            # A student was displaced; use preference_tracker to accurately queue them for the next preference
            displaced_student = result
            next_preference = preference_tracker[displaced_student] + 1
            print(f"Displaced student {displaced_student} re-added to queue for next preference: {next_preference}")
            displaced_students_queue.append((displaced_student, next_preference))
        else:
            # Student was not placed; try the next preference
            students_to_process.append((current_student, current_preference + 1))

    return course_matches, unplaced_students, student_course_assignment




def create_course_report(course_data, course_matches, student_marks, output_folder_path):
    course_report = []

    for course_name, course_info in course_data.items():
        original_vacancies = course_info['capacity']
        assigned_students = course_matches.get(course_name, [])
        remaining_vacancies = original_vacancies - len(assigned_students)
        num_students_posted = len(assigned_students)

        if num_students_posted > 0:
            last_ranked_student = min(assigned_students, key=lambda student: student_marks[student]['Total Score'])
            last_ranked_student_scores = [student_marks[last_ranked_student][subject] for subject in course_info['subject_criteria'].keys()]
            last_ranked_student_total_score = student_marks[last_ranked_student]['Total Score']
        else:
            last_ranked_student = "N/A"
            last_ranked_student_scores = ["N/A"] * len(course_info['subject_criteria'])
            last_ranked_student_total_score = "N/A"

        course_report.append({
            'Course Name': course_name,
            'Original Vacancies': original_vacancies,
            'Remaining Vacancies': remaining_vacancies,
            'Number of students posted': num_students_posted,
            'Last Ranked Student Posted': last_ranked_student,
            **{f'Last Ranked Student {subject} Score': score for subject, score in zip(course_info['subject_criteria'].keys(), last_ranked_student_scores)},
            'Last Ranked Student Overall Score': last_ranked_student_total_score
        })

    course_report_df = pd.DataFrame(course_report)
    course_report_output_file_path = os.path.join(output_folder_path, 'course_report.xlsx')
    with pd.ExcelWriter(course_report_output_file_path) as writer:
        course_report_df.to_excel(writer, sheet_name='Course Report', index=False)

    print(f"Course report saved to {course_report_output_file_path}")


def create_unplaced_students_report(unplaced_students, student_marks, num_preferences, output_folder_path, student_course_assignment):
    """
    Creates a report of unplaced students who have exhausted all their preferences.
    """
    unplaced_students_list = []

    for student in unplaced_students:
        # Skip students who were successfully placed
        if student in student_course_assignment:
            continue

        if student not in student_marks:
            continue  # Skip if the student has no recorded marks (this can avoid errors)
            
        student_data = student_marks[student].copy()
        student_data['Student Name'] = student
        student_data['Reason for not being placed'] = "No available courses in preferences"
        unplaced_students_list.append(student_data)

    # Check if the unplaced students list is empty
    if not unplaced_students_list:
        print("No unplaced students found.")
        return  # If no unplaced students, don't generate the report

    # Create a DataFrame with the required columns
    unplaced_students_column_order = ['Student Name', 'Reason for not being placed'] + [f'Preference {i}' for i in range(1, num_preferences + 1)]

    unplaced_students_df = pd.DataFrame(unplaced_students_list, columns=unplaced_students_column_order)
    unplaced_students_report_path = os.path.join(output_folder_path, 'unplaced_students_report.xlsx')
    
    # Save the DataFrame to Excel
    with pd.ExcelWriter(unplaced_students_report_path) as writer:
        unplaced_students_df.to_excel(writer, sheet_name='Unplaced Students Report', index=False)

    print(f"Unplaced students report saved to {unplaced_students_report_path}")



def run_matching_algorithm():
    """
    Runs the matching end-to-end, logging to both console and <output>/matcher_log.txt.
    """
    import os, sys, traceback
    from datetime import datetime
    from tkinter import messagebox

    # ----------------- Validate inputs from GUI -----------------
    global student_data_path_var, course_data_path_var, output_folder_path_var

    student_data_path = (student_data_path_var.get() or "").strip()
    course_data_path  = (course_data_path_var.get() or "").strip()
    output_folder_path = (output_folder_path_var.get() or "").strip()

    if not student_data_path:
        messagebox.showerror("Error", "Please select the Student Data file.")
        return
    if not course_data_path:
        messagebox.showerror("Error", "Please select the Course Data file.")
        return
    if not output_folder_path:
        messagebox.showerror("Error", "Please select an Output Folder.")
        return

    os.makedirs(output_folder_path, exist_ok=True)

    # ----------------- Set up tee logging to console + file -----------------
    log_file_path = os.path.join(output_folder_path, "matcher_log.txt")

    class _Tee:
        def __init__(self, *streams):
            self._streams = streams
        def write(self, data):
            for s in self._streams:
                try:
                    s.write(data)
                    s.flush()
                except Exception:
                    pass
        def flush(self):
            for s in self._streams:
                try:
                    s.flush()
                except Exception:
                    pass

    original_stdout = sys.stdout
    original_stderr = sys.stderr
    log_fh = open(log_file_path, "w", encoding="utf-8", buffering=1)
    sys.stdout = _Tee(original_stdout, log_fh)
    sys.stderr = _Tee(original_stderr, log_fh)

    print(f"\n=== Deferred Acceptance Matching Log Started {datetime.now()} ===")
    print(f"Student Data: {student_data_path}")
    print(f"Course Data:  {course_data_path}")
    print(f"Output Folder: {output_folder_path}\n")

    try:
        # -------------- Read inputs --------------
        print("Reading student data...")
        student_marks, num_preferences = read_student_data(student_data_path)

        print("Reading course data...")
        course_data = read_course_data(course_data_path)

        # Initialize course_matches for all courses with empty lists
        print("Initializing course match containers...")
        course_matches = {course_name: [] for course_name in course_data.keys()}

        # -------------- Run the core algorithm --------------
        print("Running deferred acceptance with displacement...")
        course_matches, unplaced_students, student_course_assignment = deferred_acceptance_with_displacement(
            student_marks, course_data, num_preferences
        )

        # -------------- Build per-student assignment table --------------
        print("Assembling per-student assignment table...")
        student_assignments = []
        placed_students = set()

        for course, students in course_matches.items():
            for student in students:
                if student in placed_students:
                    print(f"Warning: Skipping duplicate assignment for student {student}.")
                    continue
                placed_students.add(student)

                sdata = student_marks.get(student)
                if not sdata:
                    print(f"ERROR: Student {student} not found in student_marks!")
                    continue

                record = sdata.copy()
                record["Student Name"] = student
                record["Assigned Course"] = course
                student_assignments.append(record)
                print(f"DEBUG: Added student {student} to the assignment for course {course}")

        # Ordered columns for output (missing keys will show as NaN)
        pref_cols = [f"Preference {i}" for i in range(1, num_preferences + 1)]
        column_order = ["Student Name", "Assigned Course"] + pref_cols + ["Total Score"]

        results_df = pd.DataFrame(student_assignments)
        # Ensure we include all desired columns (add missing ones, preserve order)
        for col in column_order:
            if col not in results_df.columns:
                results_df[col] = pd.NA
        results_df = results_df[column_order]

        # -------------- Consistency checks --------------
        all_students_in_matches = {stu for lst in course_matches.values() for stu in lst}
        if all_students_in_matches != placed_students:
            print("ERROR: Mismatch between placed students and students in final matches! Check output generation.")

        output_students = set(results_df["Student Name"].dropna())
        missing_students = placed_students - output_students
        if missing_students:
            print(f"Warning: These placed students were not written to the output (possibly missing data): {missing_students}")

        # -------------- Save outputs --------------
        output_file_path = os.path.join(output_folder_path, "outputmatchingresults.xlsx")
        print(f"Writing student assignments to: {output_file_path}")
        results_df.to_excel(output_file_path, index=False)

        print("Creating course and unplaced student reports...")
        create_course_report(course_data, course_matches, student_marks, output_folder_path)
        create_unplaced_students_report(unplaced_students, student_marks, num_preferences, output_folder_path, student_course_assignment)

        print("\n=== Matching completed successfully ===")
        print(f"Files created in: {output_folder_path}")
        print(f"- outputmatchingresults.xlsx")
        print(f"- course_report.xlsx")
        print(f"- unplaced_students_report.xlsx")
        print(f"- matcher_log.txt")

        messagebox.showinfo("Success", f"Matching complete.\nResults saved to:\n{output_folder_path}")

    except Exception as e:
        print("\n=== ERROR DURING MATCHING ===")
        print("Exception:", repr(e))
        traceback.print_exc()
        messagebox.showerror("Error", f"An error occurred.\nSee matcher_log.txt in the output folder for details.")
    finally:
        # Restore console streams and close file handle
        try:
            sys.stdout = original_stdout
            sys.stderr = original_stderr
            log_fh.close()
        except Exception:
            pass



def select_student_data():
    global student_data_path_var
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    student_data_path_var.set(file_path)


def select_course_data():
    global course_data_path_var
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    course_data_path_var.set(file_path)


def select_output_folder():
    global output_folder_path_var
    folder_path = filedialog.askdirectory()
    output_folder_path_var.set(folder_path)


def main():

    # Create GUI
    root = tk.Tk()
    root.title("Deferred Acceptance Matching")

    # Variables to store file paths
    student_data_path_var = tk.StringVar()
    course_data_path_var = tk.StringVar()
    output_folder_path_var = tk.StringVar()

    # Label and entry for student data
    tk.Label(root, text="Student Data:").grid(row=0, column=0)
    tk.Entry(root, textvariable=student_data_path_var, width=50).grid(row=0, column=1)
    tk.Button(root, text="Browse", command=select_student_data).grid(row=0, column=2)

    # Label and entry for course data
    tk.Label(root, text="Course Data:").grid(row=1, column=0)
    tk.Entry(root, textvariable=course_data_path_var, width=50).grid(row=1, column=1)
    tk.Button(root, text="Browse", command=select_course_data).grid(row=1, column=2)

    # Label and entry for output folder
    tk.Label(root, text="Output Folder:").grid(row=2, column=0)
    tk.Entry(root, textvariable=output_folder_path_var, width=50).grid(row=2, column=1)
    tk.Button(root, text="Browse", command=select_output_folder).grid(row=2, column=2)

    # Button to run matching algorithm
    tk.Button(root, text="Run Matching Algorithm", command=run_matching_algorithm).grid(row=3, column=1)

    root.mainloop()

if __name__ == "__main__":
    main()

