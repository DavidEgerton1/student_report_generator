import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import json

# Helper function to round scores to the nearest whole number and handle NaN
def round_scores(df):
    return df.applymap(lambda x: round(x) if isinstance(x, (float, int)) and not pd.isna(x) else 5)

# Debugging function to check file access
def check_file_access(file_path, file_description):
    try:
        with open(file_path, "rb") as f:
            print(f"Successfully opened {file_description}: {file_path}")
    except Exception as e:
        print(f"Error opening {file_description}: {file_path}\n{e}")
        raise

# Process behavior file with rounding
def process_behavior(behavior_df):
    behavior_df = round_scores(behavior_df)  # Round scores and replace NaN with 5
    behavior_scores = {}
    for _, row in behavior_df.iterrows():
        student_code = row["student_code"]
        weekly_scores = row.filter(like="week_").values
        behavior_scores[student_code] = sum(weekly_scores) / len(weekly_scores) if len(weekly_scores) > 0 else 0
    return behavior_scores

# Process mini test files with debugging and fallback for missing columns
def process_mini_test(mini_test_df, test_name):
    print(f"{test_name} Columns:", mini_test_df.columns.tolist())  # Debugging: print column names
    mini_test_df = round_scores(mini_test_df)  # Round scores and replace NaN with 5
    test_scores = {}
    for _, row in mini_test_df.iterrows():
        student_code = row["student_code"]
        test_scores[student_code] = {
            "Pronunciation": row.get("pronunciation_and_intonation", 5),  # Default to 5 if column is missing
            "Communication & Interaction": row.get("fluency_coherence", 5),
            "Vocabulary": row.get("vocab_and_lang", 5),
            "Listening for Detail": row.get("listening_section_1", 5),
            "Listening for Main Idea": row.get("listening_section_2", 5)
        }
    return test_scores

# Generate comments from JSON
def generate_comment(skill, score, comments_database):
    categorized_score = str(min(10, max(1, int(score))))  # Ensure score is between 1 and 10
    return random.choice(comments_database[skill][categorized_score])

def generate_final_comment(student_scores, comments_database):
    skills = ["Pronunciation", "Communication & Interaction", "Vocabulary", "Listening for Detail", "Listening for Main Idea", "Behavior"]
    comments = []

    # Add introduction
    overall_avg = sum(student_scores.values()) / len(student_scores)
    introduction = generate_comment("Introduction", overall_avg, comments_database)
    comments.append(introduction)

    # Add skill-specific comments
    for skill in skills:
        comments.append(generate_comment(skill, student_scores[skill], comments_database))

    # Add conclusion
    conclusion = generate_comment("Conclusion", overall_avg, comments_database)
    comments.append(conclusion)

    return " ".join(comments)

# Consolidate all data
def consolidate_data(mini_test_1_scores, mini_test_2_scores, behavior_scores, comments_database):
    consolidated_data = []

    for student_code in mini_test_1_scores.keys():
        if student_code in mini_test_2_scores:
            # Average scores from Mini Test 1 and 2
            avg_scores = {skill: (mini_test_1_scores[student_code][skill] + mini_test_2_scores[student_code][skill]) / 2
                          for skill in mini_test_1_scores[student_code]}
            avg_scores["Behavior"] = behavior_scores.get(student_code, 0)

            # Generate final report comment
            final_comment = generate_final_comment(avg_scores, comments_database)

            # Store consolidated data
            consolidated_data.append({
                "student_code": student_code,
                "Pronunciation": avg_scores["Pronunciation"],
                "Communication & Interaction": avg_scores["Communication & Interaction"],
                "Vocabulary": avg_scores["Vocabulary"],
                "Listening for Detail": avg_scores["Listening for Detail"],
                "Listening for Main Idea": avg_scores["Listening for Main Idea"],
                "Behavior": avg_scores["Behavior"],
                "Final Report Comment": final_comment
            })

    return pd.DataFrame(consolidated_data)

# GUI to select files and generate the report
def generate_report_with_gui():
    root = tk.Tk()
    root.title("Student Report Generator")

    # File paths
    behavior_form_path = tk.StringVar()
    mini_test_1_path = tk.StringVar()
    mini_test_2_path = tk.StringVar()
    json_file_path = tk.StringVar()
    output_file_path = tk.StringVar()

    # Functions to select files
    def select_file(var):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        var.set(file_path)

    def select_json_file(var):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        var.set(file_path)

    def select_output_file(var):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        var.set(file_path)

    # Generate report function
    def generate_report():
        try:
            # Check file access
            check_file_access(behavior_form_path.get(), "Behavior Form")
            check_file_access(mini_test_1_path.get(), "Mini Test 1")
            check_file_access(mini_test_2_path.get(), "Mini Test 2")
            check_file_access(json_file_path.get(), "JSON Comment File")

            # Load data
            behavior_df = pd.read_excel(behavior_form_path.get())
            mini_test_1_df = pd.read_excel(mini_test_1_path.get())
            mini_test_2_df = pd.read_excel(mini_test_2_path.get())

            # Load comments database
            with open(json_file_path.get(), "r") as f:
                comments_database = json.load(f)

            # Process data
            behavior_scores = process_behavior(behavior_df)
            mini_test_1_scores = process_mini_test(mini_test_1_df, "Mini Test 1")
            mini_test_2_scores = process_mini_test(mini_test_2_df, "Mini Test 2")

            # Consolidate data
            consolidated_data = consolidate_data(mini_test_1_scores, mini_test_2_scores, behavior_scores, comments_database)

            # Save to output file
            output_path = output_file_path.get()
            consolidated_data.to_excel(output_path, index=False)
            messagebox.showinfo("Success", f"Report generated successfully at: {output_path}")

        except PermissionError:
            messagebox.showerror("Error", "Permission denied. Please ensure the file is not open or in use.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    # Layout for the GUI
    tk.Label(root, text="Behavior Form:").grid(row=0, column=0, sticky="e")
    tk.Entry(root, textvariable=behavior_form_path, width=50).grid(row=0, column=1)
    tk.Button(root, text="Browse", command=lambda: select_file(behavior_form_path)).grid(row=0, column=2)

    tk.Label(root, text="Mini Test 1:").grid(row=1, column=0, sticky="e")
    tk.Entry(root, textvariable=mini_test_1_path, width=50).grid(row=1, column=1)
    tk.Button(root, text="Browse", command=lambda: select_file(mini_test_1_path)).grid(row=1, column=2)

    tk.Label(root, text="Mini Test 2:").grid(row=2, column=0, sticky="e")
    tk.Entry(root, textvariable=mini_test_2_path, width=50).grid(row=2, column=1)
    tk.Button(root, text="Browse", command=lambda: select_file(mini_test_2_path)).grid(row=2, column=2)

    tk.Label(root, text="JSON Comment File:").grid(row=3, column=0, sticky="e")
    tk.Entry(root, textvariable=json_file_path, width=50).grid(row=3, column=1)
    tk.Button(root, text="Browse", command=lambda: select_json_file(json_file_path)).grid(row=3, column=2)

    tk.Label(root, text="Output File:").grid(row=4, column=0, sticky="e")
    tk.Entry(root, textvariable=output_file_path, width=50).grid(row=4, column=1)
    tk.Button(root, text="Browse", command=lambda: select_output_file(output_file_path)).grid(row=4, column=2)

    tk.Button(root, text="Generate Report", command=generate_report).grid(row=5, column=1, pady=10)

    root.mainloop()

# Run the GUI
generate_report_with_gui()



