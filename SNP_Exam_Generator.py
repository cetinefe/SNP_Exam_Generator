import sys
import pandas as pd
import json
import os
import argparse
import html

def check_requirements():
    """Check if required packages are installed."""
    try:
        import pandas
        import openpyxl  # Required for Excel support
    except ImportError:
        print("Required packages are missing. Please install them:")
        print("pip install pandas openpyxl")
        sys.exit(1)

# Run package check before imports
check_requirements()

def validate_excel_structure(sheet_data):
    """Ensure required columns exist in the DataFrame."""
    required_columns = [
        "Occurrence", "Exam Number", "Correct Answers & Selections", 
        "Question Text", "Selections", "Selection Criteria", 
        "Exam #", "Question #", "Difficulty Level", "Domain"
    ]
    for column in required_columns:
        if column not in sheet_data.columns:
            sheet_data[column] = None if column != "Occurrence" else 0
    return sheet_data

def generate_exam_html(excel_file_path, output_dir, sample_size=40):
    try:
        if not os.path.exists(excel_file_path):
            raise FileNotFoundError(f"Excel file not found: {excel_file_path}")
        
        data = pd.ExcelFile(excel_file_path)
        sheet_data = data.parse('Sheet1')  # Adjust sheet name if necessary
        sheet_data = validate_excel_structure(sheet_data)
        
        max_exam_number = sheet_data['Exam Number'].max() if pd.notna(sheet_data['Exam Number']).any() else 0
        new_exam_number = int(max_exam_number) + 1

        if len(sheet_data) < sample_size:
            sample_size = len(sheet_data)  # Use all available questions
        questions = sheet_data.sample(sample_size)
        
        sheet_data.loc[questions.index, "Occurrence"] += 1
        sheet_data.loc[questions.index, "Exam Number"] = new_exam_number

        sheet_data.to_excel(excel_file_path, index=False)
        
        output_html_path = os.path.join(output_dir, f"shuffle_exam_test_{new_exam_number}.html")
        html_content = create_html_content(questions, new_exam_number)

        with open(output_html_path, 'w', encoding='utf-8') as file:
            file.write(html_content)

        print(f"HTML file successfully written to {output_html_path}")
    except Exception as e:
        log_error(e)

def create_html_content(questions, new_exam_number):
    """Create HTML content for the exam."""
    html_header = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Random Scoped Exam Test #{new_exam_number}</title>
    <style>
        body {{font-family: Arial, sans-serif;}}
        h1 {{text-align: center;}}
        .test-container {{margin-bottom: 50px; border: 1px solid #ccc; padding: 20px; border-radius: 8px;}}
        .question {{margin-bottom: 20px;}}
        .options {{margin-left: 20px;}}
        .metadata {{font-size: 0.9em; color: grey; margin-top: 5px;}}
        .correct-answer {{color: green; font-weight: bold;}}
        .incorrect-answer {{color: red; font-weight: bold;}}
    </style>
    <script>
        function checkAnswers(testId) {{
            const correctAnswers = 
    """

    correct_answers_dict = {
        i: [html.escape(ans.strip()) for ans in str(row['Correct Answers & Selections']).split(' + ') if pd.notna(row['Correct Answers & Selections'])]
        for i, (_, row) in enumerate(questions.iterrows(), start=1)
    }
    html_header += f"{json.dumps(correct_answers_dict, indent=4)};\n"

    html_js = """
            let score = 0;
            Object.keys(correctAnswers).forEach((key) => {
                const selectedOptions = Array.from(
                    document.querySelectorAll(`#${testId} input[name="question${key}"]:checked`)
                ).map(opt => opt.value.trim());
                const correct = correctAnswers[key];
                if (selectedOptions.sort().toString() === correct.sort().toString()) {
                    score++;
                }
            });
            document.querySelector(`#${testId} .score`).textContent = `Your score is: ${score} out of ${Object.keys(correctAnswers).length}`;
        });
    }
    </script>
</head>
<body>
    <h1>Random Scoped Exam Test #{new_exam_number}</h1>
    <div id="test1" class="test-container">
        <div class="score">Your score is: 0 out of {len(correct_answers_dict)}</div>
    """
    question_html = ""
    for i, (_, row) in enumerate(questions.iterrows(), start=1):
        selection_criteria = row['Selection Criteria'] if pd.notna(row['Selection Criteria']) else ''
        input_type = "checkbox" if selection_criteria else "radio"
        metadata = f"{row['Exam #']} | {row['Question #']} | Difficulty: {row['Difficulty Level']} | Domain: {row['Domain']}"
        question_html += f"""
        <div class="question">
            <b>Question {i}: {html.escape(row['Question Text'])}</b>
            {'<br><i>' + html.escape(selection_criteria) + '</i>' if selection_criteria else ''}
            <div class="metadata">{metadata}</div>
            <div class="options">
        """
        if pd.notna(row['Selections']):
            for option in row['Selections'].split(' + '):
                escaped_option = html.escape(option.strip())
                question_html += f'<input type="{input_type}" name="question{i}" value="{escaped_option}"> <label>{escaped_option}</label><br>'
        else:
            question_html += '<div>No options available for this question.</div>'
        question_html += """</div></div>"""
    html_footer = """
            <button onclick="checkAnswers('test1')">Check Answers</button>
        </div>
    </body>
</html>
    """
    return html_header + html_js + question_html + html_footer

def log_error(e):
    """Log errors to a file."""
    with open("error_log.txt", "a") as log_file:
        log_file.write(f"Error: {e}\n")
    print(f"An error occurred: {e}")

def main():
    parser = argparse.ArgumentParser(description='Generate exam HTML from Excel file')
    parser.add_argument('--excel', '-e', required=True, help='Path to Excel file')
    parser.add_argument('--output', '-o', default='output', help='Output directory for HTML files')
    parser.add_argument('--sample-size', '-n', type=int, default=40, help='Number of questions to sample')
    args = parser.parse_args()
    try:
        generate_exam_html(args.excel, args.output, args.sample_size)
    except Exception as e:
        sys.exit(1)

if __name__ == "__main__":
    main()
