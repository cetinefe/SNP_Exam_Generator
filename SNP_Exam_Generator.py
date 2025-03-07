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
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        if not os.path.exists(excel_file_path):
            raise FileNotFoundError(f"Excel file not found: {excel_file_path}")
        
        data = pd.ExcelFile(excel_file_path)
        sheet_data = data.parse('Sheet1')  # Adjust sheet name if necessary
        sheet_data = validate_excel_structure(sheet_data)
        
        # Fix the exam number calculation
        max_exam_number = 0
        if 'Exam Number' in sheet_data.columns:
            # Convert to numeric, treating non-numeric values as NaN
            sheet_data['Exam Number'] = pd.to_numeric(sheet_data['Exam Number'], errors='coerce')
            # Get the maximum value, ignoring NaN
            valid_exam_numbers = sheet_data['Exam Number'].dropna()
            if not valid_exam_numbers.empty:
                max_exam_number = int(valid_exam_numbers.max())

        new_exam_number = max_exam_number + 1

        if len(sheet_data) < sample_size:
            sample_size = len(sheet_data)  # Use all available questions
        questions = sheet_data.sample(sample_size)
        
        # Update with new exam number
        sheet_data.loc[questions.index, "Occurrence"] = sheet_data.loc[questions.index, "Occurrence"].fillna(0) + 1
        sheet_data.loc[questions.index, "Exam Number"] = new_exam_number

        # Save changes back to Excel
        sheet_data.to_excel(excel_file_path, index=False)
        
        output_html_path = os.path.join(output_dir, f"shuffle_exam_test_{new_exam_number}.html")
        html_content = create_html_content(questions, new_exam_number)

        with open(output_html_path, 'w', encoding='utf-8') as file:
            file.write(html_content)

        print(f"HTML file successfully written to {output_html_path}")
    except Exception as e:
        log_error(e)
        raise

def create_html_content(questions, new_exam_number):
    """Create HTML content for the exam."""
    # Create the correct answers dictionary FIRST
    correct_answers_dict = {}
    for i, (_, row) in enumerate(questions.iterrows(), start=1):
        if pd.notna(row['Correct Answers & Selections']):
            answers = [ans.strip() for ans in str(row['Correct Answers & Selections']).split('+')]
            correct_answers_dict[str(i)] = answers

    html_header = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Random Scoped Exam Test {new_exam_number}</title>
    <style>
        body {{font-family: Arial, sans-serif;}}
        h1 {{text-align: center;}}
        .test-container {{margin-bottom: 50px; border: 1px solid #ccc; padding: 20px; border-radius: 8px;}}
        .question {{margin-bottom: 20px; padding: 10px;}}
        .options {{margin-left: 20px;}}
        .metadata {{font-size: 0.9em; color: grey; margin-top: 5px;}}
        .correct {{background-color: #dff0d8;}}
        .incorrect {{background-color: #f2dede;}}
        .missing {{background-color: #d9edf7;}}
        .correct-answer {{color: green; font-weight: bold;}}
        .label-container {{display: inline-block; margin-left: 5px;}}
    </style>
    <script>
        const correctAnswers = {json.dumps(correct_answers_dict, indent=4)};
        
        function checkAnswers(testId) {{
            let score = 0;
            const totalQuestions = document.querySelectorAll(`#${{testId}} .question`).length;
            
            for (let i = 1; i <= totalQuestions; i++) {{
                const key = i.toString();
                const questionDiv = document.querySelector(`#${{testId}} div[data-question="${{key}}"]`);
                
                if (!questionDiv) continue;
                
                const selectedOptions = Array.from(
                    questionDiv.querySelectorAll('input:checked')
                ).map(opt => opt.value.trim());
                
                // Remove any existing status classes
                questionDiv.classList.remove('correct', 'incorrect', 'missing');
                
                // Reset previous markings
                questionDiv.querySelectorAll('.label-container').forEach(label => {{
                    label.classList.remove('correct-answer');
                }});
                
                // Skip questions with no correct answers defined
                if (!correctAnswers[key]) continue;
                
                const correct = correctAnswers[key].map(ans => ans.trim());
                
                if (selectedOptions.length === 0) {{
                    // No answer selected
                    questionDiv.classList.add('missing');
                }} else {{
                    // Sort both arrays for comparison
                    const sortedSelected = selectedOptions.sort();
                    const sortedCorrect = correct.sort();
                    const isEqual = JSON.stringify(sortedSelected) === JSON.stringify(sortedCorrect);
                    
                    if (isEqual) {{
                        score++;
                        questionDiv.classList.add('correct');
                    }} else {{
                        questionDiv.classList.add('incorrect');
                    }}
                }}
                
                // Highlight correct answers
                correct.forEach(correctAns => {{
                    const correctLabel = questionDiv.querySelector(`label[for="q${{key}}_${{correctAns.replace(/\\s/g, '_')}}"]`);
                    if (correctLabel) {{
                        correctLabel.parentElement.classList.add('correct-answer');
                    }}
                }});
            }}
            
            document.querySelector(`#${{testId}} .score`).textContent = 
                `Your score is: ${{score}} out of ${{Object.keys(correctAnswers).length}}`;
        }}
    </script>
</head>
<body>
    <h1>Random Scoped Exam Test {new_exam_number}</h1>
    <div id="test1" class="test-container">
        <div class="score">Your score is: 0 out of {len(questions)}</div>
"""

    question_html = ""
    for i, (_, row) in enumerate(questions.iterrows(), start=1):
        selection_criteria = row['Selection Criteria'] if pd.notna(row['Selection Criteria']) else ''
        input_type = "checkbox" if selection_criteria else "radio"
        metadata = f"{row['Exam #']} | {row['Question #']} | Difficulty: {row['Difficulty Level']} | Domain: {row['Domain']}"
        question_html += f"""
        <div class="question" data-question="{i}">
            <b>Question {i}: {html.escape(str(row['Question Text']))}</b>
            {'<br><i>' + html.escape(selection_criteria) + '</i>' if selection_criteria else ''}
            <div class="metadata">{metadata}</div>
            <div class="options">
        """
        if pd.notna(row['Selections']):
            for option in row['Selections'].split('+'):
                option = option.strip()
                escaped_option = html.escape(option)
                option_id = f"q{i}_{escaped_option.replace(' ', '_')}"
                question_html += f'''
                    <div>
                        <input type="{input_type}" name="question{i}" id="{option_id}" value="{escaped_option}">
                        <div class="label-container">
                            <label for="{option_id}">{escaped_option}</label>
                        </div>
                    </div>
                '''
        else:
            question_html += '<div>No options available for this question.</div>'
        question_html += """</div></div>"""

    html_footer = """
        <button onclick="checkAnswers('test1')">Check Answers</button>
    </div>
</body>
</html>
    """
    
    return html_header + question_html + html_footer

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
