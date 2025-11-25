import sys
import pandas as pd
import json
import os
import argparse
import html
import random
from datetime import datetime


def check_requirements():
    """Check if required packages are installed."""
    try:
        import pandas
    except ImportError:
        print("Required packages are missing. Please install them:")
        print("pip install pandas")
        sys.exit(1)


# Run package check before imports
check_requirements()


def validate_csv_structure(sheet_data):
    """Ensure required columns exist in the DataFrame."""
    required_columns = [
        "Occurrence", "Exam Number", "Correct Answers & Selections",
        "Question Text", "Selections", "Selection Criteria",
        "Exam #", "Question #", "Difficulty Level", "Domain"
    ]
    missing_columns = [col for col in required_columns if col not in sheet_data.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
    return sheet_data


def generate_exam_html(csv_file_path, output_dir, sample_size=40, seed=None):
    """Main logic to read CSV, sample questions, and generate HTML exam."""
    try:
        os.makedirs(output_dir, exist_ok=True)

        if not os.path.exists(csv_file_path):
            raise FileNotFoundError(f"CSV file not found: {csv_file_path}")

        # Apply reproducible random seed
        if seed is not None:
            random.seed(seed)
            print(f"🧬 Using random seed: {seed}")

        # Read CSV
        sheet_data = pd.read_csv(csv_file_path)
        sheet_data = validate_csv_structure(sheet_data)

        # Determine next exam number
        max_exam_number = 0
        if 'Exam Number' in sheet_data.columns:
            sheet_data['Exam Number'] = pd.to_numeric(sheet_data['Exam Number'], errors='coerce')
            valid_exam_numbers = sheet_data['Exam Number'].dropna()
            if not valid_exam_numbers.empty:
                max_exam_number = int(valid_exam_numbers.max())

        new_exam_number = max_exam_number + 1

        # Adjust sample size
        if len(sheet_data) < sample_size:
            sample_size = len(sheet_data)

        # Random sample of questions (reproducible if seed is set)
        questions = sheet_data.sample(sample_size, random_state=seed).copy()

        # Update exam number and occurrence
        sheet_data.loc[questions.index, "Occurrence"] = sheet_data.loc[questions.index, "Occurrence"].fillna(0) + 1
        sheet_data.loc[questions.index, "Exam Number"] = new_exam_number

        # Save updated data
        sheet_data.to_csv(csv_file_path, index=False)

        # Timestamp for filename
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_html_path = os.path.join(
            output_dir,
            f"shuffle_exam_test_{new_exam_number}_{timestamp}.html"
        )

        # Generate HTML content
        html_content = create_html_content(questions, new_exam_number, timestamp, seed)

        with open(output_html_path, 'w', encoding='utf-8') as file:
            file.write(html_content)

        print(f"\n✅ HTML file successfully written to:\n{output_html_path}")

    except Exception as e:
        log_error(e)
        raise


def create_html_content(questions, new_exam_number, timestamp, seed):
    """Create HTML content for the exam."""
    correct_answers_dict = {}
    question_option_map = {}

    for i, (_, row) in enumerate(questions.iterrows(), start=1):
        # Correct answers
        if pd.notna(row['Correct Answers & Selections']):
            answers = [ans.strip() for ans in str(row['Correct Answers & Selections']).split('+')]
            correct_answers_dict[str(i)] = answers

        # Shuffle the Selections (reproducible if seed is set)
        if pd.notna(row['Selections']):
            options = [opt.strip() for opt in str(row['Selections']).split('+')]
            random.shuffle(options)
            question_option_map[str(i)] = options
        else:
            question_option_map[str(i)] = []

    # Use raw f-string to preserve regex and backslashes
    html_header = fr"""
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

                questionDiv.classList.remove('correct', 'incorrect', 'missing');
                questionDiv.querySelectorAll('.label-container').forEach(label => {{
                    label.classList.remove('correct-answer');
                }});

                if (!correctAnswers[key]) continue;

                const correct = correctAnswers[key].map(ans => ans.trim());

                if (selectedOptions.length === 0) {{
                    questionDiv.classList.add('missing');
                }} else {{
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
                    const escapedCorrectAns = correctAns.replace(/([!"#$%&'()*+,.\\/:;<=>?@[\\\]^`{{|}}~])/g, '\\$1');
                    const correctLabel = questionDiv.querySelector(`label[for="q${{key}}_${{escapedCorrectAns.replace(/\\s/g, '_')}}"]`);
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
    <p style="text-align:center; font-size: 0.9em; color: gray;">
        Generated on {timestamp} {'(Seed: ' + str(seed) + ')' if seed is not None else ''}
    </p>
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
        if question_option_map[str(i)]:
            for option in question_option_map[str(i)]:
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
        log_file.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error: {e}\n")
    print(f"⚠️ An error occurred: {e}")


def main():
    parser = argparse.ArgumentParser(description='Generate exam HTML from CSV file')
    parser.add_argument('--csv', '-c', required=True, help='Path to CSV file')
    parser.add_argument('--output', '-o', default='output', help='Output directory for HTML files')
    parser.add_argument('--sample-size', '-n', type=int, default=40, help='Number of questions to sample')
    parser.add_argument('--seed', '-s', type=int, help='Optional random seed for reproducibility')
    args = parser.parse_args()
    try:
        generate_exam_html(args.csv, args.output, args.sample_size, args.seed)
    except Exception as e:
        log_error(e)
        sys.exit(1)


if __name__ == "__main__":
    main()
