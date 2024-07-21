from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)


EXCEL_FILE = 'feedback.xlsx'

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
       
        name = request.form['name']
        roll_no = request.form['rollno']
        class_ = request.form['class']
        school = request.form.getlist('school')
        feedback = request.form['feedback']

       
        teacher_data = []
        subject_data = []
        rating_data = []

        for i in range(1, 8): 
            teacher_key = f'teacher{i}'
            subject_key = f'subject{i}'
            rating_key = f'rating{i}'

            if teacher_key in request.form and subject_key in request.form and rating_key in request.form:
                teacher_data.append(request.form[teacher_key])
                subject_data.append(request.form[subject_key])
                rating_data.append(request.form[rating_key])

      
        new_data = {
            'Name': name,
            'Roll No': roll_no,
            'Class': class_,
            'School': ', '.join(school),
            'Feedback': feedback
        }

       
        for i in range(len(teacher_data)):
            new_data[f'Teacher{i+1}'] = teacher_data[i]
            new_data[f'Subject{i+1}'] = subject_data[i]
            new_data[f'Rating{i+1}'] = rating_data[i]

     
        try:
            df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        except FileNotFoundError:
            df = pd.DataFrame(columns=new_data.keys())

       
        new_row = pd.DataFrame([new_data])
        df = pd.concat([df, new_row], ignore_index=True)

        df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

        return 'Form submitted successfully!'

    except Exception as e:
        return f"An error occurred: {e}"

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
