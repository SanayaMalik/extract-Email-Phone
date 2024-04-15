import os
from flask import Flask, render_template, request, send_file
from assignment import process_folder, write_to_excel

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
        # Ensure the folder path is correctly received from the form
        folder_path = request.form['folder_path']
        # Ensure the folder path is an absolute path
        folder_path = os.path.abspath(folder_path)
        # Ensure the folder path exists
        if not os.path.exists(folder_path):
            return "Folder path does not exist."
        
        output_path = os.path.join(folder_path, 'output.xls')
        data = process_folder(folder_path)
        if data:
            write_to_excel(data, output_path)
            return send_file(output_path, as_attachment=True)
        else:
            return "No email addresses or phone numbers found in the resumes."

if __name__ == '__main__':
    app.run(debug=True)
