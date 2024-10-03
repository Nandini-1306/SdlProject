from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import csv
import pdfplumber
import pandas as pd


app = Flask(__name__)  # Corrected here
CORS(app)

UPLOAD_FOLDER = 'uploads'
RESULTS_BASE_FOLDER = 'results'
CSV_BASE_FOLDER = 'csv'
EXCEL_BASE_FOLDER = 'excel'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def create_folders_for_pdf(pdf_name):
    base_name = os.path.splitext(pdf_name)[0]
    results_folder = os.path.join(RESULTS_BASE_FOLDER, base_name)
    csv_folder = os.path.join(CSV_BASE_FOLDER, base_name)
    excel_folder = os.path.join(EXCEL_BASE_FOLDER, base_name)
    
    os.makedirs(results_folder, exist_ok=True)
    os.makedirs(csv_folder, exist_ok=True)
    os.makedirs(excel_folder, exist_ok=True)
    
    return results_folder, csv_folder, excel_folder

def extract_data_to_csv(pdf_path, csv_folder):
    csv_path = os.path.join(csv_folder, 'data.csv')
    with pdfplumber.open(pdf_path) as pdf:
        with open(csv_path, 'w', newline='') as csv_file:
            writer = csv.writer(csv_file)
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    writer.writerows(table)
    return csv_path

# def create_files(df, results_folder, excel_folder, subjects):
    sub = [df.iat[1, 0], df.iat[1, 1], df.iat[1, 2]]
    
    for key, value in subjects.items():
        list_length = len(value)
        head = []
        
        if list_length == 6:
            col = value[0]
            type = df.iat[1, col]
            type1 = df.iat[1, col + 3]
            head = sub.copy()
            head.extend([
                f'{key} {type} Total',
                f'{key} {type} Attended',
                f'{key} {type} Percentage',
                f'{key} {type1} Total',
                f'{key} {type1} Attended',
                f'{key} {type1} Percentage'
            ])
        else:
            col = value[0]
            type = df.iat[1, col]
            head = sub.copy()
            head.extend([
                f'{key} {type} Total',
                f'{key} {type} Attended',
                f'{key} {type} Percentage'
            ])
        
        pdf_path = os.path.join(results_folder, f"{key}_report.pdf")
        excel_path = os.path.join(excel_folder, f"{key}_report.xlsx")
        
        data = []
        for i in range(2, len(df)):
            row_data = [
                df.iat[i, 0],  # Assuming this is S.No. or similar
                df.iat[i, 1],  # Assuming this is Enrollment No. or similar
                df.iat[i, 2],  # Assuming this is Name
            ]
            
            if list_length == 6:
                row_data.extend([
                    df.iat[i, value[0]], df.iat[i, value[1]], df.iat[i, value[2]],
                    df.iat[i, value[3]], df.iat[i, value[4]], df.iat[i, value[5]]
                ])
            else:
                row_data.extend([df.iat[i, value[0]], df.iat[i, value[1]], df.iat[i, value[2]]])
            
            data.append(row_data)

        pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
        elements = []
        table = Table([head] + data)  # Include headers and data
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOX', (0, 0), (-1, -1), 2, colors.black),
        ]))
        elements.append(table)
        pdf.build(elements)

        subject_df = pd.DataFrame(data, columns=head)  # Create DataFrame for Excel
        subject_df.to_excel(excel_path, index=False)

# def process_csv_to_df(csv_path, results_folder, excel_folder):
    df = pd.read_csv(csv_path, skiprows=4)

    if df.empty:
        raise ValueError("CSV file is empty or could not be read.")
    
    map_dict = {}
    row_number = 0
    
    for i in range(3, len(df.columns) - 3):  
        if pd.notna(df.iat[row_number, i]):
            sub = df.iat[row_number, i]
            map_dict[sub] = []
            map_dict[sub].append(i)
        else:
            map_dict[sub].append(i)

    create_files(df, results_folder, excel_folder, map_dict)
    return df

# def process_data(pdf_path, pdf_name):
    results_folder, csv_folder, excel_folder = create_folders_for_pdf(pdf_name)
    csv_path = extract_data_to_csv(pdf_path, csv_folder)
    df = process_csv_to_df(csv_path, results_folder, excel_folder)
    return df

@app.route('/')
def index():
    return '''
    <h1>Welcome to the PDF Processing API</h1>
    <p>Use the <a href="/upload">/upload</a> endpoint to upload PDF files.</p>
    <p>Use the <a href="/download/csv">/download/csv</a> endpoint to download the generated CSV file.</p>
    <p>Use the <a href="/download/filename">/download/filename</a> endpoint to download generated PDF reports.</p>
    <p>Use the <a href="/download/excel/filename">/download/excel/filename</a> endpoint to download generated Excel reports.</p>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'pdfFile' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['pdfFile']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and file.filename.endswith('.pdf'):
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        try:
            df = process_data(file_path, file.filename)
            csv_file = 'data.csv'
            base_name = os.path.splitext(file.filename)[0]
            results_folder = os.path.join(RESULTS_BASE_FOLDER, base_name)
            excel_folder = os.path.join(EXCEL_BASE_FOLDER, base_name)
            pdf_files = [f for f in os.listdir(results_folder) if f.endswith('.pdf')]
            excel_files = [f for f in os.listdir(excel_folder) if f.endswith('.xlsx')]
            return jsonify({'csv_file': csv_file, 'pdf_files': pdf_files, 'excel_files': excel_files})
        except ValueError as e:
            return jsonify({'error': str(e)}), 400
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Invalid file format'}), 400

@app.route('/download/<filename>')
def download_file(filename):
    safe_filename = os.path.basename(filename)  # Prevent directory traversal
    base_name = os.path.splitext(filename)[0]
    results_folder = os.path.join(RESULTS_BASE_FOLDER, base_name)
    if filename.endswith('.pdf'):
        return send_from_directory(results_folder, safe_filename)
    elif filename.endswith('.xlsx'):
        excel_folder = os.path.join(EXCEL_BASE_FOLDER, base_name)
        return send_from_directory(excel_folder, safe_filename)
    else:
        return jsonify({'error': 'File not found'}), 404

@app.route('/download/csv')
def download_csv():
    # Ensure a PDF was uploaded before allowing the download
    if not os.listdir(UPLOAD_FOLDER):
        return jsonify({'error': 'No PDF uploaded. Please upload a file first.'}), 400

    # Proceed to fetch the latest CSV if available
    if not os.path.exists(CSV_BASE_FOLDER):
        return jsonify({'error': 'No CSV generated yet. Please upload and process a PDF first.'}), 404

    latest_folder = max([os.path.join(CSV_BASE_FOLDER, d) for d in os.listdir(CSV_BASE_FOLDER)], key=os.path.getctime)
    csv_path = os.path.join(latest_folder, 'data.csv')
    
    if os.path.exists(csv_path):
        return send_from_directory(latest_folder, 'data.csv')
    else:
        return jsonify({'error': 'CSV file not found. Process a PDF first.'}), 404


# @app.route('/download/excel/<filename>')
# def download_excel(filename):
#     safe_filename = os.path.basename(filename)  # Prevent directory traversal
#     base_name = os.path.splitext(filename)[0]
#     excel_folder = os.path.join(EXCEL_BASE_FOLDER, base_name)
#     if os.path.exists(os.path.join(excel_folder, safe_filename)):
#         return send_from_directory(excel_folder, safe_filename)
#     return jsonify({'error': 'Excel file not found'}), 404

if __name__ == '__main__':
    app.run(debug=True)
