from flask import Flask, render_template, request, jsonify
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/compare', methods=['POST'])
def compare_files():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': '두 파일을 모두 업로드해주세요.'}), 400
    
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    if file1.filename == '' or file2.filename == '':
        return jsonify({'error': '파일이 선택되지 않았습니다.'}), 400

    path1 = os.path.join(app.config['UPLOAD_FOLDER'], 'file1.xlsx')
    path2 = os.path.join(app.config['UPLOAD_FOLDER'], 'file2.xlsx')
    
    file1.save(path1)
    file2.save(path2)

    try:
        df1 = pd.read_excel(path1)
        df2 = pd.read_excel(path2)
        
        # TODO: Implement actual comparison logic based on user columns
        # For now, returning column names to let user choose or just raw data preview
        
        return jsonify({
            'message': '파일 업로드 성공',
            'file1_columns': df1.columns.tolist(),
            'file2_columns': df2.columns.tolist(),
            'file1_preview': df1.head().to_dict(orient='records'),
            'file2_preview': df2.head().to_dict(orient='records')
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
