from flask import Flask, request, render_template_string, send_file
import os
from spaceNinja import process_pptx_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['pptx_file']
        if file and file.filename.endswith('.pptx'):
            filename = secure_filename(file.filename)
            input_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(input_path)
            output_path = os.path.join(UPLOAD_FOLDER, filename.replace('.pptx', '_spaced.pptx'))
            process_pptx_file(input_path, output_path)
            return send_file(output_path, as_attachment=True)
        return 'PPTXファイルをアップロードしてください。'
    return render_template_string('''
        <h1>SpaceNinja PPTX スペース挿入ツール</h1>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="pptx_file" accept=".pptx" required>
            <input type="submit" value="アップロードして変換">
        </form>
    ''')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)
