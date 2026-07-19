import os
import json
import uuid
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file

# ІМПОРТУЄМО САМ КЛАС, А НЕ ФУНКЦІЮ
from docx_handler import DocxHandler

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return jsonify({"error": "Файл не знайдено"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Файл не обрано"}), 400

    # 1. Збираємо дані з форми у словник
    client_data = {
        "standard_rate": request.form.get("standard_rate", ""),
        "reduced_rate": request.form.get("reduced_rate", ""),
        "promo_rate": request.form.get("promo_rate", ""),
        "preferential_rate": request.form.get("preferential_rate", ""),
        "fee": request.form.get("fee", ""),
        "otp": request.form.get("otp", "")
    }

    # 2. Зберігаємо базовий файл з унікальним іменем
    file_id = str(uuid.uuid4())
    base_filename = f"{file_id}.docx"
    base_path = os.path.join(UPLOAD_FOLDER, base_filename)
    file.save(base_path)

    # 3. Визначаємо шлях для збереження готового файлу
    mod_filename = f"mod_{file_id}.docx"
    mod_path = os.path.join(UPLOAD_FOLDER, mod_filename)

    # 4. Обробляємо DOCX НАПРЯМУ ЧЕРЕЗ КЛАС
    try:
        # Ініціалізуємо твій клас, передаючи йому дані з форми
        handler = DocxHandler(client_data=client_data, const_path="const.json")
        # Викликаємо метод обробки
        handler.process_file(base_path, mod_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    # 5. Повертаємо ID файлу для розблокування кнопки скачування
    return jsonify({"status": "success", "file_id": file_id})


@app.route('/download/<file_id>')
def download(file_id):
    base_path = os.path.join(UPLOAD_FOLDER, f"{file_id}.docx")
    mod_path = os.path.join(UPLOAD_FOLDER, f"mod_{file_id}.docx")

    if not os.path.exists(mod_path):
        return "Файл не знайдено або вже видалено", 404

    # Читаємо модифікований файл в оперативну пам'ять
    return_data = BytesIO()
    with open(mod_path, 'rb') as fo:
        return_data.write(fo.read())
    return_data.seek(0)

    # Очищаємо сервер від тимчасових файлів
    if os.path.exists(base_path):
        os.remove(base_path)
    if os.path.exists(mod_path):
        os.remove(mod_path)

    # Віддаємо файл користувачу
    return send_file(
        return_data,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='ready_contract.docx'
    )


if __name__ == '__main__':
    # Запускаємо сервер
    app.run(debug=True, port=5000)
