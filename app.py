import os
import uuid
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file
import google.generativeai as genai
from docx_handler import DocxHandler

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
API_KEY = os.environ.get("GEMINI_API_KEY", "")
# Ініціалізуємо Gemini. Встав свій ключ або використовуй змінну середовища
if API_KEY:
    genai.configure(api_key=API_KEY)


def get_rate_in_words(rate_val):
    """Викликає Gemini для точного юридичного перекладу числа в текст за шаблоном"""
    if not API_KEY:
        return "відсутній API ключ Gemini"
    try:
        model = genai.GenerativeModel('gemini-3.1-flash-lite')

        # Примусово форматуємо число до вигляду з двома знаками після крапки
        try:
            formatted_val = f"{float(str(rate_val).replace(',', '.')):.2f}"
        except ValueError:
            formatted_val = str(rate_val)

        prompt = (
            f"Твоє завдання — перетворити дробове число {formatted_val} у суворий юридичний текст українською мовою. "
            f"Ти повинен чітко розписати цілу та дробову частини (соті), використовуючи слова 'процента' або 'процентів'.\n\n"
            f"ОБОВ'ЯЗКОВО притримуйся саме цього шаблону:\n"
            f"- Якщо число закінчується на 1 (наприклад, 1.01): 'одна ціла, одна сота процента'\n"
            f"- Якщо число закінчується на 2, 3, 4 (наприклад, 0.32 чи 2.02): 'нуль цілих, тридцять дві сотих процента' / 'дві цілих, дві сотих процента'\n"
            f"- Для інших випадків (наприклад, 0.74): 'нуль цілих, сімдесят чотири сотих процентів'\n"
            f"- Для рівних чисел (наприклад, 1.00 чи 2.00): 'одна ціла, нуль сотих процентів' / 'дві цілих, нуль сотих процентів'\n\n"
            f"Приклади для точного наслідування:\n"
            f"'0.32' -> 'нуль цілих, тридцять дві сотих процента'\n"
            f"'2.00' -> 'дві цілих, нуль сотих процентів'\n"
            f"'0.65' -> 'нуль цілих, шістдесят п'ять сотих процентів'\n"
            f"'1.50' -> 'одна ціла, п'ятдесят сотих процентів'\n"
            f"'0.02' -> 'нуль цілих, дві сотих процента'\n\n"
            f"Вхідне число: {formatted_val}\n"
            f"Напиши ТІЛЬКИ фінальний текст нижнім регістром (малими літерами). Жодних пояснень, дужок чи крапок наприкінці."
        )

        response = model.generate_content(prompt)
        return response.text.strip().strip("()").strip(".").lower()
    except Exception as e:
        print(f"Помилка Gemini: {e}")
        return "помилка генерації тексту"


# ================= МАРШРУТИ FLASK =================

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

    # Зберігаємо оригінальне ім'я файлу
    original_filename = file.filename

    raw_rate_181 = request.form.get("rate_181", "").strip()
    rate_181 = raw_rate_181
    rate_181_words = ""
    if raw_rate_181:
        try:
            rate_181 = f"{float(raw_rate_181.replace(',', '.')):.2f}"
        except ValueError:
            pass
        rate_181_words = get_rate_in_words(raw_rate_181)

    raw_std = request.form.get("standard_rate", "").strip()
    standard_rate = raw_std
    if raw_std:
        try:
            standard_rate = f"{float(raw_std.replace(',', '.')):.2f}"
        except ValueError:
            pass

    client_data = {
        "standard_rate": standard_rate,
        "rate_181": rate_181,
        "rate_181_words": rate_181_words,
        "reduced_rate": request.form.get("reduced_rate", ""),
        "promo_rate": request.form.get("promo_rate", ""),
        "preferential_rate": request.form.get("preferential_rate", ""),
        "fee": request.form.get("fee", ""),
        "otp": request.form.get("otp", "")
    }

    file_id = str(uuid.uuid4())
    base_path = os.path.join(UPLOAD_FOLDER, f"{file_id}.docx")
    file.save(base_path)

    mod_path = os.path.join(UPLOAD_FOLDER, f"mod_{file_id}.docx")

    try:
        handler = DocxHandler(client_data=client_data, const_path="const.json")
        handler.process_file(base_path, mod_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    # Повертаємо оригінальне ім'я на фронтенд
    return jsonify({"status": "success", "file_id": file_id, "filename": original_filename})


@app.route('/download/<file_id>')
def download(file_id):
    base_path = os.path.join(UPLOAD_FOLDER, f"{file_id}.docx")
    mod_path = os.path.join(UPLOAD_FOLDER, f"mod_{file_id}.docx")

    if not os.path.exists(mod_path):
        return "Файл не знайдено або вже видалено", 404

    return_data = BytesIO()
    with open(mod_path, 'rb') as fo:
        return_data.write(fo.read())
    return_data.seek(0)

    if os.path.exists(base_path): os.remove(base_path)
    if os.path.exists(mod_path): os.remove(mod_path)

    # Отримуємо ім'я файлу з параметрів запиту
    download_name = request.args.get('name', 'ready_contract.docx')

    return send_file(
        return_data,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=download_name
    )


if __name__ == '__main__':
    app.run(debug=True, port=5000)
