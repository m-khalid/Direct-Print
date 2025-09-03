from flask import Flask, request, jsonify
from flask_cors import CORS
import win32print
import win32api
import tempfile
import os

app = Flask(__name__)
CORS(app)

@app.route('/printers', methods=['GET'])
def list_printers():
    printers = [printer[2] for printer in win32print.EnumPrinters(2)]
    return jsonify({"printers": printers})

@app.route('/print', methods=['POST', 'OPTIONS'])
def print_text():
    if request.method == 'OPTIONS':
        return '', 200

    try:
        data = request.get_json()
        content = data.get("content", "")
        printer_name = data.get("printer") or win32print.GetDefaultPrinter()

        fd, temp_path = tempfile.mkstemp(suffix=".txt")
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(content)

        win32print.SetDefaultPrinter(printer_name)
        win32api.ShellExecute(0, "print", temp_path, None, ".", 0)

        return jsonify({
            "status": "success",
            "message": f"Printed to: {printer_name}"
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

if __name__ == '__main__':
    app.run(port=5005)
