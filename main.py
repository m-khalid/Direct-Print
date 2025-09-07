from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import os, tempfile, win32print, win32ui
from PIL import Image, ImageWin
import imgkit

app = Flask(__name__)
CORS(app)

# Path to wkhtmltoimage.exe
WKHTMLTOIMAGE_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe"


def print_image(printer_name, image_path):
    """Scale image to printer page size and print"""
    # Open printer
    hprinter = win32print.OpenPrinter(printer_name)
    printer_info = win32print.GetPrinter(hprinter, 2)
    pdc = win32ui.CreateDC()
    pdc.CreatePrinterDC(printer_name)

    # Get printable area and paper size
    HORZRES = 8
    VERTRES = 10
    PHYSICALWIDTH = 110
    PHYSICALHEIGHT = 111
    PHYSICALOFFSETX = 112
    PHYSICALOFFSETY = 113

    printable_area = (pdc.GetDeviceCaps(HORZRES), pdc.GetDeviceCaps(VERTRES))
    printer_size = (pdc.GetDeviceCaps(PHYSICALWIDTH), pdc.GetDeviceCaps(PHYSICALHEIGHT))
    margins = (pdc.GetDeviceCaps(PHYSICALOFFSETX), pdc.GetDeviceCaps(PHYSICALOFFSETY))

    # Load image
    img = Image.open(image_path)
    img_width, img_height = img.size

    # Scale image to fit page width/height
    ratio = min(printable_area[0] / img_width, printable_area[1] / img_height)
    scaled_width = int(img_width * ratio)
    scaled_height = int(img_height * ratio)

    # Center image on page
    x1 = int((printer_size[0] - scaled_width) / 2)
    y1 = int((printer_size[1] - scaled_height) / 2)
    x2 = x1 + scaled_width
    y2 = y1 + scaled_height

    # Start print job
    pdc.StartDoc("HTML Print Job")
    pdc.StartPage()

    dib = ImageWin.Dib(img)
    dib.draw(pdc.GetHandleOutput(), (x1, y1, x2, y2))

    pdc.EndPage()
    pdc.EndDoc()
    pdc.DeleteDC()
    win32print.ClosePrinter(hprinter)


@app.route('/')
def index():
    return render_template('printer.html')


@app.route('/print-html', methods=['POST'])
def print_html():
    print("inini")
    try:
        data = request.json
        html_content = data.get('html')
        printer_name = data.get('printer', '')

        if not html_content:
            return jsonify({'status': 'error', 'message': 'No HTML content provided'})

        # Save HTML -> PNG
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_img:
            temp_img_path = temp_img.name

        config = imgkit.config(wkhtmltoimage=WKHTMLTOIMAGE_PATH)
        imgkit.from_string(html_content, temp_img_path, config=config)

        # Get default printer if not passed
        if not printer_name:
            printer_name = win32print.GetDefaultPrinter()
        # Print
        print_image(printer_name, temp_img_path)

        # Cleanup
        os.unlink(temp_img_path)

        return jsonify({'status': 'success', 'message': 'Printed successfully'})

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/get-printers', methods=['GET'])
def get_printers():
    printers = [p[2] for p in win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    return jsonify({
        'status': 'success',
        'printers': printers,
        'default_printer': win32print.GetDefaultPrinter()
    })


if __name__ == '__main__':
    app.run(debug=True, port=5000)
