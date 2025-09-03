from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import subprocess
import os
import tempfile
import time

app = Flask(__name__)
CORS(app)  # Enable CORS

@app.route('/')
def index():
    return render_template('printer.html')

@app.route('/print-html', methods=['POST'])
def print_html():
    try:
        # Get HTML content from request
        data = request.json
        html_content = data.get('html')
        printer_name = data.get('printer', '')  # Optional printer name

        if not html_content:
            return jsonify({'status': 'error', 'message': 'No HTML content provided'})

        # Create temporary HTML file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as temp_file:
            temp_file.write(html_content)
            temp_html_path = temp_file.name

        # Print based on operating system
        success = print_html_file(temp_html_path, printer_name)

        # Clean up
        time.sleep(1)  # Wait a bit before deleting
        os.unlink(temp_html_path)

        if success:
            return jsonify({'status': 'success', 'message': 'HTML sent to printer successfully!'})
        else:
            return jsonify({'status': 'error', 'message': 'Failed to print. Check printer connection.'})

    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Error: {str(e)}'})

def print_html_file(html_file_path, printer_name=''):
    """Print HTML file using system commands"""
    try:
        if os.name == 'nt':  # Windows
            # Method 1: Using Microsoft Edge (recommended for Windows)
            edge_path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
            if os.path.exists(edge_path):
                cmd = [
                    edge_path,
                    '--headless',
                    '--print-to-pdf',
                    '--print-to-pdf-no-header',
                    f'--print-to-pdf={html_file_path}.pdf',
                    html_file_path
                ]
                subprocess.run(cmd, check=True)

                # Print the PDF
                if printer_name:
                    pdf_print_cmd = [
                        'powershell',
                        '-Command',
                        f'Start-Process -FilePath "{html_file_path}.pdf" -Verb PrintTo -ArgumentList "{printer_name}"'
                    ]
                else:
                    pdf_print_cmd = [
                        'AcroRd32',
                        '/t',
                        f'{html_file_path}.pdf'
                    ]
                subprocess.run(pdf_print_cmd, check=True)
                return True

            # Method 2: Fallback - print as text
            else:
                cmd = ['notepad', '/p', html_file_path]
                if printer_name:
                    cmd.extend(['/pt', html_file_path, printer_name])
                subprocess.run(cmd, check=True)
                return True

        else:  # Linux/macOS
            # Convert HTML to PDF first for better printing quality
            try:
                # Try using wkhtmltopdf if available
                subprocess.run(['wkhtmltopdf', html_file_path, f'{html_file_path}.pdf'], check=True)
                print_cmd = ['lp']
                if printer_name:
                    print_cmd.extend(['-d', printer_name])
                print_cmd.append(f'{html_file_path}.pdf')
                subprocess.run(print_cmd, check=True)
            except:
                # Fallback: print directly as text
                print_cmd = ['lp']
                if printer_name:
                    print_cmd.extend(['-d', printer_name])
                print_cmd.append(html_file_path)
                subprocess.run(print_cmd, check=True)

            return True

    except subprocess.CalledProcessError:
        return False
    except Exception:
        return False

@app.route('/get-printers', methods=['GET'])
def get_printers():
    """Get list of available printers"""
    try:
        printers = []

        if os.name == 'nt':  # Windows
            try:
                import win32print
                printer_list = win32print.EnumPrinters(2)  # Get all printers
                printers = [printer[2] for printer in printer_list]
            except ImportError:
                # Fallback if win32print not available
                result = subprocess.run(['wmic', 'printer', 'get', 'name'],
                                      capture_output=True, text=True)
                if result.returncode == 0:
                    printers = [line.strip() for line in result.stdout.split('\n')
                               if line.strip() and not line.startswith('Name')]

        else:  # Linux/macOS
            result = subprocess.run(['lpstat', '-a'], capture_output=True, text=True)
            if result.returncode == 0:
                printers = [line.split()[0] for line in result.stdout.split('\n') if line]

        return jsonify({'status': 'success', 'printers': printer_list})

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
