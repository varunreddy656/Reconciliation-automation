"""
Zomato Reconciliation Tool - Flask Web Application
Fixed template version - users only upload invoices
"""

from flask import Flask, render_template, request, jsonify, send_file
import os
import shutil
import tempfile
from werkzeug.utils import secure_filename
from datetime import datetime
import uuid
import time
import threading
import gc

# Import backend processing
from process_invoices import process_zomato_recon

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['TEMPLATE_FILE'] = 'template.xlsx'  # ✅ Fixed template path

# Create folders if they don't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def cleanup_folder_delayed(folder_path, delay=3):
    """Cleanup folder after delay in background thread"""
    def cleanup():
        time.sleep(delay)
        try:
            if os.path.exists(folder_path):
                shutil.rmtree(folder_path)
                print(f"✅ Cleaned up: {folder_path}")
        except Exception as e:
            print(f"⚠️  Could not cleanup {folder_path}: {e}")

    thread = threading.Thread(target=cleanup, daemon=True)
    thread.start()


@app.route('/')
def index():
    """Render main page"""
    # Check if template exists
    template_exists = os.path.exists(app.config['TEMPLATE_FILE'])
    return render_template('index.html', template_exists=template_exists)


@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file upload and processing"""
    session_folder = None

    try:
        # Check if template file exists
        if not os.path.exists(app.config['TEMPLATE_FILE']):
            return jsonify({
                'success': False,
                'message': 'Template file not found! Please contact administrator.'
            })

        # Validate invoice files
        if 'invoices' not in request.files:
            return jsonify({'success': False, 'message': 'No invoice files uploaded'})

        invoice_files = request.files.getlist('invoices')

        if not invoice_files or invoice_files[0].filename == '':
            return jsonify({'success': False, 'message': 'No invoice files selected'})

        # Get form data
        month = request.form.get('month', 'October')
        client_name = request.form.get('client_name', '').strip() or None

        # ✅ GET WEEK DATE RANGES
        first_week_start = request.form.get('first_week_start')
        first_week_end = request.form.get('first_week_end')
        last_week_start = request.form.get('last_week_start')
        last_week_end = request.form.get('last_week_end')

        # ✅ VALIDATE WEEK DATES
        if not all([first_week_start, first_week_end, last_week_start, last_week_end]):
            if session_folder and os.path.exists(session_folder):
                shutil.rmtree(session_folder)
            return jsonify({
                'success': False,
                'message': 'All week date fields are required (First Week Start, First Week End, Last Week Start, Last Week End)'
            })

        # Create unique session folder
        session_id = str(uuid.uuid4())[:8]
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)

        # Save invoice files
        invoice_folder = os.path.join(session_folder, 'invoices')
        os.makedirs(invoice_folder, exist_ok=True)

        saved_invoices = []
        for invoice in invoice_files:
            if invoice and allowed_file(invoice.filename):
                filename = secure_filename(invoice.filename)
                filepath = os.path.join(invoice_folder, filename)
                invoice.save(filepath)
                saved_invoices.append(filepath)

        if not saved_invoices:
            if session_folder and os.path.exists(session_folder):
                shutil.rmtree(session_folder)
            return jsonify({'success': False, 'message': 'No valid invoice files uploaded'})

        # Generate output path
        output_filename = f"Zomato_Recon_{month}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # ✅ Process reconciliation using FIXED template
        result = process_zomato_recon(
            invoice_folder_path=invoice_folder,
            template_recon_path=app.config['TEMPLATE_FILE'],
            output_path=output_path,
            client_name=client_name,
            month=month,
            first_week_start=first_week_start,
            first_week_end=first_week_end,
            last_week_start=last_week_start,
            last_week_end=last_week_end
        )


        # ✅ Force garbage collection to release file handles
        gc.collect()
        time.sleep(0.5)  # Small delay to ensure handles are released

        # ✅ Cleanup session folder in BACKGROUND (delayed)
        if session_folder and os.path.exists(session_folder):
            cleanup_folder_delayed(session_folder, delay=2)

        if result.get('success'):
            return jsonify({
                'success': True,
                'message': f"Successfully processed {result['weeks_processed']} weeks",
                'download_url': f"/download/{output_filename}",
                'weeks_processed': result['weeks_processed']
            })
        else:
            return jsonify({
                'success': False,
                'message': result.get('message', 'Processing failed')
            })

    except Exception as e:
        import traceback
        traceback.print_exc()

        # Cleanup on error
        if session_folder and os.path.exists(session_folder):
            try:
                time.sleep(1)
                shutil.rmtree(session_folder)
            except:
                cleanup_folder_delayed(session_folder, delay=2)

        return jsonify({
            'success': False,
            'message': f'Error: {str(e)}'
        })


@app.route('/download/<filename>')
def download_file(filename):
    """Download processed file"""
    try:
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True)
        else:
            return "File not found", 404
    except Exception as e:
        return f"Error: {str(e)}", 500


@app.route('/cleanup', methods=['POST'])
def cleanup_old_files():
    """Cleanup old files (optional maintenance endpoint)"""
    try:
        cleaned = 0
        now = time.time()

        # Clean uploads older than 1 hour
        for item in os.listdir(app.config['UPLOAD_FOLDER']):
            item_path = os.path.join(app.config['UPLOAD_FOLDER'], item)
            if os.path.isdir(item_path):
                age = now - os.path.getmtime(item_path)
                if age > 3600:  # 1 hour
                    try:
                        shutil.rmtree(item_path)
                        cleaned += 1
                    except:
                        pass

        # Clean outputs older than 24 hours
        for filename in os.listdir(app.config['OUTPUT_FOLDER']):
            filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
            if os.path.isfile(filepath):
                age = now - os.path.getmtime(filepath)
                if age > 86400:  # 24 hours
                    try:
                        os.remove(filepath)
                        cleaned += 1
                    except:
                        pass

        return jsonify({'success': True, 'message': f'Cleaned {cleaned} items'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


if __name__ == '__main__':
    # Check if template exists on startup
    if not os.path.exists(app.config['TEMPLATE_FILE']):
        print("⚠️  WARNING: template.xlsx not found!")
        print("📋 Please place your template.xlsx file in the root directory")
    else:
        print("✅ Template file found!")

    app.run(debug=True, host='0.0.0.0', port=5000)
