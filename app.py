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
from swiggy_process import process_invoices_web
import swiggy_dineout_process
import zomato_pay_process
from zomato_consolidated_process import process_zomato_consolidated
import paytm_process

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['TEMPLATE_FILE'] = 'template.xlsx'  # ✅ Fixed Zomato template path
app.config['SWIGGY_TEMPLATE_FILE'] = 'template_files/recon_template.xlsx' # Swiggy template path
app.config['SWIGGY_DINEOUT_TEMPLATE'] = 'template_files/dineout_template.xlsx' # New Template
app.config['ZOMATO_PAY_TEMPLATE'] = 'template_files/zpay_template.xlsx' # Zomato Pay Template
app.config['PAYTM_TEMPLATE'] = 'template_files/paytm_template.xlsx' # Paytm Template

# Create folders if they don't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_formatted_filename(client_name, recon_type, month_name):
    """Format filename as 'Client - Recon Type Summary - Mon'YY.xlsx'"""
    client = str(client_name or "Unknown").strip()
    
    # Shorten month (e.g., January -> Jan)
    try:
        month_dt = datetime.strptime(month_name.strip().capitalize(), "%B")
        mon = month_dt.strftime("%b")
    except:
        mon = month_name[:3].capitalize()
    
    # Get current year last 2 digits
    year_short = datetime.now().strftime("%y")
    
    return f"{client} - {recon_type} Summary - {mon}'{year_short}.xlsx"


def cleanup_folder_delayed(folder_path, delay=3):
    """Cleanup folder after delay in background thread"""
    def cleanup():
        time.sleep(delay)
        try:
            if os.path.exists(folder_path):
                shutil.rmtree(folder_path)
                print(f"✅ Cleaned up session folder: {folder_path}")
            
            # Also clean up the progress file if possible
            # Get session_id from folder name
            session_id = os.path.basename(folder_path)
            # We don't necessarily know the task_id here, but we can search for it
            # if we standardized task_id to follow session_id or just clean up old ones.
            # For now, rely on cleanup_old_files for the .progress files if we can't link them easily.
        except Exception as e:
            print(f"⚠️  Could not cleanup {folder_path}: {e}")

    thread = threading.Thread(target=cleanup, daemon=True)
    thread.start()


# Task Progress Tracking
def update_progress(task_id, progress):
    """Update progress for a specific task using a temporary file"""
    if task_id:
        try:
            progress_file = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}.progress")
            with open(progress_file, 'w') as f:
                f.write(str(progress))
            print(f"Task {task_id} progress: {progress}%")
        except Exception as e:
            print(f"⚠️ Error updating progress file: {e}")

@app.route('/progress/<task_id>')
def get_progress(task_id):
    """Get current progress for a task from its progress file"""
    try:
        progress_file = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}.progress")
        if os.path.exists(progress_file):
            with open(progress_file, 'r') as f:
                progress = f.read().strip()
                return jsonify({'progress': int(progress)})
    except Exception as e:
        print(f"⚠️ Error reading progress file: {e}")
    
    return jsonify({'progress': 0})

@app.route('/')
def index():
    """Render main page"""
    # Check if template exists
    template_exists = os.path.exists(app.config['TEMPLATE_FILE'])
    return render_template('index.html', template_exists=template_exists)


@app.route('/upload/swiggy-dineout', methods=['POST'])
def upload_swiggy_dineout():
    """Handle Swiggy Dineout file upload"""
    try:
        if 'invoices' not in request.files:
            return jsonify({'success': False, 'message': 'No invoice files uploaded'})
            
        invoice_files = request.files.getlist('invoices')
        task_id = request.form.get('task_id')
        client_name = request.form.get('clientName', '')
        month = request.form.get('month', '')
        
        if not invoice_files or invoice_files[0].filename == '':
            return jsonify({'success': False, 'message': 'No invoice files selected'})
            
        # Optional: Save template if user provided one? 
        # For now assume static template path key
        
        # Define progress callback
        p_func = lambda p: update_progress(task_id, p)
        
        output_filename = get_formatted_filename(client_name, "Swiggy Dineout", month)
        
        output_file, error = swiggy_dineout_process.process_swiggy_dineout(
            invoice_files,
            app.config['SWIGGY_DINEOUT_TEMPLATE'],
            app.config['OUTPUT_FOLDER'],
            p_func,
            client_name=client_name,
            month=month,
            forced_filename=output_filename # Pass filename
        )
        
        if error:
            return jsonify({'success': False, 'message': f"Error: {error}"})
            
        download_url = f"/download/{output_file}"
        return jsonify({
            'success': True, 
            'message': 'Swiggy Dineout Reconciliation Completed!',
            'download_url': download_url
        })

    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/upload/zomato-pay', methods=['POST'])
def upload_zomato_pay():
    """Handle Zomato Pay file upload"""
    try:
        if 'invoices' not in request.files:
            return jsonify({'success': False, 'message': 'No invoice files uploaded'})
            
        invoice_files = request.files.getlist('invoices')
        task_id = request.form.get('task_id')
        client_name = request.form.get('clientName', '')
        month = request.form.get('month', '')

        # Week ranges
        f_start = request.form.get('firstWeekStart')
        f_end = request.form.get('firstWeekEnd')
        l_start = request.form.get('lastWeekStart')
        l_end = request.form.get('lastWeekEnd')
        
        if not invoice_files or invoice_files[0].filename == '':
            return jsonify({'success': False, 'message': 'No invoice files selected'})
            
        # Define progress callback
        p_func = lambda p: update_progress(task_id, p)
        
        output_filename = get_formatted_filename(client_name, "Zomato Pay", month)

        output_file, error = zomato_pay_process.process_zomato_pay(
            invoice_files,
            app.config['ZOMATO_PAY_TEMPLATE'],
            app.config['OUTPUT_FOLDER'],
            p_func,
            client_name=client_name,
            month=month,
            first_start=f_start,
            first_end=f_end,
            last_start=l_start,
            last_end=l_end,
            forced_filename=output_filename # Pass filename
        )
        
        if error:
            return jsonify({'success': False, 'message': f"Error: {error}"})
            
        download_url = f"/download/{output_file}"
        return jsonify({
            'success': True, 
            'message': 'Zomato Pay Reconciliation Completed!',
            'download_url': download_url
        })

    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


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
        recon_mode = request.form.get('recon_mode', 'weekly')

        # ✅ GET WEEK DATE RANGES (Only for weekly)
        first_week_start = request.form.get('first_week_start')
        first_week_end = request.form.get('first_week_end')
        last_week_start = request.form.get('last_week_start')
        last_week_end = request.form.get('last_week_end')

        # ✅ VALIDATE WEEK DATES (If weekly)
        if recon_mode == 'weekly' and not all([first_week_start, first_week_end, last_week_start, last_week_end]):
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
        output_filename = get_formatted_filename(client_name, "Zomato", month)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # Get Task ID for progress tracking
        task_id = request.form.get('task_id')
        if task_id:
            TASK_PROGRESS[task_id] = 5 # Initial progress
          # Run processing in background if many files, or synchronous if simple
        try:
            p_func = lambda p: update_progress(task_id, p)
            
            if recon_mode == 'consolidated':
                result = process_zomato_consolidated(
                    invoice_folder,
                    app.config['TEMPLATE_FILE'],
                    output_path,
                    client_name=client_name,
                    month=month,
                    first_week_start=first_week_start,
                    first_week_end=first_week_end,
                    last_week_start=last_week_start,
                    last_week_end=last_week_end,
                    progress_callback=p_func
                )
            else: # Default to weekly or other modes handled by process_zomato_recon
                result = process_zomato_recon(
                    invoice_folder,
                    app.config['TEMPLATE_FILE'],
                    output_path,
                    client_name=client_name,
                    month=month,
                    first_week_start=first_week_start,
                    first_week_end=first_week_end,
                    last_week_start=last_week_start,
                    last_week_end=last_week_end,
                    progress_callback=p_func
                )
        except Exception as e:
            # Re-raise or handle specific processing errors
            raise e


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


@app.route('/upload/swiggy', methods=['POST'])
def upload_swiggy_files():
    print("Swiggy Upload endpoint hit")
    session_folder = None

    try:
        # Check if template file exists
        if not os.path.exists(app.config['SWIGGY_TEMPLATE_FILE']):
            return jsonify({
                'success': False,
                'message': 'Swiggy Template file not found! Please contact administrator.'
            })

        if 'invoices' not in request.files:
            return jsonify({'success': False, 'message': 'No invoice files uploaded'})

        invoice_files = request.files.getlist('invoices')
        if not invoice_files or invoice_files[0].filename == '':
            return jsonify({'success': False, 'message': 'No invoice files selected'})

        bank_file = request.files.get('bankFile')

        client_name = request.form.get('clientName', '').strip()
        month = request.form.get('month', '').strip()

        try:
            first_week_start = int(request.form.get('firstWeekStart'))
            first_week_end = int(request.form.get('firstWeekEnd'))
            last_week_start = int(request.form.get('lastWeekStart'))
            last_week_end = int(request.form.get('lastWeekEnd'))
        except (ValueError, TypeError):
            return jsonify({'success': False, 'message': 'Invalid week range input'}), 400

        # Create unique session folder
        session_id = str(uuid.uuid4())[:8]
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], f"swiggy_{session_id}")
        os.makedirs(session_folder, exist_ok=True)

        # Save invoices
        saved_count = 0
        for f in invoice_files:
            if f and allowed_file(f.filename):
                filename = secure_filename(f.filename)
                f.save(os.path.join(session_folder, filename))
                saved_count += 1
        
        if saved_count == 0:
            shutil.rmtree(session_folder)
            return jsonify({'success': False, 'message': 'No valid invoice files uploaded'})

        # Save optional bank file
        bank_file_path = None
        if bank_file and bank_file.filename != '':
            if allowed_file(bank_file.filename):
                bank_filename = secure_filename(bank_file.filename)
                # Save just outside session folder or inside? Inside is cleaner for cleanup.
                # But process_invoices_web might expect it elsewhere? 
                # The original code saved it in UPLOAD_FOLDER with a unique name.
                # Let's save it in session_folder for easier cleanup.
                bank_file_path = os.path.join(session_folder, f"bank_{bank_filename}")
                bank_file.save(bank_file_path)
            else:
                shutil.rmtree(session_folder)
                return jsonify({'success': False, 'message': 'Invalid bank file format'})

        output_filename = get_formatted_filename(client_name, "Swiggy", month)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # Get Task ID for progress tracking
        task_id = request.form.get('task_id')
        if task_id:
            TASK_PROGRESS[task_id] = 5 # Initial progress

        result = process_invoices_web(
            invoice_folder_path=session_folder,
            template_recon_path=app.config['SWIGGY_TEMPLATE_FILE'],
            output_path=output_path,
            client_name=client_name,
            month=month,
            first_week_start=first_week_start,
            first_week_end=first_week_end,
            last_week_start=last_week_start,
            last_week_end=last_week_end,
            bank_file_path=bank_file_path,
            progress_callback=lambda p: update_progress(task_id, p)
        )

        # Cleanup
        gc.collect()
        if session_folder and os.path.exists(session_folder):
            cleanup_folder_delayed(session_folder, delay=2)

        if result['success']:
            return jsonify({
                'success': True,
                'message': result.get('message', 'Processed successfully'),
                'download_url': f"/download/{output_filename}"
            })
        else:
             return jsonify({
                'success': False,
                'message': result.get('message', 'Processing failed')
            })

    except Exception as e:
        import traceback
        traceback.print_exc()
        if session_folder and os.path.exists(session_folder):
             cleanup_folder_delayed(session_folder, delay=2)
        return jsonify({
            'success': False,
            'message': f'Error: {str(e)}'
        })



@app.route('/upload/paytm', methods=['POST'])
def upload_paytm():
    """Handle Paytm file upload and processing"""
    session_folder = None
    try:
        if not os.path.exists(app.config['PAYTM_TEMPLATE']):
            return jsonify({'success': False, 'message': 'Paytm template file not found!'})

        if 'invoices' not in request.files:
            return jsonify({'success': False, 'message': 'No file uploaded'})

        invoice_files = request.files.getlist('invoices')
        if not invoice_files or invoice_files[0].filename == '':
            return jsonify({'success': False, 'message': 'No file selected'})

        client_name = request.form.get('clientName', 'Client').strip()
        month = request.form.get('month', 'October')

        # Get week ranges
        first_week_start = request.form.get('firstWeekStart')
        first_week_end = request.form.get('firstWeekEnd')
        last_week_start = request.form.get('lastWeekStart')
        last_week_end = request.form.get('lastWeekEnd')

        session_id = str(uuid.uuid4())[:8]
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)

        # Save the first file (Paytm is expected as single file)
        file = invoice_files[0]
        filename = secure_filename(file.filename)
        filepath = os.path.join(session_folder, filename)
        file.save(filepath)

        output_filename = get_formatted_filename(client_name, "Paytm", month)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        task_id = request.form.get('task_id')
        p_func = lambda p: update_progress(task_id, p)
        if task_id: update_progress(task_id, 10)

        result = paytm_process.process_paytm(
            filepath,
            app.config['PAYTM_TEMPLATE'],
            output_path,
            client_name=client_name,
            month=month,
            first_week_start=first_week_start,
            first_week_end=first_week_end,
            last_week_start=last_week_start,
            last_week_end=last_week_end,
            progress_callback=p_func
        )

        if session_folder and os.path.exists(session_folder):
            cleanup_folder_delayed(session_folder, delay=2)

        if result['success']:
            return jsonify({
                'success': True,
                'message': 'Paytm Reconciliation Complete',
                'download_url': f"/download/{output_filename}"
            })
        else:
            return jsonify({'success': False, 'message': result.get('message', 'Processing failed')})

    except Exception as e:
        import traceback
        traceback.print_exc()
        if session_folder and os.path.exists(session_folder):
            cleanup_folder_delayed(session_folder, delay=2)
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})


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
            elif item.endswith('.progress'):
                age = now - os.path.getmtime(item_path)
                if age > 3600: # 1 hour
                    try:
                        os.remove(item_path)
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
