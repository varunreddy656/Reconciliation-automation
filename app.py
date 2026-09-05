"""
Zomato Reconciliation Tool - Flask Web Application
Fixed template version - users only upload invoices
"""

from flask import Flask, render_template, request, jsonify, send_file
import os
import shutil
from dotenv import load_dotenv

# Load environment variables IMMEDIATELY
load_dotenv()
print(f"DEBUG: GROQ_API_KEY loaded: {'Yes' if os.environ.get('GROQ_API_KEY') else 'No'}")
import tempfile
from werkzeug.utils import secure_filename
from datetime import datetime
import uuid
import time
import threading
import gc
import json
import base64
from groq import Groq
from duckduckgo_search import DDGS
from pos_cleaner import clean_pos_excel, aggregate_pos, generate_pos_report

def search_web(query):
    """Perform a web search using DuckDuckGo to get latest info."""
    try:
        with DDGS() as ddgs:
            # We search for the user's query and get top snippets
            results = list(ddgs.text(query, max_results=3))
            if not results:
                return "No search results found."
            
            formatted_results = []
            for r in results:
                formatted_results.append(f"Source: {r.get('href')}\nContent: {r.get('body')}")
            
            return "\n\n".join(formatted_results)
    except Exception as e:
        print(f"⚠️ Search error: {e}")
        return "Search failed."

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

# --- Chatbot Configuration & Rate Limiting ---
CHAT_LIMITS_FILE = os.path.join(app.config['UPLOAD_FOLDER'], 'chat_limits.json')
CHAT_MAX_MESSAGES = 4  # Max messages per minute
CHAT_RESET_INTERVAL = 60  # 1 minute in seconds

def load_chat_limits():
    if os.path.exists(CHAT_LIMITS_FILE):
        try:
            with open(CHAT_LIMITS_FILE, 'r') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_chat_limits(limits):
    try:
        with open(CHAT_LIMITS_FILE, 'w') as f:
            json.dump(limits, f)
    except:
        pass

chat_limits = load_chat_limits()

RESTRO_SYSTEM_PROMPT = """
You are the Restro AI Assistant, an expert in the entire spectrum of Indian Restaurant Accounting, Finance, and Operations. 
You represent 'Restro AI', a high-end tool for restaurant financial management.

STRICT OPERATING RULES:
1. SCOPE: Answer ALL questions related to restaurant accounting, including COGS (Cost of Goods Sold), Menu Engineering, Payroll, Vendor Management, Cashflow, GST/TDS, and software integration (Zoho, Tally, Petpooja).
2. RECONCILIATION: While you are a specialist in Zomato/Swiggy reconciliation, you are also a general authority on restaurant P&L and Balance Sheets.
3. ADVICE: Provide actionable advice on how to improve restaurant profitability, manage food costs, and resolve report discrepancies.
4. DISALLOWED TOPICS: If a query is completely unrelated to the restaurant industry or finance, respond with: "I am specialized in the restaurant business and accounting. I cannot assist with unrelated topics."
5. Be comprehensive yet concise. Use bullet points and professional formatting.
"""

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
            print(f"Task {task_id} progress: {progress}%", flush=True)
        except Exception as e:
            print(f"⚠️ Error updating progress file: {e}", flush=True)

@app.route('/progress/<task_id>')
def get_progress(task_id):
    """Get current progress for a task from its progress file"""
    try:
        progress_file = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}.progress")
        if os.path.exists(progress_file):
            with open(progress_file, 'r') as f:
                progress = f.read().strip()
                if progress:
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

        # Create unique session folder
        session_id = str(uuid.uuid4())[:8]
        session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)

        # ✅ VALIDATE WEEK DATES (If weekly)
        if recon_mode == 'weekly' and not all([first_week_start, first_week_end, last_week_start, last_week_end]):
            if os.path.exists(session_folder):
                shutil.rmtree(session_folder)
            return jsonify({
                'success': False,
                'message': 'All week date fields are required (First Week Start, First Week End, Last Week Start, Last Week End)'
            })

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

        # ✅ GET OPTIONAL BANK FILE
        bank_file = request.files.get('bankFile')
        bank_file_path = None
        if bank_file and bank_file.filename != '':
            if allowed_file(bank_file.filename):
                bank_filename = secure_filename(bank_file.filename)
                bank_file_path = os.path.join(session_folder, f"bank_{bank_filename}")
                bank_file.save(bank_file_path)
            else:
                if os.path.exists(session_folder):
                    shutil.rmtree(session_folder)
                return jsonify({'success': False, 'message': 'Invalid bank file format'})

        # Generate output path
        output_filename = get_formatted_filename(client_name, "Zomato", month)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # Get Task ID for progress tracking
        task_id = request.form.get('task_id')
        if task_id:
            update_progress(task_id, 5) # Initial progress
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
                    bank_file_path=bank_file_path,
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
                    bank_file_path=bank_file_path,
                    progress_callback=p_func
                )
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"❌ Processing Error:\n{error_details}")
            return jsonify({
                'success': False,
                'message': f"Processing Error: {str(e)}",
                'traceback': error_details
            }), 500


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
                'message': result.get('message', 'Processing failed'),
                'traceback': result.get('traceback', '')
            })

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"❌ Submission Error:\n{error_details}")

        # Cleanup on error
        try:
            if 'session_folder' in locals() and session_folder and os.path.exists(session_folder):
                time.sleep(1)
                shutil.rmtree(session_folder)
        except:
            pass

        return jsonify({
            'success': False,
            'message': f"Submission Error: {str(e)}",
            'traceback': error_details
        }), 500

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
            update_progress(task_id, 5) # Initial progress

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
        error_details = traceback.format_exc()
        print(f"❌ Swiggy Error:\n{error_details}")
        return jsonify({
            'success': False,
            'message': f"Submission Error: {str(e)}",
            'traceback': error_details
        }), 500



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
        error_details = traceback.format_exc()
        print(f"❌ Paytm Error:\n{error_details}")
        return jsonify({
            'success': False,
            'message': f"Submission Error: {str(e)}",
            'traceback': error_details
        }), 500


@app.route('/parse-pos-structure', methods=['POST'])
def parse_pos_structure():
    """Uses Groq Llama-3.2 Vision to read irregular date ranges from a screenshot."""
    if 'image' not in request.files:
        return jsonify({"error": "No image uploaded"}), 400
    
    api_key = os.environ.get('GROQ_API_KEY')
    if not api_key:
        return jsonify({"error": "GROQ_API_KEY not configured in environment"}), 500
    
    client = Groq(api_key=api_key)
    file = request.files['image']
    
    try:
        # Read image bytes and encode to base64
        img_bytes = file.read()
        base64_image = base64.b64encode(img_bytes).decode('utf-8')
        
        prompt = """
        Analyze this image of a table header containing date ranges.
        Extract every period mentioned (e.g., '1st to 3rd', '4th to 7th').
        Return ONLY a clean JSON array of objects:
        [
          {"label": "1st to 3rd", "start_day": 1, "end_day": 3},
          ...
        ]
        Extract all columns from left to right. Ensure start_day and end_day are integers.
        Return only the JSON array, no markdown or explanation.
        """
        
        completion = client.chat.completions.create(
            model="meta-llama/llama-4-scout-17b-16e-instruct",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}",
                            },
                        },
                    ],
                }
            ],
            temperature=0.1,
            max_tokens=1024,
        )
        
        # Parse response
        resp_text = completion.choices[0].message.content
        clean_json = resp_text.replace('```json', '').replace('```', '').strip()
        ranges = json.loads(clean_json)
        
        return jsonify({"ranges": ranges})
        
    except Exception as e:
        print(f"❌ Groq AI Parsing Error: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/extract-pos', methods=['POST'])
def extract_pos():
    """Manual POS extraction and aggregation with Excel generation"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': 'No file uploaded'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'message': 'No file selected'})

        # 1. Clean Data
        week_ranges = request.form.get('week_ranges')
        if week_ranges:
            week_ranges = json.loads(week_ranges)
            
        custom_dinein_ranges = request.form.get('custom_dinein_ranges')
        if custom_dinein_ranges:
            custom_dinein_ranges = json.loads(custom_dinein_ranges)

        clean_result = clean_pos_excel(file)
        pos_type = clean_result['pos_type']
        cleaned_df = clean_result['dataframe']
        
        if cleaned_df.empty:
            return jsonify({'success': False, 'message': 'Could not extract any data from Excel'})

        # 3. Hardcoded Aggregation
        data = aggregate_pos(cleaned_df, pos_type, week_ranges, custom_dinein_ranges)
        
        # 3. Generate Report
        report_filename = f"POS_Summary_{uuid.uuid4().hex[:8]}.xlsx"
        report_path = os.path.join(app.config['OUTPUT_FOLDER'], report_filename)
        generate_pos_report(data, report_path)
        
        return jsonify({
            'success': True,
            'pos_type': pos_type,
            'data': data,
            'download_url': f"/download/{report_filename}"
        })

    except Exception as e:
        print(f"❌ POS Extraction Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})



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


@app.route('/chat', methods=['POST'])
def chat():
    """Restro AI Chatbot with Knowledge Base and Web Search"""
    data = request.json
    user_message = data.get('message', '').strip()
    user_id = request.remote_addr
    
    if not user_message:
        return jsonify({"error": "Empty message"}), 400
    
    if len(user_message) > 500:
        return jsonify({"error": "Message too long. Please keep questions under 500 characters."}), 400
    
    # Rate Limiting Logic
    now = time.time()
    if user_id not in chat_limits:
        chat_limits[user_id] = {'count': 0, 'last_reset': now}
    
    if now - chat_limits[user_id]['last_reset'] > CHAT_RESET_INTERVAL:
        chat_limits[user_id] = {'count': 0, 'last_reset': now}
    
    if chat_limits[user_id]['count'] >= CHAT_MAX_MESSAGES:
        return jsonify({
            "reply": "You are sending messages too quickly. Please wait a minute before asking more questions.",
            "error": "Rate limit exceeded"
        }), 429
    
    # Load Internal Knowledge Base
    kb_content = ""
    kb_path = 'restro_knowledge.txt'
    if os.path.exists(kb_path):
        try:
            with open(kb_path, 'r', encoding='utf-8') as f:
                kb_content = f.read()
        except:
            pass

    # Determine if we need Web Search (expanded to all restaurant accounting)
    accounting_keywords = [
        'zoho', 'tally', 'petpooja', 'pet pooja', 'gst', 'tds', 'payroll', 
        'cogs', 'food cost', 'inventory', 'vendor', 'p&l', 'profit and loss',
        'balance sheet', 'cashflow', 'swiggy', 'zomato', 'dineout', 'accounting',
        'reconciliation', 'margin', 'ebitda', 'operating cost'
    ]
    web_context = ""
    should_search = any(kw in user_message.lower() for kw in accounting_keywords) or len(user_message.split()) > 6
    
    if should_search:
        print(f"🔍 Performing Web Search for: {user_message}")
        web_context = search_web(user_message)

    # Build Final Prompt
    enhanced_system_prompt = f"{RESTRO_SYSTEM_PROMPT}\n\n"
    
    if kb_content:
        enhanced_system_prompt += f"INTERNAL COMPANY KNOWLEDGE (Prioritize this):\n{kb_content}\n\n"
        
    if web_context:
        enhanced_system_prompt += f"LIVE WEB SEARCH CONTEXT (Use for software specifics):\n{web_context}\n\n"

    # Process with Groq
    api_key = os.environ.get('GROQ_API_KEY')
    if not api_key:
        return jsonify({"error": "AI configuration missing"}), 500
    
    try:
        client = Groq(api_key=api_key)
        completion = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": enhanced_system_prompt},
                {"role": "user", "content": user_message}
            ],
            temperature=0.3,
            max_tokens=300
        )
        
        chat_limits[user_id]['count'] += 1
        save_chat_limits(chat_limits)
        
        return jsonify({
            "reply": completion.choices[0].message.content,
            "remaining": CHAT_MAX_MESSAGES - chat_limits[user_id]['count']
        })
        
    except Exception as e:
        print(f"? Chatbot Error: {e}")
        return jsonify({"error": "Restro AI Assistant is currently resting. Try again in a bit."}), 500


if __name__ == '__main__':
    # Check if template exists on startup
    if not os.path.exists(app.config['TEMPLATE_FILE']):
        print("⚠️  WARNING: template.xlsx not found!")
        print("📋 Please place your template.xlsx file in the root directory")
    else:
        print("✅ Template file found!")

    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)
