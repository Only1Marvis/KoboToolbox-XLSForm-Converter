# ── Imports and Setup ─────────────────────────────────────────────────────────
import os
import uuid
from flask import (
    Flask, request, render_template_string,
    send_file, jsonify
)
from convert import convert

# ── Create the Flask app ──────────────────────────────────────────────────────
# Flask is the web framework that turns our Python script into a website.
# __name__ tells Flask where to find files relative to this script.
app = Flask(__name__)

# ── Folder configuration ──────────────────────────────────────────────────────
# UPLOAD_FOLDER is where we temporarily save Word files the user uploads.
# OUTPUT_FOLDER is where we save the converted XLSForm Excel files.
# We create both folders if they don't already exist.
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# ── Allowed file types ────────────────────────────────────────────────────────
# We only accept Word documents (.docx).
# This prevents users from accidentally uploading wrong file types.
ALLOWED_EXTENSIONS = {'docx'}


def allowed_file(filename):
    """
    Checks if the uploaded file has an allowed extension.
    Returns True if the file ends with .docx, False otherwise.
    '.' in filename ensures the file has an extension at all.
    """
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ── HTML Template ─────────────────────────────────────────────────────────────
# This is the complete web page users will see in their browser.
# We store it as a Python string using triple quotes.
# Flask's render_template_string() will serve this as a real web page.
HTML_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KoboToolbox XLSForm Converter</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            background: #f0f2f5;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 2rem;
        }

        .container {
            background: white;
            border-radius: 16px;
            padding: 2.5rem;
            width: 100%;
            max-width: 860px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
        }

        /* ── Side by side layout ── */
        .main-layout {
            display: flex;
            gap: 2rem;
            align-items: flex-start;
            margin-bottom: 1.5rem;
            justify-content: center;
        }

        .left-panel {
            flex: 1;
            min-width: 280px;
            max-width: 420px;
        }

        .right-panel {
            flex-shrink: 0;
            width: auto;
        }

        .logo {
            text-align: center;
            margin-bottom: 2rem;
        }

        .logo h1 {
            font-size: 1.6rem;
            font-weight: 700;
            color: #1a1a2e;
            margin-bottom: 0.4rem;
        }

        .logo p {
            font-size: 0.9rem;
            color: #6b7280;
        }

        .upload-area {
            border: 2px dashed #d1d5db;
            border-radius: 12px;
            padding: 2.5rem 1.5rem;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s ease;
            margin-bottom: 1.5rem;
            background: #fafafa;
        }

        .upload-area:hover {
            border-color: #4472C4;
            background: #f0f4ff;
        }

        .upload-area.dragover {
            border-color: #4472C4;
            background: #e8eeff;
        }

        .upload-icon {
            font-size: 3rem;
            margin-bottom: 1rem;
        }

        .upload-area h3 {
            font-size: 1rem;
            font-weight: 600;
            color: #374151;
            margin-bottom: 0.4rem;
        }

        .upload-area p {
            font-size: 0.85rem;
            color: #9ca3af;
        }

        .file-input {
            display: none;
        }

        .selected-file {
            display: none;
            background: #f0f4ff;
            border: 1px solid #c7d2fe;
            border-radius: 8px;
            padding: 0.75rem 1rem;
            margin-bottom: 1.5rem;
            font-size: 0.875rem;
            color: #3730a3;
            align-items: center;
            gap: 0.5rem;
        }

        .selected-file.visible {
            display: flex;
        }

        .convert-btn {
            width: 100%;
            padding: 0.875rem;
            background: #4472C4;
            color: white;
            border: none;
            border-radius: 10px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.2s ease;
        }

        .convert-btn:hover {
            background: #3461b3;
        }

        .convert-btn:disabled {
            background: #9ca3af;
            cursor: not-allowed;
        }

        .status-box {
            display: none;
            margin-top: 1.5rem;
            padding: 1rem 1.25rem;
            border-radius: 10px;
            font-size: 0.9rem;
            text-align: center;
        }

        .status-box.loading {
            display: block;
            background: #fef9c3;
            color: #854d0e;
            border: 1px solid #fde68a;
        }

        .status-box.success {
            display: block;
            background: #f0fdf4;
            color: #166534;
            border: 1px solid #bbf7d0;
        }

        .status-box.error {
            display: block;
            background: #fef2f2;
            color: #991b1b;
            border: 1px solid #fecaca;
        }

        .download-btn {
            display: inline-block;
            margin-top: 0.75rem;
            padding: 0.6rem 1.5rem;
            background: #166534;
            color: white;
            border-radius: 8px;
            text-decoration: none;
            font-size: 0.875rem;
            font-weight: 600;
            transition: background 0.2s ease;
        }

        .download-btn:hover {
            background: #14532d;
        }

        .rules-note {
            margin-top: 2rem;
            padding-top: 1.5rem;
            border-top: 1px solid #f3f4f6;
            font-size: 0.8rem;
            color: #9ca3af;
            text-align: center;
            line-height: 1.6;
        }

        /* ── Repeat format selector ── */
        .repeat-section {
            margin-bottom: 0;
        }

        .repeat-section label.section-label {
            display: block;
            font-size: 0.82rem;
            font-weight: 600;
            color: #374151;
            margin-bottom: 0.4rem;
        }

        .repeat-options {
            display: flex;
            flex-direction: column;
            gap: 0.3rem;
            align-items: stretch;
            width: fit-content;
        }

        .repeat-option {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            padding: 0.5rem 0.75rem;
            border: 1px solid #e5e7eb;
            border-radius: 7px;
            cursor: pointer;
            transition: all 0.15s ease;
            position: relative;
        }

        .repeat-option .option-text {
            flex: 0 0 auto;
        }

        .repeat-option:hover {
            border-color: #4472C4;
            background: #f8faff;
        }

        .repeat-option input[type='radio'] {
            accent-color: #4472C4;
            flex-shrink: 0;
            margin: 0;
        }

        .repeat-option .option-title {
            font-size: 0.8rem;
            font-weight: 600;
            color: #1f2937;
            line-height: 1.2;
        }

        .repeat-option .badge-default {
            font-size: 0.62rem;
            background: #4472C4;
            color: white;
            padding: 1px 5px;
            border-radius: 10px;
            font-weight: 600;
            margin-left: 4px;
            vertical-align: middle;
        }

        .repeat-option:has(input:checked) {
            border-color: #4472C4;
            background: #f0f4ff;
        }

        .repeat-option:has(input:checked) .option-title {
            color: #4472C4;
        }

        /* ── Output Options section ── */
        .output-options {
            margin-top: 1.2rem;
            padding: 1rem 1.25rem;
            background: #f8faff;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
        }

        .output-options label.section-label {
            display: block;
            font-size: 0.82rem;
            font-weight: 600;
            color: #374151;
            margin-bottom: 0.6rem;
        }

        .output-option-row {
            display: inline-flex;
            align-items: center;
            gap: 0.6rem;
            position: relative;
            cursor: pointer;
            padding: 0.4rem 0.6rem;
            border: 1px solid #e5e7eb;
            border-radius: 7px;
            transition: all 0.15s ease;
        }

        .output-option-row:hover {
            border-color: #4472C4;
            background: #f0f4ff;
        }

        .output-option-row input[type='checkbox'] {
            accent-color: #4472C4;
            flex-shrink: 0;
        }

        .output-option-row .opt-text {
            font-size: 0.81rem;
            font-weight: 600;
            color: #1f2937;
            white-space: nowrap;
        }

        .output-option-row .opt-tooltip {
            visibility: hidden;
            opacity: 0;
            width: 280px;
            white-space: normal;
            word-wrap: break-word;
            background: #1f2937;
            color: #fff;
            font-size: 0.72rem;
            line-height: 1.5;
            border-radius: 6px;
            padding: 8px 11px;
            position: absolute;
            z-index: 100;
            bottom: calc(100% + 10px);
            left: 0;
            transition: opacity 0.2s;
            pointer-events: none;
            box-shadow: 0 4px 14px rgba(0,0,0,0.18);
        }

        .output-option-row .opt-tooltip::after {
            content: '';
            position: absolute;
            top: 100%;
            left: 20px;
            border: 6px solid transparent;
            border-top-color: #1f2937;
        }

        .output-option-row:hover .opt-tooltip {
            visibility: visible;
            opacity: 1;
        }
        .repeat-option .tooltip-text {
            visibility: hidden;
            opacity: 0;
            width: 240px;
            white-space: normal;
            word-wrap: break-word;
            background: #1f2937;
            color: #fff;
            font-size: 0.72rem;
            line-height: 1.5;
            border-radius: 6px;
            padding: 8px 11px;
            position: absolute;
            z-index: 100;
            bottom: calc(100% + 10px);
            left: 50%;
            transform: translateX(-50%);
            transition: opacity 0.2s;
            pointer-events: none;
            box-shadow: 0 4px 14px rgba(0,0,0,0.18);
        }

        /* Arrow pointing downward toward the option */
        .repeat-option .tooltip-text::after {
            content: '';
            position: absolute;
            top: 100%;
            left: 50%;
            transform: translateX(-50%);
            border: 6px solid transparent;
            border-top-color: #1f2937;
        }

        .repeat-option:hover .tooltip-text {
            visibility: visible;
            opacity: 1;
        }

        /* ── Spinner animation for loading state ── */
        @keyframes spin {
            0%   { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .spinner {
            display: inline-block;
            width: 16px;
            height: 16px;
            border: 2px solid #854d0e;
            border-top-color: transparent;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            vertical-align: middle;
            margin-right: 8px;
        }
    </style>
</head>
<body>
    <div class="container">

        <!-- Header -->
        <div class="logo">
            <h1>KoboToolbox XLSForm Converter</h1>
            <p>Upload a Word questionnaire and get a ready-to-use XLSForm</p>
        </div>

        <!-- Side by side layout -->
        <div class="main-layout">

            <!-- Left panel: upload + convert -->
            <div class="left-panel">

                <!-- Upload Area -->
                <div class="upload-area" id="uploadArea">
                    <div class="upload-icon">📄</div>
                    <h3>Drop your Word document here</h3>
                    <p>or click to browse — .docx files only</p>
                    <input
                        type="file"
                        id="fileInput"
                        class="file-input"
                        accept=".docx"
                    />
                </div>

                <!-- Selected File Display -->
                <div class="selected-file" id="selectedFile">
                    <span>📎</span>
                    <span id="fileName">No file selected</span>
                </div>

                <!-- Convert Button — sits directly under upload -->
                <button
                    class="convert-btn"
                    id="convertBtn"
                    disabled
                >
                    Convert to XLSForm
                </button>

            </div>

            <!-- Right panel: repeat format -->
            <div class="right-panel">
                <div class="repeat-section">
                    <label class="section-label">
                        🔁 Repeat Group Format
                    </label>
                    <div class="repeat-options">

                        <label class="repeat-option">
                            <input type="radio" name="repeatFormat" value="3" />
                            <div class="option-text">
                                <div class="option-title">
                                    Automated Repeat Loop
                                </div>
                            </div>
                            <span class="tooltip-text">
                                True begin_repeat driven by a prior select_multiple.
                                Adapts automatically per respondent — the number of
                                iterations matches how many items were selected.
                                Best for large or variable lists.
                            </span>
                        </label>

                        <label class="repeat-option">
                            <input type="radio" name="repeatFormat" value="2" />
                            <div class="option-text">
                                <div class="option-title">Brand-Fixed Group Loop</div>
                            </div>
                            <span class="tooltip-text">
                                Each brand gets its own permanently fixed group
                                that appears conditionally when selected.
                                Best for brand equity tracking, retail audits
                                and competitive benchmarking.
                            </span>
                        </label>

                        <label class="repeat-option">
                            <input type="radio" name="repeatFormat" value="1" />
                            <div class="option-text">
                                <div class="option-title">Sequential Positional Loop</div>
                            </div>
                            <span class="tooltip-text">
                                Fixed unrolled slot groups — the enumerator picks
                                one brand per slot. Already-selected brands are
                                excluded from subsequent slots automatically.
                                Best for product testing, sensory evaluation
                                and ranked preference studies.
                            </span>
                        </label>

                        <label class="repeat-option">
                            <input type="radio" name="repeatFormat" value="4" checked />
                            <div class="option-text">
                                <div class="option-title">
                                    🧠 Smart Auto-Select
                                    <span class="badge-default">Default</span>
                                </div>
                            </div>
                            <span class="tooltip-text">
                                The engine evaluates each repeat group individually —
                                considering the number of options, how the study is
                                structured, and what works best for the raw data output.
                                Different groups in the same questionnaire may use
                                different formats.
                            </span>
                        </label>

                        <label class="repeat-option">
                            <input type="radio" name="repeatFormat" value="5" />
                            <div class="option-text">
                                <div class="option-title">✍️ Direct Scripting Mode</div>
                            </div>
                            <span class="tooltip-text">
                                No system repeat groups are applied. Questions are
                                scripted exactly as written in the questionnaire —
                                group by group, question by question, with no
                                automated repeat wrapping.
                            </span>
                        </label>

                    </div>
                </div>
            </div>

        </div>
        <!-- end main-layout -->

        <!-- Output Options -->
        <div class="output-options">
            <label class="section-label">⚙️ Output Options</label>

            <div class="output-option-row">
                <input type="checkbox" id="commaDelimiter" name="commaDelimiter" checked />
                <div class="opt-text">
                    Use comma (,) delimiter for multiple choice labels
                </div>
                <span class="opt-tooltip">
                    Recommended when option labels contain multiple words.
                    Separates selected labels with commas instead of spaces
                    for cleaner data analysis. Supported on KoboToolbox,
                    ODK Collect (v1.30+) and SurveyCTO. On older or
                    unsupported platforms this setting is safely ignored —
                    the form will still run normally with no errors.
                </span>
            </div>

        </div>

        <!-- Status Messages -->
        <div class="status-box" id="statusBox"></div>

        <!-- Footer Note -->
        <div class="rules-note">
            Applies all KoboToolbox XLSForm rules automatically —
            sections, groups, grids, SO questions, choice filters and more.
        </div>

        <!-- Credit line -->
        <div style="
            margin-top: 1.5rem;
            padding-top: 1rem;
            border-top: 1px solid #f3f4f6;
            text-align: center;
            font-size: 0.75rem;
            color: #c4c4c4;
            font-style: italic;
            letter-spacing: 0.03em;
        ">
            By Marvis Onyenwenu Enubiaka
        </div>

    </div>

    <script>
        // ── Get references to all the HTML elements we need ──────────────────
        const uploadArea   = document.getElementById('uploadArea');
        const fileInput    = document.getElementById('fileInput');
        const convertBtn   = document.getElementById('convertBtn');
        const selectedFile = document.getElementById('selectedFile');
        const fileName     = document.getElementById('fileName');
        const statusBox    = document.getElementById('statusBox');

        // ── Open file browser when upload area is clicked ─────────────────────
        uploadArea.addEventListener('click', () => fileInput.click());

        // ── Handle drag and drop ──────────────────────────────────────────────
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) handleFileSelect(file);
        });

        // ── Handle file selection from browser dialog ─────────────────────────
        fileInput.addEventListener('change', () => {
            const file = fileInput.files[0];
            if (file) handleFileSelect(file);
        });

        // ── Show selected file name and enable convert button ─────────────────
        function handleFileSelect(file) {
            if (!file.name.endsWith('.docx')) {
                showStatus('error', 'Please upload a .docx Word document only.');
                return;
            }
            fileName.textContent = file.name;
            selectedFile.classList.add('visible');
            convertBtn.disabled = false;
            statusBox.style.display = 'none';
        }

        // ── Handle convert button click ───────────────────────────────────────
        convertBtn.addEventListener('click', async () => {
            const file = fileInput.files[0] || null;
            if (!file) return;

            // Get selected repeat format
            const repeatFormatEl = document.querySelector(
                'input[name="repeatFormat"]:checked'
            );
            const repeatFormat = repeatFormatEl ? repeatFormatEl.value : '4';

            // Get comma delimiter option
            const commaDelimiter = document.getElementById('commaDelimiter').checked;

            // Show loading state immediately
            convertBtn.disabled = true;
            showStatus('loading',
                '<span class="spinner"></span> Converting your questionnaire... please wait.'
            );

            // Small delay to guarantee the loading state renders visibly
            await new Promise(resolve => setTimeout(resolve, 80));

            // Build form data to send the file and options
            const formData = new FormData();
            formData.append('file', file);
            formData.append('repeat_format', repeatFormat);
            formData.append('comma_delimiter', commaDelimiter ? '1' : '0');

            try {
                const response = await fetch('/convert', {
                    method: 'POST',
                    body: formData
                });

                const result = await response.json();

                if (result.success) {
                    let html = '✅ Conversion successful! Your XLSForm is ready.' +
                        '<br><a class="download-btn" href="/download/' +
                        result.filename + '">⬇ Download XLSForm</a>';

                    if (result.errors && result.errors.length > 0) {
                        html += '<div style="margin-top:1rem; padding:0.75rem 1rem;' +
                            'background:#fef9c3; border:1px solid #fde68a;' +
                            'border-radius:8px; text-align:left;">' +
                            '<strong style="color:#854d0e;">⚠ ' +
                            result.errors.length +
                            ' issue(s) found — please review:</strong>' +
                            '<ol style="margin:0.5rem 0 0.5rem 1.2rem;' +
                            'color:#854d0e; font-size:0.82rem; line-height:1.7;">';
                        result.errors.forEach(err => {
                            html += '<li>' + err + '</li>';
                        });
                        html += '</ol>' +
                            '<p style="margin:0.5rem 0 0; font-size:0.82rem;' +
                            'color:#854d0e; font-style:italic;">' +
                            'Kindly effect these corrections on the Word document ' +
                            'questionnaire and re-upload.</p>' +
                            '</div>';
                    }

                    showStatus('success', html);
                } else {
                    showStatus('error', '❌ Error: ' + result.error);
                }

            } catch (err) {
                showStatus('error',
                    '❌ Something went wrong. Please try again.'
                );
            }

            convertBtn.disabled = false;
        });

        // ── Helper: show a status message ─────────────────────────────────────
        // We explicitly set display to 'block' here because the CSS hides
        // the status box by default (display: none). Setting the class alone
        // is not enough — we must also force it visible with display: block.
        function showStatus(type, message) {
            statusBox.className    = 'status-box ' + type;
            statusBox.style.display = 'block';
            statusBox.innerHTML    = message;
        }
    </script>

</body>
</html>
"""


# ── ROUTE 1: Home Page ────────────────────────────────────────────────────────
# This is the route that serves the main page of the web app.
# When a user visits http://localhost:5000 in their browser,
# Flask runs this function and returns the HTML page we built above.
@app.route('/')
def index():
    return render_template_string(HTML_PAGE)


# ── ROUTE 2: Convert ──────────────────────────────────────────────────────────
# This route handles the file upload and conversion.
# It only accepts POST requests — meaning it only runs when the
# user clicks the Convert button and the browser sends the file.
@app.route('/convert', methods=['POST'])
def convert_file():

    # ── Check a file was actually sent ───────────────────────────────────────
    if 'file' not in request.files:
        return jsonify({
            'success': False,
            'error': 'No file was uploaded. Please select a .docx file.'
        })

    file = request.files['file']

    # ── Check the file has a name ─────────────────────────────────────────────
    if file.filename == '':
        return jsonify({
            'success': False,
            'error': 'No file selected. Please choose a .docx file.'
        })

    # ── Check the file is an allowed type ─────────────────────────────────────
    if not allowed_file(file.filename):
        return jsonify({
            'success': False,
            'error': 'Invalid file type. Please upload a .docx Word document.'
        })

    try:
        # ── Generate a unique ID for this conversion ──────────────────────────
        # uuid4() generates a random unique ID like:
        # '3f2504e0-4f89-11d3-9a0c-0305e82c3301'
        # We use only the first 8 characters to keep filenames short.
        unique_id = str(uuid.uuid4())[:8]

        # ── Save the uploaded Word file temporarily ───────────────────────────
        # We give it a unique name to avoid conflicts with other uploads.
        original_name  = os.path.splitext(file.filename)[0]
        input_filename = f'{original_name}_{unique_id}.docx'
        input_path     = os.path.join(app.config['UPLOAD_FOLDER'],
                                      input_filename)
        file.save(input_path)

        # ── Set the output Excel file path ────────────────────────────────────
        output_filename = f'{original_name}_{unique_id}_XLSForm.xlsx'
        output_path     = os.path.join(app.config['OUTPUT_FOLDER'],
                                       output_filename)

        # ── Run the conversion ────────────────────────────────────────────────
        repeat_format   = int(request.form.get('repeat_format', 4))
        comma_delimiter = request.form.get('comma_delimiter', '1') == '1'
        output_path, errors = convert(
            filepath        = input_path,
            output_path     = output_path,
            use_ai          = False,
            repeat_format   = repeat_format,
            comma_delimiter = comma_delimiter
        )

        # ── Clean up the uploaded Word file ───────────────────────────────────
        # We delete the input file after conversion since we no longer need it.
        # The output Excel file stays until the user downloads it.
        os.remove(input_path)

        # ── Return success response with any errors ───────────────────────────
        return jsonify({
            'success':  True,
            'filename': output_filename,
            'errors':   errors
        })

    except Exception as e:
        # ── Return error response if anything went wrong ──────────────────────
        # str(e) gives us the actual error message from Python
        # so we can show it to the user for debugging.
        return jsonify({
            'success': False,
            'error':   f'Conversion failed: {str(e)}'
        })


# ── ROUTE 3: Download ─────────────────────────────────────────────────────────
# This route serves the converted Excel file back to the user's browser.
# When the user clicks the Download button, their browser visits this URL
# and Flask sends the file as a download attachment.
@app.route('/download/<filename>')
def download_file(filename):

    # ── Security check ────────────────────────────────────────────────────────
    # We make sure the filename only contains safe characters.
    # This prevents attackers from requesting files outside our output folder
    # by using paths like '../../sensitive_file.txt'
    safe_filename = os.path.basename(filename)
    file_path     = os.path.join(app.config['OUTPUT_FOLDER'], safe_filename)

    # ── Check the file exists ─────────────────────────────────────────────────
    if not os.path.exists(file_path):
        return jsonify({
            'success': False,
            'error':   'File not found. It may have already been downloaded.'
        }), 404

    # ── Send the file to the browser ──────────────────────────────────────────
    # as_attachment=True tells the browser to download it rather than
    # trying to open it in the browser window.
    return send_file(
        file_path,
        as_attachment = True,
        download_name = safe_filename
    )


# ── MAIN: Start the web server ────────────────────────────────────────────────
# This block only runs when you execute app.py directly.
# debug=True means Flask will automatically restart when you
# make changes to the code — very useful during development.
# The app runs on http://localhost:5000 by default.
if __name__ == '__main__':
    print('Starting KoboToolbox XLSForm Converter...')
    print('Open your browser and go to: http://localhost:5000')
    app.run(debug=True)