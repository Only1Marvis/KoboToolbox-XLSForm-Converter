# ── Imports and Setup ─────────────────────────────────────────────────────────
import os
import uuid
from flask import Flask, request, render_template_string, send_file, jsonify
from convert import convert

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KoboToolbox XLSForm Converter</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
            min-height: 100vh;
            display: flex;
            align-items: flex-start;
            justify-content: center;
            padding: 2rem 1rem;
        }
        .container {
            background: white;
            border-radius: 16px;
            padding: 2.5rem;
            width: 100%;
            max-width: 860px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
        }
        h1 { font-size: 1.8rem; font-weight: 700; color: #1a1a2e; text-align: center; margin-bottom: 0.4rem; }
        .subtitle { text-align: center; color: #6b7280; font-size: 0.9rem; margin-bottom: 2rem; }
        .main-layout { display: flex; gap: 2rem; align-items: flex-start; margin-bottom: 1.5rem; }
        .left-panel { flex: 1; min-width: 0; }
        .right-panel { flex-shrink: 0; width: auto; }
        .upload-area {
            border: 2px dashed #c7d2fe;
            border-radius: 12px;
            padding: 2rem 1.5rem;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s ease;
            background: #fafbff;
            position: relative;
            margin-bottom: 0.75rem;
        }
        .upload-area:hover, .upload-area.dragover { border-color: #4472C4; background: #f0f4ff; }
        .upload-icon { font-size: 2.5rem; margin-bottom: 0.5rem; }
        .upload-area h3 { font-size: 1rem; color: #374151; margin-bottom: 0.25rem; }
        .upload-area p { font-size: 0.82rem; color: #9ca3af; }
        .file-input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
        .selected-file {
            display: flex; align-items: center; gap: 0.5rem;
            padding: 0.5rem 0.75rem; background: #f3f4f6;
            border-radius: 8px; font-size: 0.82rem; color: #374151;
            margin-bottom: 0.75rem; min-height: 36px;
        }
        .convert-btn {
            width: 100%; padding: 0.85rem; background: #4472C4;
            color: white; border: none; border-radius: 10px;
            font-size: 1rem; font-weight: 600; cursor: pointer;
            transition: all 0.2s ease;
        }
        .convert-btn:hover:not(:disabled) { background: #3461b0; }
        .convert-btn:disabled { background: #9ca3af; cursor: not-allowed; }
        .repeat-section label.section-label {
            display: block; font-size: 0.82rem; font-weight: 600;
            color: #374151; margin-bottom: 0.4rem;
        }
        .repeat-options { display: flex; flex-direction: column; gap: 0.3rem; width: fit-content; }
        .repeat-option {
            display: flex; align-items: center; gap: 0.55rem;
            padding: 0.5rem 0.75rem; border: 1px solid #e5e7eb;
            border-radius: 7px; cursor: pointer;
            transition: all 0.15s ease; position: relative;
            white-space: nowrap;
        }
        .repeat-option:hover { border-color: #4472C4; background: #f8faff; }
        .repeat-option input[type="radio"] { accent-color: #4472C4; flex-shrink: 0; margin: 0; }
        .repeat-option .option-title { font-size: 0.8rem; font-weight: 600; color: #1f2937; }
        .badge-default {
            font-size: 0.62rem; background: #4472C4; color: white;
            padding: 1px 5px; border-radius: 10px; font-weight: 600;
            margin-left: 4px; vertical-align: middle;
        }
        .repeat-option:has(input:checked) { border-color: #4472C4; background: #f0f4ff; }
        .repeat-option:has(input:checked) .option-title { color: #4472C4; }
        .repeat-option .tooltip-text {
            visibility: hidden; opacity: 0; width: 240px;
            white-space: normal; word-wrap: break-word;
            background: #1f2937; color: #fff; font-size: 0.72rem;
            line-height: 1.5; border-radius: 6px; padding: 8px 11px;
            position: absolute; z-index: 200;
            bottom: calc(100% + 10px); left: 0;
            transition: opacity 0.2s; pointer-events: none;
            box-shadow: 0 4px 14px rgba(0,0,0,0.25);
        }
        .repeat-option .tooltip-text::after {
            content: ""; position: absolute; top: 100%; left: 20px;
            border: 6px solid transparent; border-top-color: #1f2937;
        }
        .repeat-option:hover .tooltip-text { visibility: visible; opacity: 1; }
        .output-options {
            margin-top: 1.2rem; padding: 0.9rem 1.1rem;
            background: #f8faff; border: 1px solid #e5e7eb; border-radius: 10px;
        }
        .output-options .section-label {
            display: block; font-size: 0.82rem; font-weight: 600;
            color: #374151; margin-bottom: 0.6rem;
        }
        .opt-row {
            display: inline-flex; align-items: center; gap: 0.55rem;
            padding: 0.45rem 0.65rem; border: 1px solid #e5e7eb;
            border-radius: 7px; cursor: pointer; position: relative;
            transition: all 0.15s ease; white-space: nowrap;
        }
        .opt-row:hover { border-color: #4472C4; background: #f0f4ff; }
        .opt-row input[type="checkbox"] { accent-color: #4472C4; flex-shrink: 0; }
        .opt-row .opt-label { font-size: 0.8rem; font-weight: 600; color: #1f2937; }
        .opt-row:has(input:checked) { border-color: #4472C4; background: #f0f4ff; }
        .opt-row .opt-tooltip {
            visibility: hidden; opacity: 0; width: 260px;
            white-space: normal; word-wrap: break-word;
            background: #1f2937; color: #fff; font-size: 0.72rem;
            line-height: 1.5; border-radius: 6px; padding: 8px 11px;
            position: absolute; z-index: 200;
            bottom: calc(100% + 10px); left: 0;
            transition: opacity 0.2s; pointer-events: none;
            box-shadow: 0 4px 14px rgba(0,0,0,0.25);
        }
        .opt-row .opt-tooltip::after {
            content: ""; position: absolute; top: 100%; left: 20px;
            border: 6px solid transparent; border-top-color: #1f2937;
        }
        .opt-row:hover .opt-tooltip { visibility: visible; opacity: 1; }
        .status-box {
            margin-top: 1.2rem; padding: 1rem 1.25rem; border-radius: 10px;
            font-size: 0.88rem; line-height: 1.6; display: none; text-align: center;
        }
        .status-box.loading { background: #fefce8; border: 1px solid #fde68a; color: #854d0e; display: block; }
        .status-box.success { background: #f0fdf4; border: 1px solid #86efac; color: #166534; display: block; }
        .status-box.error   { background: #fef2f2; border: 1px solid #fca5a5; color: #991b1b; display: block; }
        .download-btn {
            display: inline-block; margin-top: 0.75rem; padding: 0.6rem 1.5rem;
            background: #4472C4; color: white; border-radius: 8px;
            text-decoration: none; font-weight: 600; font-size: 0.88rem;
        }
        .download-btn:hover { background: #3461b0; }
        .spinner {
            display: inline-block; width: 14px; height: 14px;
            border: 2px solid #fde68a; border-top-color: #854d0e;
            border-radius: 50%; animation: spin 0.7s linear infinite;
            margin-right: 6px; vertical-align: middle;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .rules-note {
            margin-top: 1.5rem; padding-top: 1rem; border-top: 1px solid #f3f4f6;
            font-size: 0.8rem; color: #9ca3af; text-align: center; line-height: 1.6;
        }
        .credit {
            margin-top: 0.75rem; padding-top: 0.75rem; border-top: 1px solid #f3f4f6;
            text-align: center; font-size: 0.75rem; color: #c4c4c4;
            font-style: italic; letter-spacing: 0.03em;
        }
    </style>
</head>
<body>
<div class="container">

    <h1>KoboToolbox XLSForm Converter</h1>
    <p class="subtitle">Upload a Word questionnaire and get a ready-to-use XLSForm</p>

    <div class="main-layout">

        <div class="left-panel">
            <div class="upload-area" id="uploadArea">
                <div class="upload-icon">📄</div>
                <h3>Drop your Word document here</h3>
                <p>or click to browse — .docx files only</p>
                <input type="file" id="fileInput" class="file-input" accept=".docx" />
            </div>
            <div class="selected-file" id="selectedFile">
                <span>📎</span>
                <span id="fileName">No file selected</span>
            </div>
            <button class="convert-btn" id="convertBtn" disabled>
                Convert to XLSForm
            </button>
        </div>

        <div class="right-panel">
            <div class="repeat-section">
                <label class="section-label">🔁 Repeat Group Format</label>
                <div class="repeat-options">

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="3" />
                        <div class="option-title">Automated Repeat Loop</div>
                        <span class="tooltip-text">True begin_repeat driven by a prior select_multiple. Adapts automatically per respondent. Best for large or variable lists.</span>
                    </label>

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="2" />
                        <div class="option-title">Brand-Fixed Group Loop</div>
                        <span class="tooltip-text">Each brand gets its own permanently fixed group that appears conditionally when selected. Best for brand equity tracking, retail audits and competitive benchmarking.</span>
                    </label>

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="1" />
                        <div class="option-title">Sequential Positional Loop</div>
                        <span class="tooltip-text">Fixed unrolled slot groups — enumerator picks one brand per slot. Already-selected brands are excluded from subsequent slots. Best for product testing, sensory evaluation and ranked preference studies.</span>
                    </label>

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="4" checked />
                        <div class="option-title">🧠 Smart Auto-Select <span class="badge-default">Default</span></div>
                        <span class="tooltip-text">The engine evaluates each repeat group individually — considering the number of options, how the study is structured, and what works best for the raw data output. Different groups in the same questionnaire may use different formats.</span>
                    </label>

                    <label class="repeat-option">
                        <input type="radio" name="repeatFormat" value="5" />
                        <div class="option-title">✍️ Direct Scripting Mode</div>
                        <span class="tooltip-text">No system repeat groups are applied. Questions are scripted exactly as written in the questionnaire — group by group, question by question, with no automated repeat wrapping.</span>
                    </label>

                </div>
            </div>
        </div>

    </div>

    <div class="output-options">
        <span class="section-label">⚙️ Output Options</span>
        <div style="display:flex; gap:0.75rem; flex-wrap:wrap;">
            <label class="opt-row">
                <input type="checkbox" id="includeCode" name="includeCode" checked />
                <span class="opt-label">Include question code in label</span>
                <span class="opt-tooltip">When checked, question labels include the question code — e.g. "RSD1a. What is your name?". When unchecked, only the question text appears — e.g. "What is your name?". Recommended to keep checked for easier data reference.</span>
            </label>
        </div>
    </div>

    <div class="status-box" id="statusBox"></div>

    <div class="rules-note">
        Applies all KoboToolbox XLSForm rules automatically —
        sections, groups, grids, SO questions, choice filters and more.
    </div>

    <div class="credit">By Marvis Onyenwenu Enubiaka</div>

</div>

<script>
    const fileInput    = document.getElementById("fileInput");
    const uploadArea   = document.getElementById("uploadArea");
    const fileName     = document.getElementById("fileName");
    const convertBtn   = document.getElementById("convertBtn");
    const statusBox    = document.getElementById("statusBox");

    fileInput.addEventListener("change", () => {
        const file = fileInput.files[0];
        if (file) {
            fileName.textContent = file.name;
            convertBtn.disabled  = false;
            statusBox.className  = "status-box";
            statusBox.innerHTML  = "";
        }
    });

    uploadArea.addEventListener("dragover", e => { e.preventDefault(); uploadArea.classList.add("dragover"); });
    uploadArea.addEventListener("dragleave", () => { uploadArea.classList.remove("dragover"); });
    uploadArea.addEventListener("drop", e => {
        e.preventDefault();
        uploadArea.classList.remove("dragover");
        const file = e.dataTransfer.files[0];
        if (file && file.name.endsWith(".docx")) {
            const dt = new DataTransfer();
            dt.items.add(file);
            fileInput.files = dt.files;
            fileName.textContent = file.name;
            convertBtn.disabled  = false;
        }
    });

    function showStatus(type, message) {
        statusBox.className = "status-box " + type;
        statusBox.innerHTML = message;
    }

    convertBtn.addEventListener("click", async () => {
        const file = fileInput.files[0];
        if (!file) return;

        const repeatFormatEl = document.querySelector("input[name=\"repeatFormat\"]:checked");
        const repeatFormat   = repeatFormatEl ? repeatFormatEl.value : "4";
        const includeCode    = document.getElementById("includeCode").checked;

        convertBtn.disabled = true;
        showStatus("loading", "<span class=\"spinner\"></span> Converting your questionnaire... please wait.");
        await new Promise(r => setTimeout(r, 80));

        const formData = new FormData();
        formData.append("file", file);
        formData.append("repeat_format", repeatFormat);
        formData.append("include_code", includeCode ? "1" : "0");

        try {
            const response = await fetch("/convert", { method: "POST", body: formData });
            const result   = await response.json();

            if (result.success) {
                let html = "✅ Conversion successful! Your XLSForm is ready.<br>" +
                    "<a class=\"download-btn\" href=\"/download/" + result.filename + "\">⬇ Download XLSForm</a>";

                if (result.errors && result.errors.length > 0) {
                    html += "<div style=\"margin-top:1rem;padding:0.75rem 1rem;background:#fef9c3;border:1px solid #fde68a;border-radius:8px;text-align:left;\">" +
                        "<strong style=\"color:#854d0e;\">⚠ " + result.errors.length + " issue(s) found — please review:</strong>" +
                        "<ol style=\"margin:0.5rem 0 0.5rem 1.2rem;color:#854d0e;font-size:0.82rem;line-height:1.7;\">";
                    result.errors.forEach(err => { html += "<li>" + err + "</li>"; });
                    html += "</ol><p style=\"margin:0.5rem 0 0;font-size:0.82rem;color:#854d0e;font-style:italic;\">Kindly effect these corrections on the Word document questionnaire and re-upload.</p></div>";
                }
                showStatus("success", html);
            } else {
                showStatus("error", "❌ Error: " + result.error);
            }
        } catch (err) {
            showStatus("error", "❌ Something went wrong. Please try again.");
        }
        convertBtn.disabled = false;
    });
</script>
</body>
</html>"""

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file was uploaded.'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file selected.'})
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': 'Invalid file type. Please upload a .docx file.'})
    try:
        unique_id       = str(uuid.uuid4())[:8]
        original_name   = os.path.splitext(file.filename)[0]
        input_filename  = f'{original_name}_{unique_id}.docx'
        input_path      = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(input_path)
        output_filename = f'{original_name}_{unique_id}_XLSForm.xlsx'
        output_path     = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        repeat_format   = int(request.form.get('repeat_format', 4))
        include_code    = request.form.get('include_code', '1') == '1'
        output_path, errors = convert(
            filepath=input_path, output_path=output_path,
            use_ai=False, repeat_format=repeat_format, include_code=include_code
        )
        os.remove(input_path)
        return jsonify({'success': True, 'filename': output_filename, 'errors': errors})
    except Exception as e:
        return jsonify({'success': False, 'error': f'Conversion failed: {str(e)}'})

@app.route('/download/<filename>')
def download_file(filename):
    safe_filename = os.path.basename(filename)
    file_path     = os.path.join(app.config['OUTPUT_FOLDER'], safe_filename)
    if not os.path.exists(file_path):
        return jsonify({'success': False, 'error': 'File not found.'}), 404
    return send_file(file_path, as_attachment=True, download_name=safe_filename)

if __name__ == '__main__':
    print('Starting KoboToolbox XLSForm Converter...')
    print('Open your browser and go to: http://localhost:5000')
    app.run(debug=True)
