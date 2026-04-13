#!/usr/bin/env python3
"""
Web app for Eureka Forbes P&L Analyzer.
Upload Excel -> Get PPT back.
"""
import os
import uuid
import tempfile
from flask import Flask, request, send_file, render_template_string, jsonify

from sample_data import generate_sample_excel
from analyzer import analyze_pnl
from ppt_generator import generate_ppt
from llm_insights import generate_insights

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

UPLOAD_DIR = tempfile.mkdtemp(prefix="efl_pnl_")

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Eureka Forbes — P&L Analyzer</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: #f0f4f8;
            color: #1a1a2e;
            min-height: 100vh;
        }

        .navbar {
            background: #1F4E79;
            padding: 16px 40px;
            display: flex;
            align-items: center;
            gap: 16px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }

        .navbar h1 {
            color: #fff;
            font-size: 20px;
            font-weight: 600;
        }

        .navbar .badge {
            background: rgba(255,255,255,0.15);
            color: #b8d4e8;
            font-size: 11px;
            padding: 3px 10px;
            border-radius: 12px;
        }

        .container {
            max-width: 800px;
            margin: 40px auto;
            padding: 0 20px;
        }

        .card {
            background: #fff;
            border-radius: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08), 0 4px 16px rgba(0,0,0,0.04);
            padding: 36px;
            margin-bottom: 24px;
        }

        .card h2 {
            font-size: 22px;
            margin-bottom: 8px;
            color: #1F4E79;
        }

        .card p.subtitle {
            color: #666;
            font-size: 14px;
            margin-bottom: 24px;
            line-height: 1.5;
        }

        .upload-zone {
            border: 2px dashed #c8d6e5;
            border-radius: 10px;
            padding: 48px 24px;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s;
            background: #fafbfc;
            position: relative;
        }

        .upload-zone:hover, .upload-zone.dragover {
            border-color: #2E75B6;
            background: #f0f6ff;
        }

        .upload-zone .icon {
            font-size: 48px;
            margin-bottom: 12px;
        }

        .upload-zone p {
            color: #666;
            font-size: 14px;
        }

        .upload-zone p strong {
            color: #2E75B6;
        }

        .upload-zone input[type="file"] {
            position: absolute;
            inset: 0;
            opacity: 0;
            cursor: pointer;
        }

        .file-selected {
            display: none;
            align-items: center;
            gap: 12px;
            padding: 14px 18px;
            background: #e8f4e8;
            border-radius: 8px;
            margin-top: 16px;
            font-size: 14px;
            color: #2d6a2d;
        }

        .file-selected.show { display: flex; }

        .file-selected .name { flex: 1; font-weight: 500; }

        .file-selected .remove {
            cursor: pointer;
            color: #999;
            font-size: 18px;
        }

        .options {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
            margin-top: 24px;
        }

        .options label {
            font-size: 13px;
            font-weight: 500;
            color: #444;
            display: block;
            margin-bottom: 6px;
        }

        .options select, .options input {
            width: 100%;
            padding: 10px 14px;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 14px;
            background: #fff;
            color: #333;
        }

        .options select:focus, .options input:focus {
            outline: none;
            border-color: #2E75B6;
            box-shadow: 0 0 0 3px rgba(46,117,182,0.1);
        }

        .btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 14px 32px;
            border: none;
            border-radius: 8px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            margin-top: 28px;
        }

        .btn-primary {
            background: #1F4E79;
            color: #fff;
            width: 100%;
            justify-content: center;
        }

        .btn-primary:hover { background: #163d5e; }

        .btn-primary:disabled {
            background: #a0b4c8;
            cursor: not-allowed;
        }

        .btn-secondary {
            background: #f0f4f8;
            color: #1F4E79;
            border: 1px solid #c8d6e5;
        }

        .btn-secondary:hover { background: #e0e8f0; }

        .progress-area {
            display: none;
            margin-top: 24px;
        }

        .progress-area.show { display: block; }

        .progress-bar-container {
            background: #e8ecf0;
            border-radius: 6px;
            height: 8px;
            overflow: hidden;
            margin-bottom: 12px;
        }

        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #2E75B6, #1F4E79);
            border-radius: 6px;
            width: 0%;
            transition: width 0.3s;
        }

        .progress-text {
            font-size: 13px;
            color: #666;
        }

        .result-area {
            display: none;
            margin-top: 24px;
            padding: 20px;
            background: #e8f5e9;
            border-radius: 10px;
            text-align: center;
        }

        .result-area.show { display: block; }

        .result-area h3 {
            color: #2d6a2d;
            margin-bottom: 8px;
        }

        .result-area .stats {
            font-size: 13px;
            color: #555;
            margin-bottom: 16px;
            line-height: 1.6;
        }

        .result-area .download-btn {
            background: #2d6a2d;
            color: #fff;
            padding: 12px 32px;
            border-radius: 8px;
            text-decoration: none;
            font-weight: 600;
            display: inline-block;
        }

        .result-area .download-btn:hover { background: #1e4e1e; }

        .error-area {
            display: none;
            margin-top: 24px;
            padding: 16px 20px;
            background: #fde8e8;
            border-radius: 10px;
            color: #a02020;
            font-size: 14px;
        }

        .error-area.show { display: block; }

        .sample-link {
            text-align: center;
            margin-top: 16px;
        }

        .sample-link a {
            color: #2E75B6;
            font-size: 13px;
            text-decoration: none;
        }

        .sample-link a:hover { text-decoration: underline; }

        .features {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
        }

        .feature {
            padding: 16px;
            background: #f8fafc;
            border-radius: 8px;
            border: 1px solid #e8ecf0;
        }

        .feature h4 {
            font-size: 13px;
            color: #1F4E79;
            margin-bottom: 4px;
        }

        .feature p {
            font-size: 12px;
            color: #888;
            line-height: 1.4;
        }

        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }

        .analyzing { animation: pulse 1.5s infinite; }
    </style>
</head>
<body>
    <nav class="navbar">
        <h1>Eureka Forbes — P&L Analyzer</h1>
        <span class="badge">Beta</span>
    </nav>

    <div class="container">
        <div class="card">
            <h2>Upload P&L Statement</h2>
            <p class="subtitle">
                Upload your monthly P&L Excel file. The tool will analyze performance vs AOP,
                previous month, last year, and detect all outliers — then generate a presentation-ready PPT.
            </p>

            <form id="uploadForm" enctype="multipart/form-data">
                <div class="upload-zone" id="dropZone">
                    <div class="icon">📊</div>
                    <p><strong>Click to upload</strong> or drag & drop your Excel file</p>
                    <p style="margin-top:6px; font-size:12px; color:#999">.xlsx files up to 50MB</p>
                    <input type="file" id="fileInput" name="file" accept=".xlsx,.xls">
                </div>

                <div class="file-selected" id="fileSelected">
                    <span>📎</span>
                    <span class="name" id="fileName"></span>
                    <span class="remove" id="removeFile">&times;</span>
                </div>

                <div class="options">
                    <div>
                        <label>Review Month</label>
                        <select name="month" id="monthSelect">
                            <option value="Apr">April</option>
                            <option value="May">May</option>
                            <option value="Jun">June</option>
                            <option value="Jul">July</option>
                            <option value="Aug">August</option>
                            <option value="Sep">September</option>
                            <option value="Oct">October</option>
                            <option value="Nov">November</option>
                            <option value="Dec">December</option>
                            <option value="Jan">January</option>
                            <option value="Feb">February</option>
                            <option value="Mar" selected>March</option>
                        </select>
                    </div>
                    <div>
                        <label>Financial Year</label>
                        <select name="fy" id="fySelect">
                            <option value="2025">FY25</option>
                            <option value="2026" selected>FY26</option>
                            <option value="2027">FY27</option>
                        </select>
                    </div>
                </div>

                <button type="submit" class="btn btn-primary" id="submitBtn" disabled>
                    Analyze & Generate PPT
                </button>
            </form>

            <div class="progress-area" id="progressArea">
                <div class="progress-bar-container">
                    <div class="progress-bar" id="progressBar"></div>
                </div>
                <p class="progress-text analyzing" id="progressText">Uploading file...</p>
            </div>

            <div class="result-area" id="resultArea">
                <h3>Analysis Complete</h3>
                <div class="stats" id="resultStats"></div>
                <a href="#" class="download-btn" id="downloadBtn">Download PPT</a>
            </div>

            <div class="error-area" id="errorArea"></div>

            <div class="sample-link">
                <a href="/sample">Or try with sample data →</a>
            </div>
        </div>

        <div class="card">
            <h2 style="font-size:16px; margin-bottom:16px;">What's in the PPT?</h2>
            <div class="features">
                <div class="feature">
                    <h4>Executive Summary</h4>
                    <p>KPI cards for Net Sales, Gross Margin, EBITDA, PBT, PAT with beat/miss indicators</p>
                </div>
                <div class="feature">
                    <h4>P&L Comparison Table</h4>
                    <p>Current month vs AOP, previous month (MoM%), and last year (YoY%)</p>
                </div>
                <div class="feature">
                    <h4>Budget Achievement</h4>
                    <p>Actual vs AOP with achievement %, variance, and YTD tracking</p>
                </div>
                <div class="feature">
                    <h4>Quarterly Comparison</h4>
                    <p>QoQ and YoY at quarter level with variance analysis</p>
                </div>
                <div class="feature">
                    <h4>Outlier Detection</h4>
                    <p>All MoM, YoY, QoQ, and vs-AOP outliers flagged with severity</p>
                </div>
                <div class="feature">
                    <h4>Cost Deep Dive</h4>
                    <p>Cost categories as % of Net Sales vs last year with delta tracking</p>
                </div>
                <div class="feature">
                    <h4>Revenue Trend</h4>
                    <p>Monthly bar chart with best/worst month and YTD KPIs</p>
                </div>
                <div class="feature">
                    <h4>Full Year View</h4>
                    <p>For March reviews: complete FY actual vs AOP vs previous year</p>
                </div>
            </div>
        </div>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const fileSelected = document.getElementById('fileSelected');
        const fileName = document.getElementById('fileName');
        const removeFile = document.getElementById('removeFile');
        const submitBtn = document.getElementById('submitBtn');
        const uploadForm = document.getElementById('uploadForm');
        const progressArea = document.getElementById('progressArea');
        const progressBar = document.getElementById('progressBar');
        const progressText = document.getElementById('progressText');
        const resultArea = document.getElementById('resultArea');
        const resultStats = document.getElementById('resultStats');
        const downloadBtn = document.getElementById('downloadBtn');
        const errorArea = document.getElementById('errorArea');

        // Drag and drop
        ['dragenter', 'dragover'].forEach(e => {
            dropZone.addEventListener(e, (ev) => {
                ev.preventDefault();
                dropZone.classList.add('dragover');
            });
        });

        ['dragleave', 'drop'].forEach(e => {
            dropZone.addEventListener(e, (ev) => {
                ev.preventDefault();
                dropZone.classList.remove('dragover');
            });
        });

        dropZone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                fileInput.files = files;
                showFile(files[0]);
            }
        });

        fileInput.addEventListener('change', () => {
            if (fileInput.files.length > 0) {
                showFile(fileInput.files[0]);
            }
        });

        function showFile(file) {
            fileName.textContent = file.name + ' (' + (file.size / 1024).toFixed(0) + ' KB)';
            fileSelected.classList.add('show');
            dropZone.style.display = 'none';
            submitBtn.disabled = false;
            resultArea.classList.remove('show');
            errorArea.classList.remove('show');
        }

        removeFile.addEventListener('click', () => {
            fileInput.value = '';
            fileSelected.classList.remove('show');
            dropZone.style.display = 'block';
            submitBtn.disabled = true;
        });

        uploadForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const formData = new FormData();
            formData.append('file', fileInput.files[0]);
            formData.append('month', document.getElementById('monthSelect').value);
            formData.append('fy', document.getElementById('fySelect').value);

            // Show progress
            submitBtn.disabled = true;
            submitBtn.textContent = 'Analyzing...';
            progressArea.classList.add('show');
            resultArea.classList.remove('show');
            errorArea.classList.remove('show');

            const steps = [
                { pct: 15, text: 'Uploading file...' },
                { pct: 30, text: 'Reading P&L data...' },
                { pct: 45, text: 'Running variance analysis...' },
                { pct: 55, text: 'Detecting outliers (MoM, YoY, QoQ, AOP)...' },
                { pct: 70, text: 'Generating AI-powered insights with Claude...' },
                { pct: 85, text: 'Building PowerPoint slides...' },
                { pct: 92, text: 'Adding narrative commentary...' },
            ];

            let stepIdx = 0;
            const interval = setInterval(() => {
                if (stepIdx < steps.length) {
                    progressBar.style.width = steps[stepIdx].pct + '%';
                    progressText.textContent = steps[stepIdx].text;
                    stepIdx++;
                }
            }, 600);

            try {
                const resp = await fetch('/analyze', {
                    method: 'POST',
                    body: formData,
                });

                clearInterval(interval);

                if (!resp.ok) {
                    const err = await resp.json();
                    throw new Error(err.error || 'Analysis failed');
                }

                const data = await resp.json();

                progressBar.style.width = '100%';
                progressText.textContent = 'Done!';
                progressText.classList.remove('analyzing');

                // Show result
                resultStats.innerHTML = `
                    <strong>${data.company}</strong> — ${data.period}<br>
                    Net Sales: Rs ${data.net_sales} Cr | EBITDA: Rs ${data.ebitda} Cr (${data.ebitda_pct}%)<br>
                    PAT: Rs ${data.pat} Cr | Outliers detected: <strong>${data.outlier_count}</strong>
                    (${data.high_severity} high severity)
                `;
                downloadBtn.href = '/download/' + data.file_id;
                resultArea.classList.add('show');

            } catch (err) {
                clearInterval(interval);
                progressArea.classList.remove('show');
                errorArea.textContent = 'Error: ' + err.message;
                errorArea.classList.add('show');
            }

            submitBtn.disabled = false;
            submitBtn.textContent = 'Analyze & Generate PPT';
        });
    </script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/analyze", methods=["POST"])
def analyze():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an Excel file (.xlsx)"}), 400

    month = request.form.get("month", "Mar")
    fy = int(request.form.get("fy", 2026))

    # Save uploaded file
    file_id = str(uuid.uuid4())[:8]
    upload_path = os.path.join(UPLOAD_DIR, f"{file_id}_input.xlsx")
    output_path = os.path.join(UPLOAD_DIR, f"{file_id}_analysis.pptx")
    file.save(upload_path)

    try:
        analysis = analyze_pnl(upload_path, review_month=month, review_fy=fy)

        # Generate LLM insights
        insights = None
        try:
            insights = generate_insights(analysis, ANTHROPIC_API_KEY)
        except Exception as llm_err:
            print(f"LLM insights failed (non-fatal): {llm_err}")

        generate_ppt(analysis, output_path, insights=insights)

        ns = analysis.current_month.get("Total Net Sales", 0)
        ebitda = analysis.current_month.get("EBITDA (post allocation)", 0)
        ebitda_pct = analysis.current_month.get("EBITDA %", 0)
        pat = analysis.current_month.get("Profit After Tax", 0)
        high_count = sum(1 for o in analysis.outliers if o.severity == "high")

        return jsonify({
            "file_id": file_id,
            "company": analysis.company,
            "period": f"{analysis.review_month} {analysis.review_fy}",
            "net_sales": f"{ns:,.1f}",
            "ebitda": f"{ebitda:,.1f}",
            "ebitda_pct": f"{ebitda_pct:.1f}",
            "pat": f"{pat:,.1f}",
            "outlier_count": len(analysis.outliers),
            "high_severity": high_count,
            "highlights": analysis.highlights,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download/<file_id>")
def download(file_id):
    output_path = os.path.join(UPLOAD_DIR, f"{file_id}_analysis.pptx")
    if not os.path.exists(output_path):
        return "File not found", 404
    return send_file(output_path, as_attachment=True,
                     download_name="EFL_PnL_Analysis.pptx")


@app.route("/sample")
def sample():
    """Generate sample data, analyze it, and return the PPT."""
    file_id = str(uuid.uuid4())[:8]
    sample_path = os.path.join(UPLOAD_DIR, f"{file_id}_sample.xlsx")
    output_path = os.path.join(UPLOAD_DIR, f"{file_id}_analysis.pptx")

    generate_sample_excel(sample_path, review_month="Mar", review_fy=2026)
    analysis = analyze_pnl(sample_path, review_month="Mar", review_fy=2026)

    insights = None
    try:
        insights = generate_insights(analysis, ANTHROPIC_API_KEY)
    except Exception as llm_err:
        print(f"LLM insights failed (non-fatal): {llm_err}")

    generate_ppt(analysis, output_path, insights=insights)

    return send_file(output_path, as_attachment=True,
                     download_name="EFL_Sample_PnL_Analysis_Mar_FY26.pptx")


if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("  Eureka Forbes P&L Analyzer")
    print("  http://localhost:5050")
    print("=" * 50 + "\n")
    app.run(host="0.0.0.0", port=5050, debug=True)
