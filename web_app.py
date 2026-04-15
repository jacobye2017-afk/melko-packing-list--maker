"""
MELKO Packing List Maker - Web Version (Flask)
"""

import os
import io
import zipfile
import tempfile
import shutil
from flask import Flask, request, send_file, render_template_string, jsonify
import converter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

HTML = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MELKO Packing List Maker</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f0f2f5; min-height: 100vh; }
        .header { background: #2c3e50; color: white; padding: 20px 0; text-align: center; }
        .header h1 { font-size: 24px; margin-bottom: 4px; }
        .header p { font-size: 13px; color: #94a3b8; }
        .container { max-width: 800px; margin: 30px auto; padding: 0 20px; }
        .card { background: white; border-radius: 12px; padding: 30px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .card h2 { font-size: 18px; color: #2c3e50; margin-bottom: 6px; }
        .card .desc { font-size: 13px; color: #888; margin-bottom: 18px; }
        .upload-zone { border: 2px dashed #cbd5e1; border-radius: 10px; padding: 40px 20px; text-align: center; cursor: pointer; transition: all 0.2s; background: #fafbfc; }
        .upload-zone:hover, .upload-zone.drag-over { border-color: #3454D1; background: #eef2ff; }
        .upload-zone input { display: none; }
        .upload-zone .icon { font-size: 40px; margin-bottom: 10px; }
        .upload-zone .text { font-size: 14px; color: #64748b; }
        .upload-zone .text b { color: #3454D1; }
        .btn { display: inline-block; padding: 12px 28px; border: none; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer; transition: all 0.2s; }
        .btn-blue { background: #3454D1; color: white; }
        .btn-blue:hover { background: #2840A0; }
        .btn-green { background: #27AE60; color: white; }
        .btn-green:hover { background: #1E8449; }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; }
        .status { margin-top: 15px; padding: 12px 16px; border-radius: 8px; font-size: 14px; display: none; }
        .status.success { display: block; background: #f0fdf4; color: #166534; border: 1px solid #bbf7d0; }
        .status.error { display: block; background: #fef2f2; color: #991b1b; border: 1px solid #fecaca; }
        .status.loading { display: block; background: #eff6ff; color: #1e40af; border: 1px solid #bfdbfe; }
        .steps { display: flex; gap: 8px; margin-bottom: 20px; }
        .step-badge { padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }
        .step-blue { background: #dbeafe; color: #1d4ed8; }
        .step-green { background: #dcfce7; color: #166534; }
        .file-name { font-size: 13px; color: #3454D1; margin-top: 8px; font-weight: 500; word-break: break-all; }
        .instructions { font-size: 13px; color: #64748b; line-height: 1.8; }
        .instructions li { margin-bottom: 4px; }
        .footer { text-align: center; padding: 20px; font-size: 12px; color: #94a3b8; }
    </style>
</head>
<body>
    <div class="header">
        <h1>MELKO Packing List Maker</h1>
        <p>Packing List & BOL Generator</p>
    </div>
    <div class="container">
        <!-- Step 1 -->
        <div class="card">
            <div class="steps"><span class="step-badge step-blue">Step 1</span></div>
            <h2>Generate Packing List / 生成派送方案单</h2>
            <p class="desc">上传客户的 Excel 源文件 (.xlsx / .xls)，自动生成标准 Packing List</p>
            <div class="upload-zone" id="zone1" onclick="document.getElementById('file1').click()">
                <input type="file" id="file1" accept=".xlsx,.xls" onchange="handleStep1(this)">
                <div class="icon">📄</div>
                <div class="text">拖拽文件到此处 或 <b>点击选择</b></div>
                <div class="file-name" id="fname1"></div>
            </div>
            <div class="status" id="status1"></div>
        </div>

        <!-- Step 2 -->
        <div class="card">
            <div class="steps">
                <span class="step-badge step-blue">Step 2</span>
                <span style="font-size:13px;color:#888;line-height:28px;">打开 Step 1 生成的 Packing List，检查修改</span>
            </div>
        </div>

        <!-- Step 3 -->
        <div class="card">
            <div class="steps"><span class="step-badge step-green">Step 3</span></div>
            <h2>Generate BOL Files / 生成 BOL 文件</h2>
            <p class="desc">上传检查修改好的 Packing List，自动生成 4 个 BOL 文件 (ZIP 下载)</p>
            <div class="upload-zone" id="zone2" onclick="document.getElementById('file2').click()">
                <input type="file" id="file2" accept=".xlsx" onchange="handleStep2(this)">
                <div class="icon">📦</div>
                <div class="text">拖拽 Packing List 到此处 或 <b>点击选择</b></div>
                <div class="file-name" id="fname2"></div>
            </div>
            <div class="status" id="status2"></div>
        </div>

        <!-- Instructions -->
        <div class="card">
            <h2>使用说明</h2>
            <ol class="instructions">
                <li><b>Step 1</b>: 上传客户提供的预报表/派送计划 Excel → 下载生成的 Packing List</li>
                <li><b>Step 2</b>: 打开下载的 Packing List，人工检查并修改</li>
                <li><b>Step 3</b>: 上传修改好的 Packing List → 下载 BOL 文件包 (ZIP)，包含:
                    <ul style="margin-top:4px;padding-left:20px;">
                        <li>Packing List (A-M 13列)</li>
                        <li>Hold List (仅 HOLD/RELABEL 行)</li>
                        <li>TRUCKING.docx (仓库板数汇总)</li>
                        <li>柜号唛头.docx (柜号 + MELKO)</li>
                    </ul>
                </li>
            </ol>
        </div>
    </div>
    <div class="footer">MELKO Logistics &copy; 2026</div>

    <script>
        // Drag-drop support
        ['zone1','zone2'].forEach(id => {
            const zone = document.getElementById(id);
            zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
            zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
            zone.addEventListener('drop', e => {
                e.preventDefault();
                zone.classList.remove('drag-over');
                const file = e.dataTransfer.files[0];
                if (!file) return;
                const input = zone.querySelector('input');
                const dt = new DataTransfer();
                dt.items.add(file);
                input.files = dt.files;
                input.dispatchEvent(new Event('change'));
            });
        });

        function setStatus(id, type, msg) {
            const el = document.getElementById(id);
            el.className = 'status ' + type;
            el.textContent = msg;
        }

        async function handleStep1(input) {
            const file = input.files[0];
            if (!file) return;
            document.getElementById('fname1').textContent = file.name;
            setStatus('status1', 'loading', '正在处理...');

            const form = new FormData();
            form.append('file', file);

            try {
                const resp = await fetch('/api/step1', { method: 'POST', body: form });
                if (!resp.ok) {
                    const err = await resp.json();
                    throw new Error(err.error || 'Unknown error');
                }
                const blob = await resp.blob();
                const cd = resp.headers.get('Content-Disposition') || '';
                const match = cd.match(/filename="?([^"]+)"?/);
                const fname = match ? match[1] : 'packing_list.xlsx';

                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = decodeURIComponent(fname); a.click();
                URL.revokeObjectURL(url);

                setStatus('status1', 'success', '✅ 已生成: ' + decodeURIComponent(fname) + ' (已自动下载)');
            } catch (e) {
                setStatus('status1', 'error', '❌ ' + e.message);
            }
            input.value = '';
        }

        async function handleStep2(input) {
            const file = input.files[0];
            if (!file) return;
            document.getElementById('fname2').textContent = file.name;
            setStatus('status2', 'loading', '正在生成 BOL 文件...');

            const form = new FormData();
            form.append('file', file);

            try {
                const resp = await fetch('/api/step2', { method: 'POST', body: form });
                if (!resp.ok) {
                    const err = await resp.json();
                    throw new Error(err.error || 'Unknown error');
                }
                const blob = await resp.blob();
                const cd = resp.headers.get('Content-Disposition') || '';
                const match = cd.match(/filename="?([^"]+)"?/);
                const fname = match ? match[1] : 'BOL.zip';

                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = decodeURIComponent(fname); a.click();
                URL.revokeObjectURL(url);

                setStatus('status2', 'success', '✅ 已生成 BOL 文件包: ' + decodeURIComponent(fname) + ' (已自动下载)');
            } catch (e) {
                setStatus('status2', 'error', '❌ ' + e.message);
            }
            input.value = '';
        }
    </script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/api/step1', methods=['POST'])
def step1():
    """Upload source Excel → return Packing List xlsx."""
    if 'file' not in request.files:
        return jsonify(error="No file uploaded"), 400

    f = request.files['file']
    if not f.filename:
        return jsonify(error="Empty filename"), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.xlsx', '.xls'):
        return jsonify(error=f"Unsupported file type: {ext}"), 400

    tmpdir = tempfile.mkdtemp()
    try:
        src_path = os.path.join(tmpdir, f.filename)
        f.save(src_path)

        # Read and transform
        raw_headers, rows, header_map, header_to_canon, metadata = converter.read_source(src_path)
        container_no = metadata.get("container") or converter._extract_container(f.filename)
        sorted_rows = converter.sort_rows(rows, header_map)
        transformed = [converter.transform_row(r, header_map) for r in sorted_rows]

        all_candidates = [
            h for h in raw_headers
            if header_to_canon.get(h) is None and h
            and converter._norm_header_key(h) not in converter.SKIP_HEADER_KEYS
        ]
        extras_headers = [
            h for h in all_candidates
            if any(sr.get(h) not in (None, "") for sr in sorted_rows)
        ]

        import openpyxl
        out_path = os.path.join(tmpdir, f"{container_no} packing list.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "\u603b\u8868"
        for col_letter, w in converter.COLUMN_WIDTHS.items():
            ws.column_dimensions[col_letter].width = w
        converter._write_header_block(ws, container_no, metadata, extras_headers)
        last_data = converter._write_data_rows(ws, transformed, sorted_rows, extras_headers)
        converter._write_total_row(ws, last_data, transformed, len(extras_headers))
        ws.freeze_panes = "A4"
        wb.save(out_path)

        return send_file(
            out_path,
            as_attachment=True,
            download_name=f"{container_no} packing list.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


@app.route('/api/step2', methods=['POST'])
def step2():
    """Upload reviewed Packing List → return ZIP of BOL files."""
    if 'file' not in request.files:
        return jsonify(error="No file uploaded"), 400

    f = request.files['file']
    if not f.filename:
        return jsonify(error="Empty filename"), 400

    tmpdir = tempfile.mkdtemp()
    try:
        src_path = os.path.join(tmpdir, f.filename)
        f.save(src_path)

        import openpyxl
        wb = openpyxl.load_workbook(src_path, data_only=True)
        ws = wb.active

        container_no = converter._norm_str(ws.cell(row=1, column=1).value)
        if not container_no:
            container_no = converter._extract_container(f.filename)

        metadata = {"container": container_no, "etd": "", "eta": ""}
        etd_val = ws.cell(row=2, column=6).value
        eta_val = ws.cell(row=2, column=11).value
        if etd_val:
            metadata["etd"] = converter._norm_str(etd_val)
        if eta_val:
            metadata["eta"] = converter._norm_str(eta_val)

        # Read data
        transformed = []
        for row in range(4, ws.max_row + 1):
            method = converter._norm_str(ws.cell(row=row, column=3).value)
            if method == "TOTAL" or not method:
                continue
            ship_id = converter._norm_str(ws.cell(row=row, column=2).value)
            if not ship_id:
                continue
            r = {
                "no": converter._norm_str(ws.cell(row=row, column=1).value),
                "ship_id": ship_id,
                "method": method,
                "ctns": converter._to_num(ws.cell(row=row, column=4).value),
                "cbm": converter._to_num(ws.cell(row=row, column=5).value),
                "kg": converter._to_num(ws.cell(row=row, column=6).value),
                "ctn_lbs": converter._to_num(ws.cell(row=row, column=7).value),
                "total_lbs": converter._to_num(ws.cell(row=row, column=8).value),
                "description": converter._norm_str(ws.cell(row=row, column=9).value),
                "fba_code": converter._norm_str(ws.cell(row=row, column=10).value),
                "address": converter._norm_str(ws.cell(row=row, column=11).value),
                "fba_id": converter._norm_str(ws.cell(row=row, column=12).value),
                "reference_id": converter._norm_str(ws.cell(row=row, column=13).value),
                "loading_order": converter._norm_str(ws.cell(row=row, column=14).value),
                "eta_date": converter._norm_str(ws.cell(row=row, column=15).value),
                "trucking_no": converter._norm_str(ws.cell(row=row, column=16).value),
                "rate": converter._norm_str(ws.cell(row=row, column=17).value),
                "remark": converter._norm_str(ws.cell(row=row, column=18).value),
            }
            if r["ctns"] == int(r["ctns"]):
                r["ctns"] = int(r["ctns"])
            transformed.append(r)

        # Generate BOL files
        bol_dir = os.path.join(tmpdir, "BOL")
        os.makedirs(bol_dir)

        files_generated = []

        bol_pl = os.path.join(bol_dir, f"{container_no} packing list.xlsx")
        converter.build_bol_packing_list(transformed, container_no, bol_pl)
        files_generated.append(bol_pl)

        bol_hl = os.path.join(bol_dir, f"{container_no} hold list.xlsx")
        if converter.build_bol_hold_list(transformed, container_no, bol_hl):
            files_generated.append(bol_hl)

        bol_tr = os.path.join(bol_dir, f"{container_no}-TRUCKING.docx")
        converter.build_bol_trucking(transformed, container_no, bol_tr)
        files_generated.append(bol_tr)

        bol_sm = os.path.join(bol_dir, f"{container_no}-\u67dc\u53f7\u551b\u5934.docx")
        converter.build_bol_ship_marks(container_no, bol_sm)
        files_generated.append(bol_sm)

        # Create ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fp in files_generated:
                zf.write(fp, os.path.basename(fp))
        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=f"{container_no} BOL.zip",
            mimetype='application/zip',
        )
    except Exception as e:
        return jsonify(error=str(e)), 500
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
