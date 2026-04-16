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

HTML = r"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MELKO Packing List Maker</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',system-ui,sans-serif;background:#f7f9ff;color:#091d2e;min-height:100vh}
.header{background:linear-gradient(135deg,#1039b9,#3454d1);padding:24px 0;text-align:center;box-shadow:0 4px 20px rgba(16,57,185,0.15)}
.header-inner{max-width:900px;margin:0 auto;padding:0 24px;display:flex;align-items:center;justify-content:center;gap:16px}
.header img{height:56px}
.header h1{font-size:22px;font-weight:800;color:#fff;letter-spacing:-0.5px}
.header p{font-size:12px;color:rgba(255,255,255,0.7)}
.container{max-width:900px;margin:32px auto;padding:0 20px;display:grid;grid-template-columns:1fr 300px;gap:24px}
@media(max-width:768px){.container{grid-template-columns:1fr}}
.main{display:flex;flex-direction:column;gap:20px}
.card{background:#fff;border-radius:16px;padding:28px;box-shadow:0 2px 12px rgba(0,0,0,0.04);border:1px solid rgba(196,197,214,0.2)}
.card h2{font-size:18px;font-weight:700;margin-bottom:4px;color:#091d2e}
.card .step-badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;margin-bottom:10px}
.step-blue{background:#dde1ff;color:#1039b9}
.step-check{background:#dcfce7;color:#166534}
.step-green{background:#d1fae5;color:#065f46}
.card .desc{font-size:13px;color:#444654;margin-bottom:16px}
.upload-zone{border:2px dashed #c4c5d6;border-radius:14px;padding:40px 20px;text-align:center;cursor:pointer;transition:all 0.25s;background:#f7f9ff}
.upload-zone:hover,.upload-zone.drag-over{border-color:#3454d1;background:#eef1ff;box-shadow:0 0 0 4px rgba(52,84,209,0.08)}
.upload-zone input{display:none}
.upload-zone .icon{font-size:44px;margin-bottom:8px}
.upload-zone .text{font-size:14px;color:#444654;font-weight:500}
.upload-zone .text b{color:#3454d1}
.upload-zone .sub{font-size:11px;color:#747685;margin-top:4px}
.upload-zone .browse{display:inline-block;margin-top:14px;padding:8px 22px;background:#fff;border:1.5px solid #c4c5d6;border-radius:10px;font-size:13px;font-weight:600;color:#1039b9;cursor:pointer;transition:all 0.2s}
.upload-zone .browse:hover{background:#eef1ff;border-color:#3454d1}
.fname{font-size:12px;color:#3454d1;margin-top:8px;font-weight:500;word-break:break-all}
.review-box{display:flex;gap:12px;align-items:flex-start;padding:20px;background:#f7f9ff;border-radius:12px;border:1px solid #dde1ff}
.review-box .r-icon{font-size:28px;flex-shrink:0}
.review-box .r-title{font-size:14px;font-weight:600}
.review-box .r-desc{font-size:12px;color:#444654;margin-top:2px}
.status{margin-top:12px;padding:10px 14px;border-radius:10px;font-size:13px;display:none}
.status.success{display:block;background:#f0fdf4;color:#166534;border:1px solid #bbf7d0}
.status.error{display:block;background:#fef2f2;color:#991b1b;border:1px solid #fecaca}
.status.loading{display:block;background:#eef1ff;color:#1039b9;border:1px solid #bfdbfe}
.sidebar{display:flex;flex-direction:column;gap:20px}
.feat-card{background:linear-gradient(135deg,#1039b9,#3454d1);color:#fff;border-radius:16px;padding:24px;box-shadow:0 8px 24px rgba(16,57,185,0.15)}
.feat-card h3{font-size:16px;font-weight:700;margin-bottom:12px}
.feat-card .feat-item{display:flex;gap:10px;margin-bottom:10px;align-items:flex-start}
.feat-card .feat-item .dot{width:20px;height:20px;border-radius:50%;background:rgba(255,255,255,0.15);display:flex;align-items:center;justify-content:center;font-size:12px;flex-shrink:0;margin-top:1px}
.feat-card .feat-item .ft{font-size:12px;font-weight:600}
.feat-card .feat-item .fd{font-size:11px;opacity:0.7}
.bol-card{background:#fff;border-radius:16px;padding:24px;border:1px solid rgba(196,197,214,0.2)}
.bol-card h3{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:#4e6073;margin-bottom:14px}
.bol-card li{display:flex;gap:10px;align-items:flex-start;margin-bottom:10px}
.bol-card .num{width:22px;height:22px;border-radius:6px;background:#f1f5f9;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#64748b;flex-shrink:0}
.bol-card .bt{font-size:12px;color:#444654}
.bol-card .bt strong{color:#091d2e}
.footer{text-align:center;padding:24px;font-size:11px;color:#94a3b8;letter-spacing:1px;text-transform:uppercase;font-weight:600}
</style>
</head>
<body>
<header class="header">
<div class="header-inner">
<img src="/static/logo.png" alt="MELKO">
<div>
<h1>Packing List Maker</h1>
<p>Packing List & BOL Generator</p>
</div>
</div>
</header>

<div class="container">
<div class="main">
<!-- Step 1 -->
<div class="card">
<span class="step-badge step-blue">Step 1</span>
<h2>生成 Packing List</h2>
<p class="desc">上传客户的 Excel 源文件 (.xlsx / .xls)，自动生成标准派送方案单</p>
<div class="upload-zone" id="zone1" onclick="document.getElementById('file1').click()">
<input type="file" id="file1" accept=".xlsx,.xls" onchange="handleStep1(this)">
<div class="icon">📄</div>
<div class="text">拖拽客户 Excel 文件到此处 或 <b>点击选择</b></div>
<div class="sub">支持 .XLSX / .XLS 格式</div>
<div class="browse" onclick="event.stopPropagation();document.getElementById('file1').click()">浏览文件</div>
<div class="fname" id="fname1"></div>
</div>
<div class="status" id="status1"></div>
</div>

<!-- Step 2 -->
<div class="card">
<span class="step-badge step-check">Step 2</span>
<h2>检查 Packing List</h2>
<div class="review-box">
<div class="r-icon">📋</div>
<div>
<p class="r-title">打开下载的 Packing List</p>
<p class="r-desc">人工检查数据是否正确（颜色、分组、件数等），修改后保存。然后进入 Step 3。</p>
</div>
</div>
</div>

<!-- Step 3 -->
<div class="card">
<span class="step-badge step-green">Step 3</span>
<h2>生成 BOL 文件</h2>
<p class="desc">上传检查修改好的 Packing List，自动生成 BOL 文件包 (ZIP 下载)</p>
<div class="upload-zone" id="zone2" onclick="document.getElementById('file2').click()">
<input type="file" id="file2" accept=".xlsx" onchange="handleStep2(this)">
<div class="icon">📦</div>
<div class="text">拖拽修改好的 Packing List 到此处 或 <b>点击选择</b></div>
<div class="sub">上传后自动生成 BOL 文件包</div>
<div class="browse" onclick="event.stopPropagation();document.getElementById('file2').click()">浏览文件</div>
<div class="fname" id="fname2"></div>
</div>
<div class="status" id="status2"></div>
</div>
</div>

<!-- Sidebar -->
<div class="sidebar">
<div class="feat-card">
<h3>✨ 自动化功能</h3>
<div class="feat-item"><div class="dot">✓</div><div><div class="ft">KG → LBS 自动转换</div><div class="fd">×2.2 自动计算 CTN/LBS 和总 LBS</div></div></div>
<div class="feat-item"><div class="dot">✓</div><div><div class="ft">颜色自动标记</div><div class="fd">FEDEX紫 / UPS褐 / HOLD红 / FF黄</div></div></div>
<div class="feat-item"><div class="dot">✓</div><div><div class="ft">仓库代码自动合并</div><div class="fd">FBA CODE 连续相同值合并单元格</div></div></div>
<div class="feat-item"><div class="dot">✓</div><div><div class="ft">支持 .xls + .xlsx</div><div class="fd">兼容多种客户模板格式</div></div></div>
<div class="feat-item"><div class="dot">✓</div><div><div class="ft">柜号自动识别</div><div class="fd">从文件名/表头自动提取柜号</div></div></div>
</div>
<div class="bol-card">
<h3>BOL 文件包含</h3>
<ul style="list-style:none">
<li><span class="num">01</span><span class="bt"><strong>Packing List</strong> (A-M 13列精简版)</span></li>
<li><span class="num">02</span><span class="bt"><strong>Hold List</strong> (仅 HOLD/RELABEL 行)</span></li>
<li><span class="num">03</span><span class="bt"><strong>TRUCKING.docx</strong> (仓库板数/CBM 汇总)</span></li>
<li><span class="num">04</span><span class="bt"><strong>柜号唛头.docx</strong> (柜号 + MELKO)</span></li>
</ul>
</div>
</div>
</div>

<div class="footer">© 2026 MELKO Logistics</div>

<script>
['zone1','zone2'].forEach(id=>{
const z=document.getElementById(id);
z.addEventListener('dragover',e=>{e.preventDefault();z.classList.add('drag-over')});
z.addEventListener('dragleave',()=>z.classList.remove('drag-over'));
z.addEventListener('drop',e=>{
e.preventDefault();z.classList.remove('drag-over');
const f=e.dataTransfer.files[0];if(!f)return;
const inp=z.querySelector('input');const dt=new DataTransfer();dt.items.add(f);inp.files=dt.files;
inp.dispatchEvent(new Event('change'));
});
});
function setStatus(id,t,m){const e=document.getElementById(id);e.className='status '+t;e.textContent=m}
async function handleStep1(inp){
const f=inp.files[0];if(!f)return;
document.getElementById('fname1').textContent=f.name;
setStatus('status1','loading','⏳ 正在处理...');
const fd=new FormData();fd.append('file',f);
try{
const r=await fetch('/api/step1',{method:'POST',body:fd});
if(!r.ok){const e=await r.json();throw new Error(e.error||'Error')}
const b=await r.blob(),cd=r.headers.get('Content-Disposition')||'',
mt=cd.match(/filename="?([^"]+)"?/),fn=mt?mt[1]:'packing_list.xlsx';
const u=URL.createObjectURL(b),a=document.createElement('a');
a.href=u;a.download=decodeURIComponent(fn);a.click();URL.revokeObjectURL(u);
setStatus('status1','success','✅ 已生成: '+decodeURIComponent(fn)+' — 请检查后进入 Step 3');
}catch(e){setStatus('status1','error','❌ '+e.message)}
inp.value='';
}
async function handleStep2(inp){
const f=inp.files[0];if(!f)return;
document.getElementById('fname2').textContent=f.name;
setStatus('status2','loading','⏳ 正在生成 BOL 文件包...');
const fd=new FormData();fd.append('file',f);
try{
const r=await fetch('/api/step2',{method:'POST',body:fd});
if(!r.ok){const e=await r.json();throw new Error(e.error||'Error')}
const b=await r.blob(),cd=r.headers.get('Content-Disposition')||'',
mt=cd.match(/filename="?([^"]+)"?/),fn=mt?mt[1]:'BOL.zip';
const u=URL.createObjectURL(b),a=document.createElement('a');
a.href=u;a.download=decodeURIComponent(fn);a.click();URL.revokeObjectURL(u);
setStatus('status2','success','✅ BOL 文件包已下载: '+decodeURIComponent(fn));
}catch(e){setStatus('status2','error','❌ '+e.message)}
inp.value='';
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
