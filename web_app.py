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
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Noto Sans SC',-apple-system,BlinkMacSystemFont,'Helvetica Neue',sans-serif;background:#f5f5f7;color:#1d1d1f;min-height:100vh;-webkit-font-smoothing:antialiased;background-image:url('/static/bg.jpg');background-size:cover;background-position:center;background-attachment:fixed;background-repeat:no-repeat}
body::before{content:'';position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(255,255,255,0.50);z-index:0;pointer-events:none}
body>*{position:relative;z-index:1}
.header{background:#fbfbfd;border-bottom:1px solid #d2d2d7;padding:14px 0}
.header-inner{max-width:1100px;margin:0 auto;padding:0 24px;display:flex;align-items:center;gap:16px}
.header img{height:60px}
.header-text h1{font-size:24px;font-weight:600;color:#1d1d1f}
.header-text p{font-size:14px;color:#86868b;margin-top:2px}
.page-hero{max-width:1100px;margin:0 auto;padding:52px 24px 16px;text-align:center}
.page-hero h2{font-size:40px;font-weight:600;color:#1d1d1f;letter-spacing:-0.5px}
.page-hero p{font-size:17px;color:#86868b;margin-top:8px}
.steps-row{max-width:1100px;margin:32px auto 0;padding:0 24px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px}
@media(max-width:900px){.steps-row{grid-template-columns:1fr}}
.step-card{background:#fbfbfd;border-radius:18px;overflow:hidden;display:flex;flex-direction:column;transition:transform 0.2s}
.step-card:hover{transform:translateY(-3px)}
.step-top{padding:20px 24px;display:flex;align-items:center;gap:12px}
.step-num{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:600;color:#fff;flex-shrink:0}
.step-num.blue{background:#3454d1}
.step-num.amber{background:#ff9500}
.step-num.green{background:#34c759}
.step-top h3{font-size:17px;font-weight:600;color:#1d1d1f}
.step-body{padding:0 24px 28px;flex:1;display:flex;flex-direction:column;align-items:center;text-align:center}
.step-body .desc{font-size:14px;color:#86868b;line-height:1.6;margin-bottom:20px}
.upload-zone{width:100%;border:1.5px dashed #d2d2d7;border-radius:12px;padding:28px 16px;cursor:pointer;transition:all 0.2s;background:#fff}
.upload-zone:hover,.upload-zone.drag-over{border-color:#3454d1;background:#f5f8ff}
.upload-zone input{display:none}
.upload-zone .uz-icon{font-size:32px;margin-bottom:6px}
.upload-zone .uz-text{font-size:14px;font-weight:500;color:#1d1d1f}
.upload-zone .uz-sub{font-size:12px;color:#86868b;margin-top:3px}
.upload-zone .browse{display:inline-block;margin-top:14px;padding:9px 22px;background:#3454d1;color:#fff;border:none;border-radius:980px;font-size:14px;font-weight:400;cursor:pointer;transition:background 0.15s}
.upload-zone .browse:hover{background:#2840a0}
.fname{font-size:12px;color:#3454d1;margin-top:8px;font-weight:500;word-break:break-all}
.status{margin-top:10px;padding:10px 14px;border-radius:10px;font-size:13px;display:none;text-align:left;width:100%}
.status.success{display:block;background:#f0faf0;color:#248a3d}
.status.error{display:block;background:#fff0f0;color:#d70015}
.status.loading{display:block;background:#f5f8ff;color:#3454d1}
.review-content{display:flex;flex-direction:column;align-items:center;gap:12px;flex:1;justify-content:center;padding:8px 0}
.review-content .r-icon{font-size:40px}
.review-content .r-text{font-size:14px;color:#86868b;line-height:1.6}
.review-content .r-text strong{color:#1d1d1f;font-weight:600}
.review-content .r-hint{font-size:13px;color:#86868b;padding:8px 16px;background:#f5f5f7;border-radius:10px}
.features-bar{max-width:1100px;margin:40px auto 0;padding:0 24px}
.features-inner{background:#3454d1;border-radius:18px;padding:28px 32px;display:flex;flex-wrap:wrap;gap:12px;justify-content:center;align-items:center}
.features-inner h3{width:100%;text-align:center;font-size:17px;font-weight:600;color:#fff;margin-bottom:6px}
.feat-pill{display:flex;align-items:center;gap:8px;background:rgba(255,255,255,0.08);padding:8px 16px;border-radius:980px}
.feat-pill .fi{font-size:16px}
.feat-pill .ft{font-size:13px;color:rgba(255,255,255,0.85)}
.bol-bar{max-width:1100px;margin:16px auto 48px;padding:0 24px}
.bol-inner{background:#fbfbfd;border-radius:18px;padding:20px 32px;display:flex;flex-wrap:wrap;gap:10px;justify-content:center;align-items:center}
.bol-inner h3{width:100%;text-align:center;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1.5px;color:#86868b;margin-bottom:2px}
.bol-tag{font-size:13px;color:#1d1d1f;background:#fff;padding:7px 14px;border-radius:10px;border:1px solid #d2d2d7}
.bol-tag strong{color:#3454d1}
.footer{text-align:center;padding:28px;font-size:12px;color:#86868b}
</style>
</head>
<body>
<!-- Header -->
<header class="header">
<div class="header-inner">
<img src="/static/logo.png" alt="MELKO">
<div class="header-text">
<h1>Packing List Maker</h1>
<p>Packing List & BOL Generator</p>
</div>
</div>
</header>

<!-- Hero -->
<div class="page-hero">
<h2>极速 3 步自动化流程</h2>
<p>上传客户的预报表 Excel，自动生成标准 Packing List 和 BOL 文件包。</p>
</div>

<!-- 3 Steps Horizontal -->
<div class="steps-row">
<!-- Step 1 -->
<div class="step-card">
<div class="step-top">
<div class="step-num blue">1</div>
<h3>上传源文件生成方案</h3>
</div>
<div class="step-body">
<div class="desc">拖拽客户 Excel 文件，系统自动完成 KG/LBS 转换、颜色标记及仓库代码合并。</div>
<div class="upload-zone" id="zone1" onclick="document.getElementById('file1').click()">
<input type="file" id="file1" accept=".xlsx,.xls" onchange="handleStep1(this)">
<div class="uz-icon">📄</div>
<div class="uz-text">拖拽客户 Excel 到此处</div>
<div class="uz-sub">支持 .XLSX / .XLS 格式</div>
<button class="browse" onclick="event.stopPropagation();document.getElementById('file1').click()">浏览文件</button>
<div class="fname" id="fname1"></div>
</div>
<div class="status" id="status1"></div>
</div>
</div>

<!-- Step 2 -->
<div class="step-card">
<div class="step-top">
<div class="step-num amber">2</div>
<h3>人工检查并保存</h3>
</div>
<div class="step-body">
<div class="review-content">
<div class="r-icon">🔍</div>
<div class="r-text">打开生成的列表，<strong>核对颜色、分组及件数</strong>等关键数据，修改后本地保存。</div>
<div class="r-hint">💾 检查完毕后进入 Step 3</div>
</div>
</div>
</div>

<!-- Step 3 -->
<div class="step-card">
<div class="step-top">
<div class="step-num green">3</div>
<h3>上传获取 BOL 文件包</h3>
</div>
<div class="step-body">
<div class="desc">上传核对后的文件，系统自动生成包含装箱单、Hold单、汇总表及唛头的 ZIP 包。</div>
<div class="upload-zone" id="zone2" onclick="document.getElementById('file2').click()">
<input type="file" id="file2" accept=".xlsx" onchange="handleStep2(this)">
<div class="uz-icon">📦</div>
<div class="uz-text">拖拽 Packing List 到此处</div>
<div class="uz-sub">自动生成 BOL 文件包 (ZIP)</div>
<button class="browse" onclick="event.stopPropagation();document.getElementById('file2').click()">浏览文件</button>
<div class="fname" id="fname2"></div>
</div>
<div class="status" id="status2"></div>
</div>
</div>
</div>

<!-- Features Bar -->
<div class="features-bar">
<div class="features-inner">
<h3>自动化功能集成，告别手动纠错</h3>
<div class="feat-pill"><span class="fi">🚛</span><span class="ft">柜号自动识别</span></div>
<div class="feat-pill"><span class="fi">🏷️</span><span class="ft">FBA 代码连续值合并</span></div>
<div class="feat-pill"><span class="fi">🎨</span><span class="ft">物流专用颜色标记</span></div>
<div class="feat-pill"><span class="fi">⚖️</span><span class="ft">KG → LBS 自动换算</span></div>
<div class="feat-pill"><span class="fi">📋</span><span class="ft">支持 .xls + .xlsx</span></div>
</div>
</div>

<!-- BOL Contents -->
<div class="bol-bar">
<div class="bol-inner">
<h3>BOL 文件包含</h3>
<div class="bol-tag"><strong>01</strong> Packing List (A-M 13列)</div>
<div class="bol-tag"><strong>02</strong> Hold List (HOLD/RELABEL)</div>
<div class="bol-tag"><strong>03</strong> TRUCKING.docx (板数/CBM)</div>
<div class="bol-tag"><strong>04</strong> 柜号唛头.docx (柜号+MELKO)</div>
</div>
</div>

<div class="footer">&copy; 2026 MELKO Logistics</div>

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
setStatus('status1','success','✅ 已生成 '+decodeURIComponent(fn)+' — 请检查后进入 Step 3');
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
