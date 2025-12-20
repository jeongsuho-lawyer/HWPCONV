"""
로컬 웹 서버 - HWP 변환기 GUI

python -m hwpconv.server 로 실행
"""

import os
import sys
import tempfile
import webbrowser
from pathlib import Path
from typing import Optional

try:
    from flask import Flask, request, jsonify, send_file, render_template_string
except ImportError:
    print("Flask가 필요합니다. 설치 중...")
    import subprocess
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "flask"])
        from flask import Flask, request, jsonify, send_file, render_template_string
    except Exception as e:
        print(f"Flask 설치 실패: {e}")
        print("수동으로 설치하세요: pip install flask")
        sys.exit(1)

from .parsers.hwpx import HwpxParser
from .parsers.hwp import HwpParser
from .converters.markdown import MarkdownConverter
from .converters.html import HtmlConverter

app = Flask(__name__)

# CORS 헤더 추가 (로컬 개발 및 다른 포트 접근 허용)
@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

# 404 핸들러
@app.errorhandler(404)
def not_found(e):
    return jsonify({'success': False, 'error': 'Not found'}), 404

# 임시 디렉토리
TEMP_DIR = Path(tempfile.gettempdir()) / "hwpconv"
TEMP_DIR.mkdir(exist_ok=True)

# HTML 템플릿
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HWP 변환기 - hwpconv</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        :root {
            --primary: #8B5CF6;
            --primary-light: #A78BFA;
            --bg-dark: #1a1a2e;
            --bg-card: #ffffff;
            --text-primary: #1f2937;
            --text-secondary: #6b7280;
            --border: #e5e7eb;
            --success: #10b981;
            --error: #ef4444;
        }
        body {
            font-family: 'Inter', -apple-system, sans-serif;
            background: linear-gradient(135deg, var(--bg-dark) 0%, #16213e 100%);
            min-height: 100vh;
        }
        .header {
            background: linear-gradient(90deg, #8B5CF6, #6366F1);
            padding: 1rem 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .logo {
            color: white;
            font-weight: 700;
            font-size: 1.25rem;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }
        .main-container {
            max-width: 800px;
            margin: 0 auto;
            padding: 2rem;
        }
        .format-info {
            color: white;
            margin-bottom: 1.5rem;
            font-size: 0.9rem;
        }
        .format-info .arrow { color: var(--primary-light); }
        .format-info .active { font-weight: 600; }
        .card {
            background: var(--bg-card);
            border-radius: 16px;
            padding: 2rem;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
        }
        .format-section { margin-bottom: 2rem; }
        .format-label {
            font-size: 0.875rem;
            color: var(--text-secondary);
            margin-bottom: 0.75rem;
        }
        .format-options { display: flex; gap: 1rem; }
        .format-option {
            flex: 1;
            padding: 1rem;
            border: 2px solid var(--border);
            border-radius: 12px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .format-option:hover { border-color: var(--primary-light); }
        .format-option.selected {
            border-color: var(--primary);
            background: rgba(139, 92, 246, 0.05);
        }
        .format-option .format-name { font-weight: 600; }
        .format-option .format-ext { font-size: 0.75rem; color: var(--text-secondary); }
        .format-option .check-icon { display: none; }
        .format-option.selected .check-icon { display: inline; color: var(--primary); }
        .drop-zone {
            border: 2px dashed var(--border);
            border-radius: 16px;
            padding: 3rem 2rem;
            text-align: center;
            transition: all 0.3s;
            cursor: pointer;
            margin-bottom: 2rem;
        }
        .drop-zone:hover, .drop-zone.drag-over {
            border-color: var(--primary);
            background: rgba(139, 92, 246, 0.05);
        }
        .drop-zone-icon { width: 48px; height: 48px; margin: 0 auto 1rem; color: var(--text-secondary); }
        .hidden-input { display: none; }
        .file-list-header {
            display: none;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
            padding-bottom: 0.75rem;
            border-bottom: 1px solid var(--border);
        }
        .file-list-header.show { display: flex; }
        .file-count { font-size: 0.875rem; color: var(--text-secondary); }
        .file-count strong { color: var(--primary); }
        .file-actions { display: flex; gap: 1rem; }
        .file-action-btn {
            font-size: 0.875rem;
            color: var(--text-secondary);
            background: none;
            border: none;
            cursor: pointer;
            padding: 0.5rem 0.75rem;
            border-radius: 8px;
            transition: all 0.2s;
        }
        .file-action-btn:hover { background: rgba(139, 92, 246, 0.1); color: var(--primary); }
        .file-list { list-style: none; }
        .file-item {
            display: flex;
            align-items: center;
            padding: 0.875rem 0;
            border-bottom: 1px solid var(--border);
            gap: 0.75rem;
        }
        .file-item:last-child { border-bottom: none; }
        .file-status {
            width: 20px; height: 20px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
            font-size: 10px;
        }
        .file-status.pending { background: var(--border); }
        .file-status.processing { background: var(--primary-light); animation: pulse 1.5s infinite; }
        .file-status.success { background: var(--success); color: white; }
        .file-status.error { background: var(--error); color: white; }
        @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.5; } }
        .file-info { flex: 1; min-width: 0; }
        .file-name { font-size: 0.875rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .file-conversion { font-size: 0.75rem; color: var(--text-secondary); }
        .file-conversion .arrow { color: var(--primary); }
        .file-buttons { display: flex; gap: 0.5rem; }
        .icon-btn {
            width: 36px; height: 36px;
            border: none;
            background: rgba(139, 92, 246, 0.1);
            border-radius: 8px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            color: var(--primary);
            transition: all 0.2s;
        }
        .icon-btn:hover { background: var(--primary); color: white; }
        .icon-btn.delete { color: var(--error); background: rgba(239, 68, 68, 0.1); }
        .icon-btn.delete:hover { background: var(--error); color: white; }
        .icon-btn:disabled { opacity: 0.5; cursor: not-allowed; }
        .convert-section { margin-top: 1.5rem; display: none; }
        .convert-section.show { display: block; }
        .convert-btn {
            width: 100%;
            padding: 1rem;
            background: linear-gradient(90deg, #8B5CF6, #6366F1);
            color: white;
            border: none;
            border-radius: 12px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }
        .convert-btn:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(139, 92, 246, 0.3); }
        .convert-btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
        .footer { text-align: center; padding: 2rem; color: rgba(255,255,255,0.5); font-size: 0.875rem; }
        .footer a { color: var(--primary-light); text-decoration: none; }
    </style>
</head>
<body>
    <header class="header">
        <div class="logo">📄 hwpconv</div>
        <div style="color: rgba(255,255,255,0.7); font-size: 0.875rem;">로컬 변환기</div>
    </header>

    <main class="main-container">
        <div class="format-info">
            <span>.hwp(x)</span>
            <span class="arrow">→</span>
            <span class="active" id="formatDisplay">.md</span>
        </div>

        <div class="card">
            <div class="format-section">
                <div class="format-label">변환 형식</div>
                <div class="format-options">
                    <div class="format-option selected" data-format="md" onclick="selectFormat('md')">
                        <span class="format-name">Markdown</span>
                        <span class="check-icon">✓</span>
                        <div class="format-ext">.md</div>
                    </div>
                    <div class="format-option" data-format="html" onclick="selectFormat('html')">
                        <span class="format-name">HTML</span>
                        <span class="check-icon">✓</span>
                        <div class="format-ext">.html</div>
                    </div>
                </div>
            </div>

            <div class="drop-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
                <svg class="drop-zone-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="17 8 12 3 7 8"></polyline>
                    <line x1="12" y1="3" x2="12" y2="15"></line>
                </svg>
                <div>파일 또는 폴더를 드래그</div>
                <div style="font-size: 0.875rem; color: #6b7280;">.hwp / .hwpx</div>
                <input type="file" class="hidden-input" id="fileInput" multiple accept=".hwp,.hwpx">
            </div>

            <div class="file-list-header" id="fileListHeader">
                <div class="file-count">
                    <span id="fileCountNum">0</span>개 파일 &nbsp; <strong>✓ <span id="completedCount">0</span></strong>
                </div>
                <div class="file-actions">
                    <button class="file-action-btn" onclick="clearAll()">지우기</button>
                    <button class="file-action-btn" onclick="downloadAll()">📥 전체 다운로드</button>
                </div>
            </div>

            <ul class="file-list" id="fileList"></ul>

            <div class="convert-section" id="convertSection">
                <button class="convert-btn" id="convertBtn" onclick="convertFiles()">🚀 변환 시작</button>
            </div>
        </div>
    </main>

    <footer class="footer">
        <p>hwpconv v0.1.0 | <a href="https://github.com/rnlaw/hwpconv">GitHub</a></p>
    </footer>

    <script>
        let files = [];
        let selectedFormat = 'md';
        let convertedFiles = {};

        function selectFormat(format) {
            selectedFormat = format;
            document.querySelectorAll('.format-option').forEach(o => {
                o.classList.toggle('selected', o.dataset.format === format);
            });
            document.getElementById('formatDisplay').textContent = '.' + format;
            updateFileList();
        }

        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');

        dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
        dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('drag-over'); handleFiles(e.dataTransfer.files); });
        fileInput.addEventListener('change', e => { handleFiles(e.target.files); fileInput.value = ''; });

        function handleFiles(fileList) {
            Array.from(fileList).filter(f => f.name.match(/\\.hwpx?$/i)).forEach(file => {
                if (!files.some(f => f.name === file.name)) files.push(file);
            });
            updateFileList();
        }

        function updateFileList() {
            const list = document.getElementById('fileList');
            list.innerHTML = '';
            
            files.forEach((file, i) => {
                const status = convertedFiles[file.name] ? 'success' : 'pending';
                const outName = file.name.replace(/\\.(hwp|hwpx)$/i, '.' + selectedFormat);
                
                list.innerHTML += `
                    <li class="file-item">
                        <div class="file-status ${status}">${status === 'success' ? '✓' : ''}</div>
                        <div class="file-info">
                            <div class="file-name">${file.name}</div>
                            <div class="file-conversion">${file.name} <span class="arrow">→</span> ${outName}</div>
                        </div>
                        <div class="file-buttons">
                            <button class="icon-btn" onclick="downloadFile(${i})" ${!convertedFiles[file.name] ? 'disabled' : ''}>📥</button>
                            <button class="icon-btn delete" onclick="removeFile(${i})">✕</button>
                        </div>
                    </li>`;
            });
            
            document.getElementById('fileListHeader').classList.toggle('show', files.length > 0);
            document.getElementById('convertSection').classList.toggle('show', files.length > 0);
            document.getElementById('fileCountNum').textContent = files.length;
            document.getElementById('completedCount').textContent = Object.keys(convertedFiles).length;
        }

        function removeFile(i) {
            delete convertedFiles[files[i].name];
            files.splice(i, 1);
            updateFileList();
        }

        function clearAll() {
            files = [];
            convertedFiles = {};
            updateFileList();
        }

        async function convertFiles() {
            const btn = document.getElementById('convertBtn');
            btn.disabled = true;
            btn.textContent = '⏳ 변환 중...';

            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                if (convertedFiles[file.name]) continue;

                const items = document.querySelectorAll('.file-item');
                items[i].querySelector('.file-status').className = 'file-status processing';

                const formData = new FormData();
                formData.append('file', file);
                formData.append('format', selectedFormat);

                try {
                    const res = await fetch('/convert', { method: 'POST', body: formData });
                    const data = await res.json();
                    
                    if (data.success) {
                        convertedFiles[file.name] = { content: data.content, format: selectedFormat };
                        items[i].querySelector('.file-status').className = 'file-status success';
                        items[i].querySelector('.file-status').textContent = '✓';
                    } else {
                        items[i].querySelector('.file-status').className = 'file-status error';
                        items[i].querySelector('.file-status').textContent = '!';
                    }
                } catch (e) {
                    items[i].querySelector('.file-status').className = 'file-status error';
                }
                
                document.getElementById('completedCount').textContent = Object.keys(convertedFiles).length;
            }

            btn.disabled = false;
            btn.textContent = '✅ 변환 완료!';
            setTimeout(() => { btn.textContent = '🚀 변환 시작'; updateFileList(); }, 2000);
        }

        function downloadFile(i) {
            const file = files[i];
            const converted = convertedFiles[file.name];
            if (!converted) return;
            
            const outName = file.name.replace(/\\.(hwp|hwpx)$/i, '.' + converted.format);
            const blob = new Blob([converted.content], { type: 'text/plain;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = outName;
            a.click();
        }

        function downloadAll() {
            files.forEach((f, i) => { if (convertedFiles[f.name]) downloadFile(i); });
        }
    </script>
</body>
</html>
'''


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/convert', methods=['POST'])
def convert():
    """파일 변환 API"""
    try:
        file = request.files.get('file')
        output_format = request.form.get('format', 'md')
        
        if not file:
            return jsonify({'success': False, 'error': 'No file provided'})
        
        # 고유한 임시 파일로 저장 (경쟁 조건 방지)
        import uuid
        filename = file.filename
        unique_name = f"{uuid.uuid4().hex}_{filename}"
        temp_path = TEMP_DIR / unique_name
        file.save(str(temp_path))
        
        try:
            # 파서 선택
            ext = Path(filename).suffix.lower()
            if ext == '.hwpx':
                doc = HwpxParser().parse(str(temp_path))
            elif ext == '.hwp':
                doc = HwpParser().parse(str(temp_path))
            else:
                return jsonify({'success': False, 'error': f'Unsupported format: {ext}'})
            
            # 변환
            if output_format == 'html':
                content = HtmlConverter().convert(doc)
            else:
                content = MarkdownConverter().convert(doc)
            
            return jsonify({'success': True, 'content': content})
            
        finally:
            # 임시 파일 삭제
            if temp_path.exists():
                temp_path.unlink()
                
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


def main(port: int = 5000, open_browser: bool = True):
    """서버 시작"""
    print(f"\n{'='*50}")
    print(f"  hwpconv 로컬 변환기")
    print(f"  http://localhost:{port}")
    print(f"{'='*50}\n")
    
    if open_browser:
        import threading
        threading.Timer(1.0, lambda: webbrowser.open(f'http://localhost:{port}')).start()
    
    app.run(host='127.0.0.1', port=port, debug=False)


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='HWP 변환기 로컬 서버')
    parser.add_argument('-p', '--port', type=int, default=5000, help='포트 번호')
    parser.add_argument('--no-browser', action='store_true', help='브라우저 자동 열기 비활성화')
    args = parser.parse_args()
    
    main(port=args.port, open_browser=not args.no_browser)
