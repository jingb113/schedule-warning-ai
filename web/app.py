# -*- coding: utf-8 -*-
"""
工期预警管控系统 - Flask Web 后端
提供文件上传、比对执行、邮件发送的 Web 服务
"""

from flask import Flask, request, jsonify, render_template
import os
from datetime import datetime
from schedule_warning_core import run_comparison

app = Flask(__name__)

# CORS：允许所有来源，仅适用于本机/内网自用；
# 若要暴露到公网，请改为固定白名单并加上鉴权，否则任何人都能上传文件并触发邮件
@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
# 只监听本机，避免被同网段其他机器直接访问
WEB_HOST = os.environ.get('WEB_HOST', '127.0.0.1')
WEB_PORT = int(os.environ.get('WEB_PORT', '8080'))

# 进度计划文件固定保存路径（只需上传一次）
PLAN_FILE_PATH = os.path.join(app.config['UPLOAD_FOLDER'], 'current_plan.xlsx')

# 确保目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


@app.route('/')
def index():
    """首页 - 返回前端页面"""
    return render_template('index.html')


@app.route('/api/status')
def status():
    """获取当前系统状态：进度计划是否已配置"""
    plan_exists = os.path.exists(PLAN_FILE_PATH)
    plan_info = None
    if plan_exists:
        mtime = os.path.getmtime(PLAN_FILE_PATH)
        plan_info = {
            'uploaded': True,
            'filename': 'current_plan.xlsx',
            'upload_time': datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
        }
    else:
        plan_info = {'uploaded': False}
    
    return jsonify({
        'plan': plan_info
    })


@app.route('/api/upload_plan', methods=['POST'])
def upload_plan():
    """上传/更新进度计划文件（只需一次）"""
    try:
        if 'plan_file' not in request.files:
            return jsonify({'success': False, 'error': '缺少进度计划文件'}), 400
        
        plan_file = request.files['plan_file']
        if plan_file.filename == '':
            return jsonify({'success': False, 'error': '文件名为空'}), 400
        
        plan_file.save(PLAN_FILE_PATH)
        
        return jsonify({
            'success': True,
            'message': '进度计划已保存，后续无需重复上传',
            'filename': plan_file.filename
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/compare', methods=['POST'])
def compare():
    """
    比对接口
    接收：日报文件（进度计划使用已保存的版本）
    返回：比对结果 JSON
    """
    # 调试日志（走标准日志，不落盘到工作目录）
    app.logger.debug('/api/compare called, files=%s, plan_exists=%s',
                     list(request.files.keys()), os.path.exists(PLAN_FILE_PATH))
    try:
        # 检查进度计划是否存在
        if not os.path.exists(PLAN_FILE_PATH):
            return jsonify({
                'success': False,
                'error': '尚未上传进度计划文件，请先上传'
            }), 400
        
        # 检查日报文件
        if 'daily_file' not in request.files:
            return jsonify({
                'success': False,
                'error': '缺少日报文件'
            }), 400
        
        daily_file = request.files['daily_file']
        if daily_file.filename == '':
            return jsonify({
                'success': False,
                'error': '日报文件名为空'
            }), 400
        
        # 保存日报文件
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        daily_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{timestamp}_daily.xlsx')
        daily_file.save(daily_path)
        
        # 执行比对（进度计划用已保存的固定文件）
        result = run_comparison(
            plan_file=PLAN_FILE_PATH,
            daily_file=daily_path,
            output_dir=app.config['OUTPUT_FOLDER'],
            send_email=True  # 自动发送邮件
        )
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'服务器错误: {str(e)}'
        }), 500


@app.route('/api/health')
def health():
    """健康检查接口"""
    return jsonify({
        'status': 'ok',
        'timestamp': datetime.now().isoformat()
    })


if __name__ == '__main__':
    print('工期预警管控系统 Web 服务')
    print(f'访问地址: http://{WEB_HOST}:{WEB_PORT}')
    app.run(host=WEB_HOST, port=WEB_PORT, debug=False, use_reloader=False)
