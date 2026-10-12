# -*- coding: utf-8 -*-
"""
工期预警管控系统 - 核心逻辑模块（Web 版，规则分析，不含大模型）
可被 Flask 后端调用，也可独立运行
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, date
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import glob
import re
import json


def _load_dotenv():
    """从本文件所在目录（或向上最多 3 层）查找 .env 并载入 os.environ。
    无第三方依赖；已存在的系统环境变量优先，不被 .env 覆盖。"""
    here = os.path.dirname(os.path.abspath(__file__))
    for _ in range(4):
        cand = os.path.join(here, '.env')
        if os.path.isfile(cand):
            try:
                with open(cand, encoding='utf-8') as fh:
                    for line in fh:
                        line = line.strip()
                        if not line or line.startswith('#') or '=' not in line:
                            continue
                        k, _, v = line.partition('=')
                        k, v = k.strip(), v.strip().strip('"').strip("'")
                        if k:
                            os.environ.setdefault(k, v)
            except OSError:
                pass
            return
        parent = os.path.dirname(here)
        if parent == here:
            break
        here = parent


_load_dotenv()

# ============================================================
# 配置区
# ============================================================
# 项目名（用于邮件标题），可用环境变量覆盖
PROJECT_NAME = os.environ.get('PROJECT_NAME', '示例项目')
# 日报文件名需包含的关键词（逗号分隔），用于自动定位日报
DAILY_FILE_KEYWORDS = [k.strip() for k in os.environ.get('DAILY_FILE_KEYWORDS', '工程日报').split(',') if k.strip()]
# 总进度计划工作表名候选，命中第一个即用；都不命中则取第一个工作表
PLAN_SHEET_CANDIDATES = [s for s in os.environ.get('PLAN_SHEET_NAMES', '').split(',') if s.strip()]

# 主任务匹配：(日报行号, 日报行名, 进度计划任务名) —— 请按自己项目的日报结构填写
MAPPING = [
    # (34, '示例日报行名', '示例计划任务名'),
]

# 子任务归类：子任务名关键词 → 所属主任务分类 —— 请按自己项目填写
SUBTASK_CATEGORY = {
    # '示例子任务名': '示例所属分类',
}

GREEN_THRESHOLD = -0.05
YELLOW_THRESHOLD = -0.10
TIME_WARN_RED = 7
TIME_WARN_YELLOW = 3

# ------------------------------------------------------------
# 邮件配置（全部通过环境变量提供，仓库内不含任何凭证）
#   SMTP_SERVER / SMTP_PORT / SMTP_USER / SMTP_PASS / EMAIL_RECIPIENTS
# SMTP_PASS 为邮箱「授权码」而非登录密码（163/QQ/126 等均需单独申请）
# 未配置 SMTP_USER / SMTP_PASS / 收件人时自动跳过发邮件，只生成报表
# ------------------------------------------------------------
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.163.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '465'))
SMTP_USER = os.environ.get('SMTP_USER', '')
SMTP_PASS = os.environ.get('SMTP_PASS', '')
EMAIL_RECIPIENTS = [x.strip() for x in os.environ.get('EMAIL_RECIPIENTS', '').split(',') if x.strip()]


# ============================================================
# 工具函数
# ============================================================
def parse_date(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        for fmt in ['%Y年%m月%d日', '%Y-%m-%d', '%Y/%m/%d', '%Y/%m%d']:
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except:
                pass
    return None


def parse_pct(val):
    if val is None or val == '\\' or val == '':
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        if '完成' in val:
            return 1.0
        elif '%' in val:
            try:
                return float(val.replace('%', '')) / 100
            except:
                return 0.0
    return 0.0


def warning_level(deviation):
    if deviation < YELLOW_THRESHOLD:
        return '红色'
    elif deviation < GREEN_THRESHOLD:
        return '黄色'
    else:
        return '绿色'


def classify_subtask(name):
    if name is None:
        return None
    name = str(name).strip()
    for keyword, category in SUBTASK_CATEGORY.items():
        if keyword in name:
            return category
    return None


def time_warning_level(delay_days):
    if delay_days >= TIME_WARN_RED:
        return '红色'
    elif delay_days >= TIME_WARN_YELLOW:
        return '黄色'
    else:
        return '绿色'


def find_latest_daily_file(directory):
    pattern = os.path.join(directory, '*.xlsx')
    files = [f for f in glob.glob(pattern)
             if not os.path.basename(f).startswith('~')
             and any(kw in os.path.basename(f) for kw in DAILY_FILE_KEYWORDS)]
    if not files:
        # 未命中关键词：回退到目录下全部 xlsx（排除临时文件）
        files = [f for f in glob.glob(pattern)
                 if not os.path.basename(f).startswith('~')]
    if not files:
        return None
    files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return files[0]


# ============================================================
# 核心比对函数
# ============================================================
def run_comparison(plan_file, daily_file, output_dir, send_email=True, today=None):
    """
    执行完整的工期比对分析
    
    Args:
        plan_file: 进度计划文件路径
        daily_file: 日报文件路径
        output_dir: 输出目录
        send_email: 是否发送邮件
        today: 指定日期，默认为今天
    
    Returns:
        dict: 包含比对结果、统计数据、文件路径
    """
    if today is None:
        today = datetime.now().date()
    
    # 前置校验：Web 版基于 MAPPING / SUBTASK_CATEGORY 固定配置解析，
    # 开源版默认置空，需按自己项目的日报结构填写，否则会返回空结果
    if not MAPPING:
        return {
            'success': False,
            'error': ('尚未配置主任务匹配表 MAPPING。Web 版按固定行号解析日报，'
                      '请打开 schedule_warning_core.py，在 MAPPING 中填入'
                      '（日报行号, 日报行名, 进度计划任务名）的对应关系，'
                      '并按需在 SUBTASK_CATEGORY 中配置子任务归类。'),
            'summary': [], 'subtasks': [],
            'stats': {'summary_counts': {'red': 0, 'yellow': 0, 'green': 0},
                      'subtask_counts': {'red': 0, 'yellow': 0, 'green': 0, 'none': 0}},
            'output_file': None, 'email_sent': False, 'email_recipients': [],
        }
    
    result = {
        'success': False,
        'error': None,
        'summary': [],
        'subtasks': [],
        'stats': {
            'summary_counts': {'red': 0, 'yellow': 0, 'green': 0},
            'subtask_counts': {'red': 0, 'yellow': 0, 'green': 0, 'none': 0}
        },
        'output_file': None,
        'email_sent': False
    }
    
    try:
        # 步骤1：读取进度计划
        # 注意：本模块按旧版计划表列位读取（任务名=第5列、开始=第7列、完成=第8列），
        # 工作表名可通过环境变量 PLAN_SHEET_NAMES 指定，否则取第一个工作表
        wb_plan = openpyxl.load_workbook(plan_file, data_only=True)
        ws_plan = None
        for cand in PLAN_SHEET_CANDIDATES:
            if cand in wb_plan.sheetnames:
                ws_plan = wb_plan[cand]
                break
        if ws_plan is None:
            ws_plan = wb_plan.worksheets[0]
        
        plan_dict = {}
        for row in ws_plan.iter_rows(min_row=2, max_row=ws_plan.max_row, values_only=True):
            name = row[4]
            if not name:
                continue
            name = str(name).strip()
            start = parse_date(row[6])
            end = parse_date(row[7])
            if start and end:
                dur_val = row[5]
                if isinstance(dur_val, str):
                    dur_days = int(re.sub(r'[^\d]', '', dur_val) or '0')
                elif isinstance(dur_val, (int, float)):
                    dur_days = int(dur_val)
                else:
                    dur_days = (end - start).days
                plan_dict[name] = {
                    'name': name, 'start': start, 'end': end,
                    'duration': max(dur_days, (end - start).days),
                }
        
        # 步骤2：读取日报
        wb_daily = openpyxl.load_workbook(daily_file, data_only=True)
        ws_daily = wb_daily['1'] if '1' in wb_daily.sheetnames else wb_daily.worksheets[0]
        
        # 汇总行
        daily_summary = {}
        for row_num, daily_name, plan_name_kw in MAPPING:
            row = list(ws_daily.iter_rows(min_row=row_num, max_row=row_num, values_only=True))[0]
            work_name = str(row[4]).strip() if row[4] else ''
            total_qty = row[5]
            pct = parse_pct(row[14])
            daily_summary[row_num] = {
                'name': work_name, 'total_qty': total_qty,
                'actual_pct': pct, 'plan_keyword': plan_name_kw,
            }
        
        # 子任务明细
        subtasks = []
        current_block = ''
        for row_num in range(37, 90):
            row = list(ws_daily.iter_rows(min_row=row_num, max_row=row_num, values_only=True))[0]
            
            col3 = str(row[3]).strip() if row[3] else ''
            col4 = str(row[4]).strip() if row[4] else ''
            
            if len(col3) == 1 and col3 in 'ABCDE':
                current_block = col3
                task_name = col4
            else:
                task_name = col3 if col3 else col4
            
            if not task_name or task_name in ['总体', '总体形况', '南区换热孔']:
                continue
            
            if '人员' in task_name or '管理人员' in task_name:
                break
            
            category = classify_subtask(task_name)
            if not category:
                continue
            
            unit = str(row[5]).strip() if row[5] else ''
            total_qty = row[6]
            plan_start = parse_date(row[7])
            plan_end = parse_date(row[8])
            actual_start = parse_date(row[9])
            actual_end = parse_date(row[10])
            status = str(row[11]).strip() if row[11] else ''
            pct = parse_pct(row[14])
            
            start_delay = None
            end_delay = None
            
            if plan_start and actual_start:
                start_delay = (actual_start - plan_start).days
            
            if plan_end:
                if (status == '完成' or pct >= 1.0) and actual_end:
                    end_delay = (actual_end - plan_end).days
                else:
                    end_delay = (today - plan_end).days
            
            st = {
                'block': current_block,
                'row': row_num,
                'name': task_name,
                'category': category,
                'unit': unit,
                'total_qty': total_qty,
                'plan_start': plan_start,
                'plan_end': plan_end,
                'actual_start': actual_start,
                'actual_end': actual_end,
                'status': status,
                'pct': pct,
                'start_delay': start_delay,
                'end_delay': end_delay,
            }
            subtasks.append(st)
        
        # 步骤3A：汇总比对
        summary_results = []
        for row_num, daily_name, plan_keyword in MAPPING:
            dd = daily_summary[row_num]
            plan_task = None
            for pname, pinfo in plan_dict.items():
                if plan_keyword == pname:
                    plan_task = pinfo
                    break
            
            if not plan_task:
                continue
            
            total_days = plan_task['duration']
            if total_days <= 0:
                continue
            
            elapsed = (today - plan_task['start']).days
            if elapsed < 0:
                continue
            
            plan_pct = min(elapsed / total_days, 1.0)
            actual_pct = dd['actual_pct']
            deviation = actual_pct - plan_pct
            dev_days = round(deviation * total_days)
            level = warning_level(deviation)
            
            sr = {
                'daily_name': dd['name'],
                'plan_name': plan_task['name'],
                'plan_start': plan_task['start'],
                'plan_end': plan_task['end'],
                'total_days': total_days,
                'elapsed_days': elapsed,
                'plan_pct': plan_pct,
                'actual_pct': actual_pct,
                'deviation': deviation,
                'dev_days': dev_days,
                'level': level,
            }
            summary_results.append(sr)
        
        summary_results.sort(key=lambda x: x['deviation'])
        
        # 步骤3B：子任务统计
        sub_red = sum(1 for st in subtasks if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_RED and st['pct'] < 1.0)
        sub_yellow = sum(1 for st in subtasks if st['end_delay'] is not None and TIME_WARN_YELLOW <= st['end_delay'] < TIME_WARN_RED and st['pct'] < 1.0)
        sub_green = sum(1 for st in subtasks if st['end_delay'] is not None and st['end_delay'] < TIME_WARN_YELLOW)
        sub_none = len(subtasks) - sub_red - sub_yellow - sub_green
        
        # 步骤4：生成Excel
        output_file = os.path.join(output_dir, f'工期预警报表_{today.strftime("%Y-%m-%d")}.xlsx')
        generate_excel_report(summary_results, subtasks, today, output_file, 
                             os.path.basename(plan_file), os.path.basename(daily_file))
        
        # 步骤5：发送邮件
        email_sent = False
        if send_email and summary_results:
            email_sent = send_warning_email(summary_results, subtasks, today, output_file, daily_file)
        
        # 组装结果
        result['success'] = True
        result['summary'] = [
            {
                'daily_name': r['daily_name'],
                'plan_name': r['plan_name'],
                'plan_start': r['plan_start'].strftime('%Y-%m-%d'),
                'plan_end': r['plan_end'].strftime('%Y-%m-%d'),
                'total_days': r['total_days'],
                'elapsed_days': r['elapsed_days'],
                'plan_pct': round(r['plan_pct'] * 100, 1),
                'actual_pct': round(r['actual_pct'] * 100, 1),
                'deviation': round(r['deviation'] * 100, 1),
                'dev_days': r['dev_days'],
                'level': r['level']
            }
            for r in summary_results
        ]
        result['subtasks'] = [
            {
                'block': st['block'],
                'name': st['name'],
                'category': st['category'],
                'status': st['status'],
                'pct': round(st['pct'] * 100),
                'start_delay': st['start_delay'],
                'end_delay': st['end_delay'],
                'time_warning': time_warning_level(st['end_delay']) if st['end_delay'] is not None and st['pct'] < 1.0 else ('绿色' if st['status'] == '完成' or st['pct'] >= 1.0 else '-')
            }
            for st in subtasks
        ]
        result['stats'] = {
            'summary_counts': {
                'red': sum(1 for r in summary_results if r['level'] == '红色'),
                'yellow': sum(1 for r in summary_results if r['level'] == '黄色'),
                'green': sum(1 for r in summary_results if r['level'] == '绿色'),
            },
            'subtask_counts': {
                'red': sub_red,
                'yellow': sub_yellow,
                'green': sub_green,
                'none': sub_none
            }
        }
        result['output_file'] = output_file
        result['email_sent'] = email_sent
        # 收件人由后端环境变量提供，前端不再硬编码
        result['email_recipients'] = list(EMAIL_RECIPIENTS)
        
        return result
        
    except Exception as e:
        result['error'] = str(e)
        return result


def generate_excel_report(summary_results, subtasks, today, output_file, plan_filename, daily_filename):
    """生成Excel预警报表"""
    wb_out = openpyxl.Workbook()
    
    # 样式定义
    header_font = Font(name='Microsoft YaHei', bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='1A2740', end_color='1A2740', fill_type='solid')
    sub_header_fill = PatternFill(start_color='233045', end_color='233045', fill_type='solid')
    data_font = Font(name='Microsoft YaHei', size=10)
    num_font = Font(name='Consolas', size=10)
    red_fill = PatternFill(start_color='3D1520', end_color='3D1520', fill_type='solid')
    red_font = Font(name='Microsoft YaHei', size=10, color='FF4060', bold=True)
    yellow_fill = PatternFill(start_color='3D3015', end_color='3D3015', fill_type='solid')
    yellow_font = Font(name='Microsoft YaHei', size=10, color='FFAA00', bold=True)
    green_fill = PatternFill(start_color='153D20', end_color='153D20', fill_type='solid')
    green_font = Font(name='Microsoft YaHei', size=10, color='00FF88')
    white_font = Font(name='Microsoft YaHei', size=10, color='FFFFFF')
    gray_font = Font(name='Microsoft YaHei', size=10, color='8899AA')
    thin_border = Border(
        left=Side(style='thin', color='2A3A50'),
        right=Side(style='thin', color='2A3A50'),
        top=Side(style='thin', color='2A3A50'),
        bottom=Side(style='thin', color='2A3A50'),
    )
    center_align = Alignment(horizontal='center', vertical='center')
    wrap_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    summary_counts = {
        'red': sum(1 for r in summary_results if r['level'] == '红色'),
        'yellow': sum(1 for r in summary_results if r['level'] == '黄色'),
        'green': sum(1 for r in summary_results if r['level'] == '绿色'),
    }
    
    # Sheet1: 汇总比对
    ws1 = wb_out.active
    ws1.title = '汇总比对'
    ws1.sheet_properties.tabColor = '1A2740'
    
    title_row = ['日报工作项', '进度计划任务', '计划开始', '计划完成', '总工期(天)',
                 '已过天数', '计划完成%', '实际完成%', '偏差%', '偏差天数', '预警等级']
    ws1.append(title_row)
    for col in range(1, len(title_row) + 1):
        cell = ws1.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    
    for r in summary_results:
        row_data = [
            r['daily_name'], r['plan_name'],
            r['plan_start'].strftime('%Y-%m-%d'), r['plan_end'].strftime('%Y-%m-%d'),
            r['total_days'], r['elapsed_days'],
            f"{r['plan_pct']*100:.1f}%", f"{r['actual_pct']*100:.1f}%",
            f"{r['deviation']*100:+.1f}%", f"{r['dev_days']:+d}天", r['level'],
        ]
        ws1.append(row_data)
        row_idx = ws1.max_row
        
        if r['level'] == '红色':
            fill, font = red_fill, red_font
        elif r['level'] == '黄色':
            fill, font = yellow_fill, yellow_font
        else:
            fill, font = green_fill, green_font
        
        for col in range(1, len(row_data) + 1):
            cell = ws1.cell(row=row_idx, column=col)
            cell.fill = fill
            cell.border = thin_border
            cell.alignment = center_align
            cell.font = data_font if col in [1, 2] else (font if col == 11 else num_font)
    
    col_widths = [16, 28, 12, 12, 10, 10, 10, 10, 10, 10, 10]
    for i, w in enumerate(col_widths, 1):
        ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    
    # Sheet2: 子任务时间偏差
    ws2 = wb_out.create_sheet('子任务时间偏差')
    ws2.sheet_properties.tabColor = 'FF8800'
    
    sub_cols = ['所属主任务', '区段', '子任务名称', '单位', '总量',
                '计划开始', '计划结束', '实际开始', '实际结束', '状态',
                '开始偏差(天)', '结束偏差(天)', '完成%', '时间预警']
    ws2.append(sub_cols)
    for col in range(1, len(sub_cols) + 1):
        cell = ws2.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    
    sorted_subtasks = sorted(subtasks, key=lambda x: (
        {'CFG桩施工': 0, '抗浮锚杆施工': 1, '桩基础施工段': 2}.get(x['category'], 99),
        x.get('block', 'Z')
    ))
    
    last_category = ''
    for st in sorted_subtasks:
        if st['category'] != last_category:
            last_category = st['category']
            ws2.append([f'▎ {st["category"]}'] + [''] * (len(sub_cols) - 1))
            row_idx = ws2.max_row
            ws2.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=len(sub_cols))
            cell = ws2.cell(row=row_idx, column=1)
            cell.font = Font(name='Microsoft YaHei', bold=True, size=10, color='FFFFFF')
            cell.fill = sub_header_fill
            cell.border = thin_border
            for c in range(2, len(sub_cols) + 1):
                ws2.cell(row=row_idx, column=c).fill = sub_header_fill
                ws2.cell(row=row_idx, column=c).border = thin_border
        
        ed = st['end_delay']
        if ed is not None and st['pct'] < 1.0:
            if ed >= TIME_WARN_RED:
                t_level = '红色'
            elif ed >= TIME_WARN_YELLOW:
                t_level = '黄色'
            else:
                t_level = '绿色'
        elif st['status'] == '完成' or st['pct'] >= 1.0:
            t_level = '绿色'
        else:
            t_level = '-'
        
        start_plan_str = st['plan_start'].strftime('%m-%d') if st['plan_start'] else '-'
        end_plan_str = st['plan_end'].strftime('%m-%d') if st['plan_end'] else '-'
        actual_start_str = st['actual_start'].strftime('%m-%d') if st['actual_start'] else '-'
        actual_end_str = st['actual_end'].strftime('%m-%d') if st['actual_end'] else '-'
        sd_str = f'{st["start_delay"]:+d}' if st['start_delay'] is not None else '-'
        ed_str = f'{ed:+d}' if ed is not None else '-'
        pct_str = f'{st["pct"]*100:.0f}%'
        
        row_data = [
            st['category'], st['block'], st['name'], st['unit'],
            st['total_qty'] if st['total_qty'] else '-',
            start_plan_str, end_plan_str, actual_start_str, actual_end_str,
            st['status'], sd_str, ed_str, pct_str, t_level,
        ]
        ws2.append(row_data)
        row_idx = ws2.max_row
        
        if t_level == '红色':
            fill, font = red_fill, red_font
        elif t_level == '黄色':
            fill, font = yellow_fill, yellow_font
        elif t_level == '绿色':
            fill, font = green_fill, green_font
        else:
            fill, font = PatternFill(), white_font
        
        for col in range(1, len(row_data) + 1):
            cell = ws2.cell(row=row_idx, column=col)
            cell.fill = fill
            cell.border = thin_border
            cell.alignment = center_align
            if col in (3, 4, 10):
                cell.font = white_font if fill == PatternFill() else data_font
            elif col in (11, 12, 14):
                cell.font = font if fill != PatternFill() else white_font
            else:
                cell.font = white_font if fill == PatternFill() else num_font
            cell.alignment = center_align if col not in (1, 3) else wrap_align
    
    sub_col_widths = [18, 6, 18, 6, 8, 10, 10, 10, 10, 10, 12, 12, 8, 10]
    for i, w in enumerate(sub_col_widths, 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    
    # Sheet3: 统计摘要
    ws3 = wb_out.create_sheet('统计摘要')
    ws3.sheet_properties.tabColor = '00AA88'
    
    sub_red = sum(1 for st in subtasks if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_RED and st['pct'] < 1.0)
    sub_yellow = sum(1 for st in subtasks if st['end_delay'] is not None and TIME_WARN_YELLOW <= st['end_delay'] < TIME_WARN_RED and st['pct'] < 1.0)
    sub_green = sum(1 for st in subtasks if st['end_delay'] is not None and st['end_delay'] < TIME_WARN_YELLOW)
    sub_none = len(subtasks) - sub_red - sub_yellow - sub_green
    
    summary_data = [
        ['工期预警统计摘要', '', '', ''],
        ['统计日期', today.strftime('%Y-%m-%d'), '', ''],
        ['数据来源（进度计划）', plan_filename, '', ''],
        ['数据来源（日报）', daily_filename, '', ''],
        ['', '', '', ''],
        ['一、三大主任务汇总预警（vs进度计划）', '', '', ''],
        ['预警等级', '数量', '任务', '偏差天数'],
        ['🔴 红色', summary_counts['red'], '', ''],
        ['🟡 黄色', summary_counts['yellow'], '', ''],
        ['🟢 绿色', summary_counts['green'], '', ''],
        ['合计', len(summary_results), '', ''],
    ]
    for r in summary_results:
        summary_data.append([f'  {r["level"]}', '', f'{r["daily_name"]} → {r["plan_name"]}', f'{r["dev_days"]:+d}天'])
    
    summary_data.append(['', '', '', ''])
    summary_data.append(['二、子任务时间偏差统计（日报内部比对）', '', '', ''])
    summary_data.append(['时间预警', '数量', '', ''])
    summary_data.append(['🔴 延迟≥7天', sub_red, '', ''])
    summary_data.append(['🟡 延迟3-6天', sub_yellow, '', ''])
    summary_data.append(['🟢 正常', sub_green, '', ''])
    summary_data.append(['- 已完成/未开始', sub_none, '', ''])
    summary_data.append(['合计子任务', len(subtasks), '', ''])
    
    red_subtasks = [st for st in subtasks if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_RED and st['pct'] < 1.0]
    if red_subtasks:
        summary_data.append(['', '', '', ''])
        summary_data.append(['🔴 严重滞后子任务明细', '', '', ''])
        for st in red_subtasks:
            summary_data.append(['', f'[{st["block"]}] {st["name"]}',
                                f'计划结束{st["plan_end"].strftime("%m-%d") if st["plan_end"] else "?"}',
                                f'已延迟{st["end_delay"]}天'])
    
    for row_data in summary_data:
        ws3.append(row_data)
    
    ws3.cell(row=1, column=1).font = Font(name='Microsoft YaHei', bold=True, size=14, color='FFFFFF')
    ws3.merge_cells('A1:D1')
    for row_idx in range(2, ws3.max_row + 1):
        for col in [1, 2, 3, 4]:
            cell = ws3.cell(row=row_idx, column=col)
            cell.font = data_font
            cell.border = thin_border
    
    ws3.column_dimensions['A'].width = 22
    ws3.column_dimensions['B'].width = 20
    ws3.column_dimensions['C'].width = 45
    ws3.column_dimensions['D'].width = 15
    
    wb_out.save(output_file)


def send_warning_email(summary_results, subtasks_list, today, output_file, daily_file):
    """发送预警邮件；未配置 SMTP 凭证或收件人时直接跳过"""
    if not (SMTP_USER and SMTP_PASS and EMAIL_RECIPIENTS):
        print('ℹ️ 未配置 SMTP_USER / SMTP_PASS / EMAIL_RECIPIENTS，跳过邮件发送')
        return False
    try:
        # 分类顺序按 SUBTASK_CATEGORY 的声明顺序，未归类的子任务单独收在最后
        subtask_groups = {c: [] for c in dict.fromkeys(SUBTASK_CATEGORY.values())}
        for st in subtasks_list:
            if st['category'] in subtask_groups:
                subtask_groups[st['category']].append(st)
        all_subtasks_flat = [st for items in subtask_groups.values() for st in items]
        
        html_body, subject_prefix = generate_email_html(summary_results, all_subtasks_flat, today, daily_file)
        
        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = ', '.join(EMAIL_RECIPIENTS)
        msg['Subject'] = f'{subject_prefix} {PROJECT_NAME}工期预警 - {today.strftime("%Y-%m-%d")}'
        
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))
        
        if os.path.exists(output_file):
            with open(output_file, 'rb') as f:
                attachment = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                attachment.set_payload(f.read())
                encoders.encode_base64(attachment)
                attachment.add_header(
                    'Content-Disposition',
                    f'attachment; filename="工期预警报表_{today.strftime("%Y-%m-%d")}.xlsx"'
                )
                msg.attach(attachment)
        
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
        server.quit()
        return True
        
    except Exception as e:
        print(f'邮件发送失败: {str(e)}')
        return False


def generate_email_html(summary_results, subtasks_list, today, daily_file):
    """生成邮件HTML内容"""
    red_count = sum(1 for r in summary_results if r['level'] == '红色')
    yellow_count = sum(1 for r in summary_results if r['level'] == '黄色')
    green_count = sum(1 for r in summary_results if r['level'] == '绿色')
    
    if red_count > 0:
        subject_prefix = '🔴'
    elif yellow_count > 0:
        subject_prefix = '🟡'
    else:
        subject_prefix = '🟢'
    
    html = f"""
    <html>
    <body style="font-family: 'Microsoft YaHei', Arial, sans-serif; background-color: #f5f5f5; padding: 20px;">
        <div style="max-width: 900px; margin: 0 auto; background-color: white; border-radius: 8px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
            <h2 style="color: #1a2740; margin-top: 0; border-bottom: 3px solid #1a2740; padding-bottom: 10px;">
                {subject_prefix} {PROJECT_NAME}工期预警日报 - {today.strftime('%Y年%m月%d日')}
            </h2>
            <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0;">
                <p style="margin: 5px 0; font-size: 14px; color: #666;">统计日期：{today.strftime('%Y-%m-%d')} | 日报：{os.path.basename(daily_file)}</p>
                <p style="margin: 5px 0; font-size: 14px;">
                    <b>三大任务汇总：</b>
                    <span style="color: #ff4060;">🔴 {red_count}</span> |
                    <span style="color: #ffaa00;">🟡 {yellow_count}</span> |
                    <span style="color: #00ff88;">🟢 {green_count}</span>
                </p>
            </div>
    """
    
    html += """
            <h3 style="color: #1a2740;">一、三大主任务汇总比对（vs进度计划）</h3>
            <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px;">
                <thead>
                    <tr style="background-color: #1a2740; color: white;">
                        <th style="padding: 10px; text-align: left;">任务</th>
                        <th style="padding: 10px; text-align: center;">计划%</th>
                        <th style="padding: 10px; text-align: center;">实际%</th>
                        <th style="padding: 10px; text-align: center;">偏差%</th>
                        <th style="padding: 10px; text-align: center;">偏差天数</th>
                        <th style="padding: 10px; text-align: center;">等级</th>
                    </tr>
                </thead>
                <tbody>
    """
    for r in summary_results:
        level_map = {'红色': ('#3d1520', '#ff4060'), '黄色': ('#3d3015', '#ffaa00'), '绿色': ('#153d20', '#00ff88')}
        bg, tc = level_map.get(r['level'], ('#153d20', '#00ff88'))
        html += f"""
                    <tr style="background-color: {bg};">
                        <td style="padding: 8px; color: white;">{r['daily_name']}</td>
                        <td style="padding: 8px; text-align: center; color: white;">{r['plan_pct']*100:.0f}%</td>
                        <td style="padding: 8px; text-align: center; color: white;">{r['actual_pct']*100:.0f}%</td>
                        <td style="padding: 8px; text-align: center; color: {tc}; font-weight: bold;">{r['deviation']*100:+.1f}%</td>
                        <td style="padding: 8px; text-align: center; color: {tc}; font-weight: bold;">{r['dev_days']:+d}天</td>
                        <td style="padding: 8px; text-align: center; color: {tc}; font-weight: bold;">{r['level']}</td>
                    </tr>
        """
    html += """
                </tbody>
            </table>
    """
    
    delayed_subs = [st for st in subtasks_list
                    if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_YELLOW and st['pct'] < 1.0]
    if delayed_subs:
        html += """
            <h3 style="color: #ff4060; margin-top: 30px;">二、子任务时间偏差预警（红色≥7天 / 黄色≥3天）</h3>
            <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 12px;">
                <thead>
                    <tr style="background-color: #1a2740; color: white;">
                        <th style="padding: 8px; text-align: left;">区段</th>
                        <th style="padding: 8px; text-align: left;">子任务</th>
                        <th style="padding: 8px; text-align: center;">计划结束</th>
                        <th style="padding: 8px; text-align: center;">实际/今天</th>
                        <th style="padding: 8px; text-align: center;">延迟天数</th>
                        <th style="padding: 8px; text-align: center;">完成%</th>
                        <th style="padding: 8px; text-align: center;">状态</th>
                    </tr>
                </thead>
                <tbody>
        """
        for st in sorted(delayed_subs, key=lambda x: -(x['end_delay'] or 0)):
            level = '红色' if st['end_delay'] >= TIME_WARN_RED else '黄色'
            bg, tc = ('#3d1520', '#ff4060') if level == '红色' else ('#3d3015', '#ffaa00')
            plan_end_str = st['plan_end'].strftime('%m-%d') if st['plan_end'] else '-'
            html += f"""
                    <tr style="background-color: {bg};">
                        <td style="padding: 8px; color: white;">{st['block']}</td>
                        <td style="padding: 8px; color: white;">{st['name']}</td>
                        <td style="padding: 8px; text-align: center; color: white;">{plan_end_str}</td>
                        <td style="padding: 8px; text-align: center; color: white;">{today.strftime('%m-%d')}</td>
                        <td style="padding: 8px; text-align: center; color: {tc}; font-weight: bold;">+{st['end_delay']}天</td>
                        <td style="padding: 8px; text-align: center; color: white;">{st['pct']*100:.0f}%</td>
                        <td style="padding: 8px; text-align: center; color: {tc}; font-weight: bold;">{level}</td>
                    </tr>
            """
        html += """
                </tbody>
            </table>
        """
    
    html += f"""
            <div style="margin-top: 30px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; font-size: 12px; color: #999;">
                <p style="margin: 5px 0;">此邮件由工期预警管控系统自动生成</p>
                <p style="margin: 5px 0;">数据来源：{os.path.basename(daily_file)}</p>
                <p style="margin: 5px 0;">汇总比对 vs 进度计划 | 子任务时间偏差 vs 日报计划日期</p>
                <p style="margin: 5px 0;">详细报表（含全部子任务）请查看附件Excel</p>
            </div>
        </div>
    </body>
    </html>
    """
    return html, subject_prefix


# ============================================================
# 独立运行入口（路径全部从环境变量读取，仓库内不含真实路径）
# ============================================================
if __name__ == '__main__':
    PLAN_FILE = os.environ.get('PLAN_FILE', '')
    DAILY_DIR = os.environ.get('DAILY_DIR', os.path.dirname(os.path.abspath(__file__)))
    OUTPUT_DIR = os.environ.get('OUTPUT_DIR', os.path.dirname(os.path.abspath(__file__)))

    if not PLAN_FILE or not os.path.exists(PLAN_FILE):
        print('❌ 请先通过环境变量 PLAN_FILE 指定总进度计划 xlsx 的路径')
        raise SystemExit(1)

    DAILY_FILE = find_latest_daily_file(DAILY_DIR)
    if DAILY_FILE:
        print(f'📄 自动找到最新日报: {os.path.basename(DAILY_FILE)}')
        result = run_comparison(PLAN_FILE, DAILY_FILE, OUTPUT_DIR, send_email=True)
        
        if result['success']:
            print(f'\n✅ 比对完成')
            print(f'   输出文件: {result["output_file"]}')
            print(f'   邮件发送: {"✅ 成功" if result["email_sent"] else "❌ 失败"}')
            print(f'   汇总任务: {len(result["summary"])} 个 | 🔴{result["stats"]["summary_counts"]["red"]} 🟡{result["stats"]["summary_counts"]["yellow"]} 🟢{result["stats"]["summary_counts"]["green"]}')
            print(f'   子任务: {len(result["subtasks"])} 个 | 🔴{result["stats"]["subtask_counts"]["red"]} 🟡{result["stats"]["subtask_counts"]["yellow"]} 🟢{result["stats"]["subtask_counts"]["green"]}')
        else:
            print(f'\n❌ 比对失败: {result["error"]}')
    else:
        print(f'❌ 在 {DAILY_DIR} 中未找到日报文件')
