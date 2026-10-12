# -*- coding: utf-8 -*-
"""
工期预警管控系统（命令行版）
- 主任务（汇总层）：日报「进行台账」1.1节「总体」区块的全部行（动态，不限数量）
- 子任务层：日报 1.2~1.6 分类的全部内容
- Sheet2 / 邮件第二板块：1.2~1.6 全部子任务明细
- 可选大模型分析：配置 LLM_API_KEY 后自动生成问题诊断与纠偏建议；未配置则回退规则分析
输入：进度计划Excel + 日报Excel
输出：工期预警报表Excel + 邮件预警（需配置 SMTP 环境变量，否则只生成报表）
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, date, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import glob
import re
import json
import time


def _load_dotenv():
    """从脚本所在目录（或向上最多 3 层）查找 .env 并载入 os.environ。
    无第三方依赖；已存在的系统环境变量优先，不被 .env 覆盖。"""
    here = os.path.dirname(os.path.abspath(__file__))
    for i in range(4):
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
# 基础目录：默认取脚本所在目录，可用环境变量 BASE_DIR 覆盖
BASE_DIR = os.environ.get('BASE_DIR') or os.path.dirname(os.path.abspath(__file__))
PLAN_FILE = os.environ.get('PLAN_FILE') or os.path.join(BASE_DIR, 'master_plan', '总进度计划.xlsx')
KEYNODE_FILE = os.environ.get('KEYNODE_FILE') or os.path.join(BASE_DIR, 'key_nodes', '关键节点汇总.xlsx')
# 日报存放目录；若子目录不存在则回退到基础目录
_DAILY_SUBDIR = os.path.join(BASE_DIR, 'daily_reports')
DAILY_DIR = _DAILY_SUBDIR if os.path.isdir(_DAILY_SUBDIR) else BASE_DIR
# 报表输出目录；若子目录不存在则回退到基础目录
_OUTPUT_SUBDIR = os.path.join(BASE_DIR, 'warning_output')
OUTPUT_DIR = _OUTPUT_SUBDIR if os.path.isdir(_OUTPUT_SUBDIR) else BASE_DIR

# 项目名（仅用于邮件标题与提示文本，可用环境变量 PROJECT_NAME 覆盖）
PROJECT_NAME = os.environ.get('PROJECT_NAME', '示例项目')
# 日报文件名需包含的关键词（用于自动定位当日日报，逗号分隔，任一命中即可）
DAILY_FILE_KEYWORDS = [k.strip() for k in os.environ.get('DAILY_FILE_KEYWORDS', '工程日报').split(',') if k.strip()]

TODAY = datetime.now().date()

# 日报工作表名（新版），找不到时回退第一个工作表
DAILY_SHEET_NAME = os.environ.get('DAILY_SHEET_NAME', '进行台账')

# 总进度计划工作表名候选，命中第一个即用；都不命中则取第一个工作表
PLAN_SHEET_CANDIDATES = [s for s in os.environ.get('PLAN_SHEET_NAMES', '').split(',') if s.strip()]

# 主任务匹配：日报总体行名称 → 进度计划任务名（仅别名精确匹配）
# 不在别名表 = 找不到对应计划任务，一律回退用日报自身计划日期计算
# 注：已取消名称模糊匹配（会误命中，如「清槽、基础施工」错配「基础施工」）
# 下面为示例，请按自己项目的日报行名与计划任务名对应关系修改
SUMMARY_NAME_ALIAS = {
    # '日报行名': '进度计划任务名',
}

# 预警阈值
GREEN_THRESHOLD = -0.05
YELLOW_THRESHOLD = -0.10

# 子任务时间偏差阈值（天）
TIME_WARN_RED = 7    # 延迟 >= 7天 红色
TIME_WARN_YELLOW = 3 # 延迟 >= 3天 黄色

# ------------------------------------------------------------
# 队伍分区对比配置
# 假设 ABCD 四个区由两家施工队分段施工：A/D 区 = 施工队A，B/C 区 = 施工队B
# 日报子任务名形如「A1（774㎡）」，据此提取分区与面积；
# 不带括号面积的临建行（如「C7/C8生活区」）不会被误判
# 不需要分区对比时，把 ZONE_TEAM 置空即可（{}）
# ------------------------------------------------------------
ZONE_TEAM = {
    'A': '施工队A', 'D': '施工队A',
    'B': '施工队B', 'C': '施工队B',
}
ZONE_TEAM_COLORS = {'施工队A': '#2f6fed', '施工队B': '#e8833a'}   # 各队主题色（邮件表格用）

# ------------------------------------------------------------
# 邮件配置（全部通过环境变量提供，仓库内不含任何凭证）
#   SMTP_SERVER / SMTP_PORT / SMTP_USER / SMTP_PASS / EMAIL_RECIPIENTS
# SMTP_PASS 为邮箱「授权码」而非登录密码（163/QQ/126 等均需单独申请）
# 未配置 SMTP_USER 或 SMTP_PASS 时自动跳过发邮件，只生成报表
# ------------------------------------------------------------
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.163.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '465'))
SMTP_USER = os.environ.get('SMTP_USER', '')
SMTP_PASS = os.environ.get('SMTP_PASS', '')
EMAIL_RECIPIENTS = [x.strip() for x in os.environ.get('EMAIL_RECIPIENTS', '').split(',') if x.strip()]
# 调试时可设环境变量 SEND_EMAIL=0 只生成报表不发邮件
SEND_EMAIL = os.environ.get('SEND_EMAIL', '1').strip().lower() not in ('0', 'false', 'no')
if SEND_EMAIL and not (SMTP_USER and SMTP_PASS and EMAIL_RECIPIENTS):
    print('ℹ️ 未配置 SMTP_USER / SMTP_PASS / EMAIL_RECIPIENTS，本次只生成报表，不发邮件')
    SEND_EMAIL = False

# ------------------------------------------------------------
# 大模型API配置（OpenAI 兼容格式：DeepSeek / 通义千问 / 智谱 / Kimi / 内部网关均可）
# 留空 LLM_API_KEY 即自动关闭大模型，回退到原规则分析，不影响发邮件
# 密钥请通过环境变量提供，不要写进代码：
#   set LLM_API_URL=...  set LLM_API_KEY=...  set LLM_MODEL=...
# ------------------------------------------------------------
LLM_API_URL = os.environ.get('LLM_API_URL', 'https://api.deepseek.com/v1/chat/completions')
LLM_API_KEY = os.environ.get('LLM_API_KEY', '').strip()
# 模型选择说明：
#   非推理模型—— 不产生思考token，返回快且稳定（推荐）
#   推理模型 —— 思考token不可控，容易吃掉 max_tokens 全部额度导致正文为空，不建议使用
LLM_MODEL = os.environ.get('LLM_MODEL', 'deepseek-chat')
LLM_TIMEOUT = 180         # 单次请求超时（秒）
LLM_MAX_TOKENS = 6000     # 回复长度上限（非推理模型正文约需2300-3000 token，6000有充足余量）
LLM_TEMPERATURE = 0.3     # 低温度，减少发挥、贴合数据
LLM_ENABLED = bool(LLM_API_KEY)


def _parse_daily_date(fname):
    """从日报文件名提取日期，支持 2026.08.09 / 2026-08-09 / 2026年08月09日"""
    m = re.search(r'(\d{4})[.\-年](\d{1,2})[.\-月](\d{1,2})', fname)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def find_daily_file_by_date(directory, target):
    """在 directory 中查找文件名日期等于 target 的日报；未命中则退回最新。
    文件名需包含 DAILY_FILE_KEYWORDS 里的全部关键词（默认只要求「工程日报」）。"""
    pattern = os.path.join(directory, '*.xlsx')
    files = [f for f in glob.glob(pattern)
             if not os.path.basename(f).startswith('~')
             and all(kw in os.path.basename(f) for kw in DAILY_FILE_KEYWORDS)]
    if not files:
        return None
    for f in files:
        if _parse_daily_date(os.path.basename(f)) == target:
            return f
    # 未精确命中目标日期：退回修改时间最新的一份，并提示
    files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    print(f'⚠️ 未找到日期为 {target} 的日报，回退到最新: {os.path.basename(files[0])}')
    return files[0]


DAILY_FILE = find_daily_file_by_date(DAILY_DIR, TODAY - timedelta(days=1))
if DAILY_FILE:
    print(f'📄 自动找到前一天日报: {os.path.basename(DAILY_FILE)}')
else:
    print(f'❌ 在 {DAILY_DIR} 中未找到日报文件')
    DAILY_FILE = ''

OUTPUT_FILE = os.path.join(OUTPUT_DIR, f'工期预警报表_{TODAY.strftime("%Y-%m-%d")}.xlsx')


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
        # 兼容 2026.7.6 这种单数字点分日期
        m = re.match(r'^(\d{4})\.(\d{1,2})\.(\d{1,2})$', val.strip())
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
    return None


def parse_pct(val):
    """解析完成比例；无法解析时返回 None（区别于 0）"""
    if val is None or val == '\\' or val == '':
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        if '完成' in val:
            return 1.0
        elif '%' in val:
            try:
                return float(val.replace('%', '')) / 100
            except:
                return None
    return None


def warning_level(deviation):
    if deviation < YELLOW_THRESHOLD:
        return '红色'
    elif deviation < GREEN_THRESHOLD:
        return '黄色'
    else:
        return '绿色'


def time_warning_level(delay_days):
    """根据时间延迟天数返回预警等级"""
    if delay_days >= TIME_WARN_RED:
        return '红色'
    elif delay_days >= TIME_WARN_YELLOW:
        return '黄色'
    else:
        return '绿色'


def _to_int(v):
    try:
        return int(float(str(v).strip()))
    except:
        return 0


def _s(v):
    """安全转字符串"""
    return str(v).strip() if v is not None else ''


# 列索引（新版日报布局：D=区段/总体, E=工作项, F=总量, G=单位,
#   H=计划开始, I=计划完成, J=实际开始, K=实际完成, L=状态,
#   M=今日完成量, N=累计完成量, O=累计完成比例, P=备注）
COL_SEQ = 1       # B列 序号
COL_CATE = 2      # C列 分类
COL_BLOCK = 3     # D列
COL_NAME = 4      # E列
COL_TOTAL = 5     # F列
COL_UNIT = 6      # G列
COL_PSTART = 7    # H列
COL_PEND = 8      # I列
COL_ASTART = 9    # J列
COL_AEND = 10     # K列
COL_STATUS = 11   # L列
COL_PCT = 14      # O列
COL_NOTE = 15     # P列


def parse_daily(ws):
    """解析日报：主任务（1.1节「总体」区块全部行）+ 子任务（1.2~1.6全部行）"""
    # 找到表头行（B列='序号'）
    header_row = None
    for r in range(1, min(ws.max_row, 30) + 1):
        if _s(ws.cell(row=r, column=COL_SEQ + 1).value) == '序号':
            header_row = r
            break
    if header_row is None:
        header_row = 6

    # 扫描各 1.x 分节的起始行
    section_starts = []  # [(row_num, section_no, category_name)]
    for r in range(header_row + 1, ws.max_row + 1):
        b = _s(ws.cell(row=r, column=COL_SEQ + 1).value)
        m = re.match(r'^1\.(\d+)$', b)
        if m:
            cate = _s(ws.cell(row=r, column=COL_CATE + 1).value)
            section_starts.append((r, b, cate))
        # 到达「二、人材机」即结束
        if b.startswith('二') or '人材机' in b:
            break

    summary_rows = []   # 主任务（总体区块）
    subtasks = []       # 子任务（1.2~1.6）

    def read_row(r):
        row = [ws.cell(row=r, column=c + 1).value for c in range(16)]
        return row

    for idx, (start_r, sec_no, cate) in enumerate(section_starts):
        end_r = section_starts[idx + 1][0] - 1 if idx + 1 < len(section_starts) else ws.max_row
        # 分节结束也受「二、人材机」限制
        for r in range(start_r, end_r + 1):
            b = _s(ws.cell(row=r, column=COL_SEQ + 1).value)
            if b.startswith('二') or '人材机' in b:
                end_r = r - 1
                break

        if sec_no == '1.1':
            # 主任务：「总体」区块（B列合并单元格标记，向下延续到1.1节末）
            zt_start = None
            for r in range(start_r, end_r + 1):
                if '总体' in _s(ws.cell(row=r, column=COL_SEQ + 1).value):
                    zt_start = r
                    break
            if zt_start is None:
                continue
            for r in range(zt_start, end_r + 1):
                # 遇到下一个非空B列标记（新的合并区块）则总体区块结束
                b_val = _s(ws.cell(row=r, column=COL_SEQ + 1).value)
                if r > zt_start and b_val and '总体' not in b_val:
                    break
                row = read_row(r)
                name = _s(row[COL_NAME])
                if not name:
                    continue
                summary_rows.append({
                    'row': r,
                    'name': name,
                    'total_qty': row[COL_TOTAL],
                    'unit': _s(row[COL_UNIT]),
                    'plan_start': parse_date(row[COL_PSTART]),
                    'plan_end': parse_date(row[COL_PEND]),
                    'actual_start': parse_date(row[COL_ASTART]),
                    'actual_end': parse_date(row[COL_AEND]),
                    'status': _s(row[COL_STATUS]),
                    'pct': parse_pct(row[COL_PCT]),
                    'note': _s(row[COL_NOTE]),
                })
        else:
            # 子任务：1.2~1.6 的全部数据行
            for r in range(start_r, end_r + 1):
                row = read_row(r)
                name = _s(row[COL_NAME])
                if not name:
                    continue
                status = _s(row[COL_STATUS])
                pct = parse_pct(row[COL_PCT])
                plan_start = parse_date(row[COL_PSTART])
                plan_end = parse_date(row[COL_PEND])
                actual_start = parse_date(row[COL_ASTART])
                actual_end = parse_date(row[COL_AEND])
                # 异常日期保护（日报偶有笔误，如与今天相差超过1年的日期置空）
                def _guard(d):
                    if d and abs((d - TODAY).days) > 365:
                        return None
                    return d
                plan_start = _guard(plan_start)
                plan_end = _guard(plan_end)
                actual_start = _guard(actual_start)
                actual_end = _guard(actual_end)

                # 计算时间偏差
                start_delay = None
                end_delay = None
                if plan_start and actual_start:
                    start_delay = (actual_start - plan_start).days
                if plan_end:
                    if (status == '完成' or (pct is not None and pct >= 1.0)) and actual_end:
                        end_delay = (actual_end - plan_end).days
                    else:
                        end_delay = (TODAY - plan_end).days

                subtasks.append({
                    'section': sec_no,
                    'category': cate if cate else sec_no,
                    'block': _s(row[COL_BLOCK]),
                    'row': r,
                    'name': name,
                    'unit': _s(row[COL_UNIT]),
                    'total_qty': row[COL_TOTAL],
                    'plan_start': plan_start,
                    'plan_end': plan_end,
                    'actual_start': actual_start,
                    'actual_end': actual_end,
                    'status': status,
                    'pct': pct if pct is not None else 0.0,
                    'start_delay': start_delay,
                    'end_delay': end_delay,
                    'note': _s(row[COL_NOTE]),
                })

    return summary_rows, subtasks


def match_plan_task(daily_name, plan_list):
    """日报总体行名称 → 进度计划任务。仅按 SUMMARY_NAME_ALIAS 别名精确匹配；
    未命中（不在别名表）即视为「找不到」，返回 None，由调用方回退用日报自身数据。
    不做名称模糊匹配——模糊匹配会误命中（如「清槽、基础施工」错配「基础施工」）。"""
    alias = SUMMARY_NAME_ALIAS.get(daily_name)
    if alias:
        for p in plan_list:
            if p['name'] == alias:
                return p
    return None


def load_plan(filepath):
    """自适应解析总进度计划，返回任务列表（list，保留重名任务）。
    - 新版布局（表头在前 8 行内，含「任务名称」+「开始/完成时间」列）：
      标识号/任务名称/工期/开始时间/完成时间，层级由「任务名称」单元格缩进推断
      （indent0→L1, 1→L2 ...）。
    - 旧版布局回退：name=row[4]、level=row[1]、duration=row[5]、start=row[6]、end=row[7]。
    工作表优先取 PLAN_SHEET_CANDIDATES 中命中的第一个，否则取第一个工作表。
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = None
    for cand in PLAN_SHEET_CANDIDATES:
        if cand in wb.sheetnames:
            ws = wb[cand]
            break
    if ws is None:
        ws = wb.worksheets[0]

    # 定位表头行 + 列名映射
    header_row = None
    col = {}
    for r in range(1, 9):
        names = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v:
                names[str(v).strip()] = c - 1
        if '任务名称' in names and ('开始时间' in names or '完成时间' in names):
            header_row = r
            col = names
            break

    def _dur(dur_val, start, end):
        if isinstance(dur_val, str):
            d = int(re.sub(r'[^\d]', '', dur_val) or '0')
        elif isinstance(dur_val, (int, float)):
            d = int(dur_val)
        else:
            d = (end - start).days
        return max(d, (end - start).days)

    plan_list = []
    if header_row:
        # 新版布局
        c_name = col['任务名称']
        c_dur = col.get('工期')
        c_start = col.get('开始时间')
        c_end = col.get('完成时间')
        for r in range(header_row + 1, ws.max_row + 1):
            cell_name = ws.cell(row=r, column=c_name + 1)
            name = cell_name.value
            if not name:
                continue
            name = str(name).strip()
            start = parse_date(ws.cell(row=r, column=c_start + 1).value) if c_start is not None else None
            end = parse_date(ws.cell(row=r, column=c_end + 1).value) if c_end is not None else None
            if not (start and end):
                continue
            dur_val = ws.cell(row=r, column=c_dur + 1).value if c_dur is not None else None
            indent = cell_name.alignment.indent if cell_name.alignment else 0
            plan_list.append({
                'name': name, 'start': start, 'end': end,
                'duration': _dur(dur_val, start, end),
                'level': f'L{int(indent) + 1}', 'type': '', 'row': r,
            })
    else:
        # 旧版布局回退
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            name = row[4]
            if not name:
                continue
            name = str(name).strip()
            start = parse_date(row[6])
            end = parse_date(row[7])
            if start and end:
                plan_list.append({
                    'name': name, 'start': start, 'end': end,
                    'duration': _dur(row[5], start, end),
                    'level': str(row[1]) if row[1] else '', 'type': str(row[2]) if row[2] else '',
                })
    return plan_list


def parse_resources(ws_daily):
    """解析日报「二、人材机」板块。
    布局：人员(col E名称/F数量/G单位"人") | 机械(col I名称/J数量) | 材料(col M名称/N数量)
    分区：col B = 2.1/2.2... ，分区名在 col C（如 总包/地源热泵）
    """
    resources = {}
    start_row = None
    for r in range(1, ws_daily.max_row + 1):
        v = ws_daily.cell(row=r, column=2).value
        if v and '人材机' in str(v):
            start_row = r
            break
    if not start_row:
        return resources

    current = None
    for r in range(start_row + 1, ws_daily.max_row + 1):
        b_val = ws_daily.cell(row=r, column=2).value
        b_str = str(b_val).strip() if b_val is not None else ''
        if b_str.startswith('三') or '风险' in b_str:
            break
        if re.match(r'^2\.\d+$', b_str):
            sec_name = ws_daily.cell(row=r, column=3).value
            sec_name = str(sec_name).strip() if sec_name else b_str
            current = {'name': sec_name, 'labor': [], 'labor_total': 0,
                       'special': 0, 'machinery': [], 'materials': []}
            resources[sec_name] = current
            # 注意：不 continue，分区标题行本身也携带人员/机械/材料数据
        if current is None:
            continue
        # 人员：E=名称 F=数量 G=单位
        name_v = ws_daily.cell(row=r, column=5).value
        cnt_v = ws_daily.cell(row=r, column=6).value
        unit_v = ws_daily.cell(row=r, column=7).value
        if name_v and unit_v and str(unit_v).strip() == '人':
            name_s = str(name_v).strip()
            cnt = _to_int(cnt_v)
            if '特殊工种' in name_s:
                current['special'] = cnt
            elif '合计' in name_s:
                current['labor_total'] = cnt
            else:
                current['labor'].append((name_s, cnt))
        # 机械：I=名称 J=数量
        mach_name = ws_daily.cell(row=r, column=9).value
        mach_cnt = _to_int(ws_daily.cell(row=r, column=10).value)
        if mach_name and mach_cnt > 0:
            current['machinery'].append((str(mach_name).strip(), mach_cnt))
        # 材料：M=名称 N=数量
        mat_name = ws_daily.cell(row=r, column=13).value
        mat_qty = ws_daily.cell(row=r, column=14).value
        if mat_name and mat_qty:
            current['materials'].append((str(mat_name).strip(), str(mat_qty).strip()))
    return resources


def get_upcoming_starts(plan_tasks, today, days=7):
    """从总进度计划中筛选出以 today 为基准、未来 days 天内计划启动的任务"""
    horizon = today + timedelta(days=days)
    items = plan_tasks.values() if isinstance(plan_tasks, dict) else plan_tasks
    upcoming = [p for p in items if p['start'] and today <= p['start'] <= horizon]
    upcoming.sort(key=lambda x: (x['start'], x['end']))
    return upcoming


def load_key_nodes(filepath):
    """读取关键节点汇总表（序号/关键节点名称/时间/类别），返回节点列表"""
    if not filepath or not os.path.exists(filepath):
        print(f'  ⚠️ 关键节点文件不存在: {filepath}')
        return []
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.worksheets[0]
    except Exception as e:
        print(f'  ⚠️ 关键节点文件读取失败: {e}')
        return []
    nodes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[1]
        if not name:
            continue
        node_date = parse_date(row[2])
        if not node_date:
            continue
        category = str(row[3]).strip() if len(row) > 3 and row[3] else ''
        nodes.append({
            'name': str(name).strip(),
            'date': node_date,
            'category': category,
        })
    return nodes


def get_upcoming_keynodes(nodes, today, days=15):
    """筛选未来 days 天内（含今天）的关键节点；已过期的单独标记"""
    horizon = today + timedelta(days=days)
    upcoming = [n for n in nodes if today <= n['date'] <= horizon]
    upcoming.sort(key=lambda x: (x['date'], x['category'], x['name']))
    return upcoming


def resource_insights(resources, subtasks):
    """结合人材机数据与预警子任务，生成资源观察与建议"""
    observations = []
    suggestions = []
    if not resources:
        return observations, suggestions

    total_labor = sum((r['labor_total'] or sum(c for _, c in r['labor'])) for r in resources.values())
    total_special = sum(r['special'] for r in resources.values())
    total_mach = sum(sum(c for _, c in r['machinery']) for r in resources.values())
    all_mats = ['、'.join(f'{n}({q})' for n, q in r['materials']) for r in resources.values() if r['materials']]
    observations.append(f"现场共投入 {total_labor} 人（含特殊工种 {total_special} 人）、机械 {total_mach} 台"
                        + (f"；主要材料：{'；'.join(all_mats)}" if all_mats else ""))

    red_names = [st['name'] for st in subtasks
                 if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_RED and st['pct'] < 1.0]
    paused_names = [st['name'] for st in subtasks if st['status'] == '工序暂停']

    # 关键词 → 资源要素映射（用于交叉验证滞后任务的资源保障）
    keyword_resource = {
        '塔吊': ('履带吊', '塔吊人员'),
        '清槽': ('挖机', '挖机司机'),
        '钎探': ('挖机', '挖机司机'),
        '抗浮锚杆': ('锚杆机', '抗浮锚杆人员'),
        '抗拔桩': ('旋挖机', '工程桩人员'),
        '地源热泵': ('挖掘机', '水平管工'),
        '降水': ('降水设备', '降水人员'),
        '静载试验': ('挖机', '工程桩人员'),
    }
    for sec_name, res in resources.items():
        zero_labor = [n for n, c in res['labor'] if c == 0]
        if zero_labor:
            observations.append(f"⚠️ {sec_name}：{'/'.join(zero_labor)} 工种零人投入")
            suggestions.append(f"核查{sec_name}零投入工种（{'/'.join(zero_labor)}）对应的子任务是否需要补充劳动力。")
        # 滞后/暂停任务与该分区资源的交叉提示
        labor_names = ' '.join(n for n, _ in res['labor'])
        for kw in keyword_resource:
            hit = any(kw in rn for rn in red_names + paused_names)
            if hit and kw in labor_names:
                cnt = sum(c for n, c in res['labor'] if kw in n)
                if cnt <= 0:
                    suggestions.append(f"滞后任务涉及「{kw}」，但{sec_name}当前无相应人员投入，建议尽快补员。")

    if total_labor > 0 and total_special / total_labor > 0.15:
        observations.append(f"特殊工种占比 {total_special/total_labor*100:.0f}%，注意持证上岗与作业安全。")
    return observations, suggestions[:2]


def _fmt_date(d, default='-'):
    return d.strftime('%Y-%m-%d') if d else default


# ============================================================
# 施工队分区进度对比（按 ZONE_TEAM 配置划分分区与施工队）
# ============================================================
_ZONE_NAME_PAT = re.compile(r'^([ABCD])\s*(\d+)\s*[\uff08(]\s*([\d.]+)')


def _zone_median(vals):
    """中位数：施工强度受小面积短任务干扰大，中位数比均值更抗噪"""
    if not vals:
        return None
    v = sorted(vals)
    n = len(v)
    return v[n // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2.0


def collect_zone_tasks(subtasks):
    """从日报子任务中挑出 A/B/C/D 分区任务，标注施工队。

    识别依据：任务名形如「A1（774㎡）」——首字母为分区、括号内为面积。
    临建板块的「C7/C8生活区」「C3/C4/C5/C6区」等不带「数字（面积）」结构，天然不会被匹配。
    """
    out = []
    for st in subtasks:
        m = _ZONE_NAME_PAT.match(_s(st['name']).strip())
        if not m:
            continue
        z = m.group(1).upper()
        team = ZONE_TEAM.get(z)
        if not team:
            continue
        try:
            area = float(m.group(3))
        except (TypeError, ValueError):
            area = 0.0
        out.append({
            'team': team, 'zone': z, 'name': _s(st['name']), 'area': area,
            'category': st['category'], 'pct': st['pct'] or 0.0, 'status': st['status'],
            'plan_end': st['plan_end'], 'actual_start': st['actual_start'],
            'actual_end': st['actual_end'],
        })
    return out


def _zone_is_done(z):
    return z['status'] == '完成' or z['pct'] >= 1.0


def team_zone_stats(items, today):
    """单队指标。

    ⚠️ 口径说明（重要，别改回「完成数/总任务数」）：
    两队的分区排期不同（如 D6~D11 计划 11-12 月才开工），直接用「已完成/全部任务」
    会把"还没到开工时间"当成"效率低"，结论完全失真（实测 36% vs 46% → 实为 86% vs 80%）。
    因此核心指标用**到期应完成履约率** = 已完成 ÷ 计划完成日已到的任务数。
    """
    due = [z for z in items if z['plan_end'] and z['plan_end'] <= today]
    due_done = [z for z in due if _zone_is_done(z)]
    due_late = [z for z in due if not _zone_is_done(z)]

    # 状态三分类必须互斥（曾把"未到开工期"与前两项重复计数，导致相加超过任务总数）
    done_items = [z for z in items if _zone_is_done(z)]
    doing = [z for z in items if not _zone_is_done(z) and z['pct'] > 0]
    todo = [z for z in items if not _zone_is_done(z) and z['pct'] <= 0]
    future = [z for z in items if not _zone_is_done(z) and z['plan_end'] and z['plan_end'] > today]

    # 已完成任务的平均偏差（实际完成 - 计划完成，正=拖后）
    devs = [(z['actual_end'] - z['plan_end']).days for z in due_done if z['actual_end'] and z['plan_end']]
    late_days = [(today - z['plan_end']).days for z in due_late if z['plan_end']]

    speeds = []
    for z in items:
        if _zone_is_done(z) and z['actual_start'] and z['actual_end']:
            d = (z['actual_end'] - z['actual_start']).days + 1
            if d >= 1:                      # 日报偶有实际完成早于开始/工期为0的笔误，剔除
                speeds.append(z['area'] / d)

    return {
        'n': len(items),
        'zones': sorted({z['zone'] for z in items}),
        'done': len(done_items),
        'doing': len(doing),
        'todo': len(todo),
        'future': len(future),
        'due': len(due), 'due_done': len(due_done), 'due_late': len(due_late),
        'fulfill': (len(due_done) / len(due)) if due else None,
        'avg_dev': (sum(devs) / len(devs)) if devs else None,
        'worst_late': max(late_days) if late_days else None,
        'speed_med': _zone_median(speeds),
        'speed_mean': (sum(speeds) / len(speeds)) if speeds else None,
        'late_list': sorted(
            [{'name': z['name'], 'zone': z['zone'], 'category': z['category'],
              'plan_end': z['plan_end'], 'pct': z['pct'],
              'over': (today - z['plan_end']).days if z['plan_end'] else 0}
             for z in due_late],
            key=lambda x: -x['over']),
    }


def _zone_team_order(teams):
    """按 ZONE_TEAM 配置里的声明顺序排，让报表列序稳定"""
    cfg_order = list(dict.fromkeys(ZONE_TEAM.values()))
    return sorted(teams, key=lambda t: (cfg_order.index(t) if t in cfg_order else 99, t))


def compare_zone_teams(subtasks, today):
    """按 ZONE_TEAM 分组统计，返回 {队名: stats}（无分区任务时返回 {}）"""
    zt = collect_zone_tasks(subtasks)
    if not zt:
        return {}
    teams = _zone_team_order({z['team'] for z in zt})
    return {t: team_zone_stats([z for z in zt if z['team'] == t], today) for t in teams}


def zone_team_by_category(subtasks, today):
    """分工序的到期履约率对比，返回 (rows, teams)；无分区任务时返回 ([], [])"""
    zt = collect_zone_tasks(subtasks)
    if not zt:
        return [], []
    teams = _zone_team_order({z['team'] for z in zt})
    cats = sorted({z['category'] for z in zt})
    rows = []
    for c in cats:
        row = {'category': c}
        for t in teams:
            row[t] = team_zone_stats([z for z in zt if z['team'] == t and z['category'] == c], today)
        rows.append(row)
    return rows, teams


def zone_team_insights(team_stats, today):
    """生成 2-3 条精简结论（不啰嗦，只讲效率差异与最该关注的事）"""
    notes = []
    if len(team_stats) < 2:
        return notes
    items = list(team_stats.items())[:2]
    (t1, s1), (t2, s2) = items[0], items[1]

    # 履约率对比（核心）
    if s1['fulfill'] is not None and s2['fulfill'] is not None:
        if abs(s1['fulfill'] - s2['fulfill']) < 0.05:
            notes.append(f"{t1}（{'/'.join(s1['zones'])}区）与{t2}（{'/'.join(s2['zones'])}区）"
                         f"到期任务履约率基本持平（{s1['fulfill']*100:.0f}% vs {s2['fulfill']*100:.0f}%），两队推进节奏相当。")
        else:
            if s1['fulfill'] >= s2['fulfill']:
                ht, hs, lt, ls = t1, s1, t2, s2
            else:
                ht, hs, lt, ls = t2, s2, t1, s1
            notes.append(f"到期任务履约率：{ht} {hs['fulfill']*100:.0f}%"
                         f"（{hs['due_done']}/{hs['due']}）高于 {lt} {ls['fulfill']*100:.0f}%"
                         f"（{ls['due_done']}/{ls['due']}），{lt}当期交付稍弱。")

    # 施工强度对比
    if s1['speed_med'] and s2['speed_med']:
        if s1['speed_med'] >= s2['speed_med']:
            ft, fs, st_, ss = t1, s1, t2, s2
        else:
            ft, fs, st_, ss = t2, s2, t1, s1
        notes.append(f"单位面积施工强度（中位数）：{ft} {fs['speed_med']:.0f}㎡/天 vs "
                     f"{st_} {ss['speed_med']:.0f}㎡/天，{ft}单块推进略快。")

    # 逾期风险
    lates = [(t, s) for t, s in team_stats.items() if s['late_list']]
    if lates:
        parts = []
        for t, s in lates:
            top = s['late_list'][0]
            parts.append(f"{t}{top['name'].split('（')[0]}逾期{top['over']}天（{top['pct']*100:.0f}%）")
        notes.append("⚠️ 到期未完成：" + "；".join(parts) + "，需优先督办。")
    return notes[:3]


def build_llm_context(summary_results, subtasks, resources, upcoming, upcoming_keynodes, today):
    """把当天所有结构化数据压成一段给大模型看的事实文本（只给数据，不夹带结论）"""
    L = []
    L.append(f"【统计日期】{today.strftime('%Y-%m-%d')}（{'星期' + '一二三四五六日'[today.weekday()]}）")
    L.append(f"【项目】{PROJECT_NAME}\n")

    L.append("【一、主任务（总体行）vs 总进度计划】")
    L.append("指标说明：计划应完成%=按工期线性推算（已过天数/总工期，未考虑工程量分布，仅供量级参考）；"
             "实际完成%=日报填报的实物量完成比例；偏差=实际-计划（负数=滞后）；"
             "偏差天数=偏差×总工期取整；「对应计划任务」为“日报计划日期”表示总进度计划中无对应条目、用日报自填的计划日期作基准。")
    if summary_results:
        for r in summary_results:
            if r['deviation'] is None:
                L.append(f"- {r['daily_name']}：无量化数据（等级={r['level']}，状态={r.get('status') or '-'}）")
            else:
                L.append(f"- {r['daily_name']}（对应计划任务：{r['plan_name']}）：计划{_fmt_date(r['plan_start'])}~{_fmt_date(r['plan_end'])}，"
                         f"总工期{r['total_days']}天，已过{r['elapsed_days']}天，计划应完成{r['plan_pct']*100:.0f}%，"
                         f"实际完成{r['actual_pct']*100:.0f}%，偏差{r['deviation']*100:+.1f}%（约{r['dev_days']:+d}天），等级={r['level']}")
    else:
        L.append("（无）")

    done = [st for st in subtasks if st['pct'] >= 1.0]
    paused = [st for st in subtasks if st['status'] == '工序暂停' and st['pct'] < 1.0]
    not_started = [st for st in subtasks if st['pct'] == 0 and st['status'] != '工序暂停'
                   and st['status'] in (None, '', '未开始')]
    delayed = [st for st in subtasks if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_YELLOW and st['pct'] < 1.0]
    in_progress = len(subtasks) - len(done) - len(paused) - len(not_started)
    L.append(f"\n【二、子任务概况】共{len(subtasks)}项：已完成{len(done)}，进行中{max(in_progress, 0)}，"
             f"未开始{len(not_started)}，工序暂停{len(paused)}；其中黄色及以上延迟{len(delayed)}项")
    if delayed:
        L.append("延迟子任务清单（结束偏差 = 实际结束日与计划结束日之差；未完成时以当天计，正数=已超期天数）：")
        for st in sorted(delayed, key=lambda x: -x['end_delay']):
            sd = f"{st['start_delay']:+d}天" if st['start_delay'] is not None else '未开始无偏差'
            _u = str(st.get('unit') or '').strip().strip('\\')
            _q = str(st.get('total_qty') or '').strip().strip('\\')
            qty_s = f"{_q}{_u}" if _q and _q.lower() not in ('none', '0') else '未填工程量'
            L.append(f"- [{st['section']} {st['block'] or '总体'}] {st['name']}：状态={st['status'] or '-'}，"
                     f"完成{st['pct']*100:.0f}%，计划{_fmt_date(st['plan_start'])}~{_fmt_date(st['plan_end'])}，"
                     f"开始偏差{sd}，结束偏差{st['end_delay']:+d}天，工程量={qty_s}")
    else:
        L.append("延迟子任务：（无黄色及以上）")

    if resources:
        L.append("\n【三、人材机投入（日报当日）】")
        for sec, res in resources.items():
            lab = '、'.join(f'{n}{c}人' for n, c in res['labor']) or '未填报'
            lt = res['labor_total'] or sum(c for _, c in res['labor'])
            mac = '、'.join(f'{n}{c}台' for n, c in res['machinery']) or '无'
            mat = '、'.join(f'{n}{q}' for n, q in res['materials']) or '无'
            L.append(f"- {sec}：人员合计{lt}人（特殊工种{res['special']}人）〔{lab}〕；机械〔{mac}〕；材料〔{mat}〕")

    if upcoming:
        L.append(f"\n【四、未来一周计划启动任务（共{len(upcoming)}项）】")
        for p in upcoming[:25]:
            L.append(f"- {p['name']}：{_fmt_date(p['start'])}~{_fmt_date(p['end'])}（工期{p['duration']}天，距今{(p['start']-today).days}天）")

    if upcoming_keynodes:
        L.append(f"\n【五、未来半个月关键节点（共{len(upcoming_keynodes)}个）】")
        for n in upcoming_keynodes[:20]:
            L.append(f"- {n['date'].strftime('%Y-%m-%d')}（{(n['date']-today).days}天后）[{n['category'] or '未分类'}] {n['name']}")

    # 施工队分区对比（只给数据，不给结论）
    try:
        team_stats = compare_zone_teams(subtasks, today)
    except Exception:
        team_stats = {}
    if team_stats:
        L.append("\n【六、施工队分区进度对比】")
        _zone_desc = '；'.join(
            f"{'/'.join(z for z, tt in ZONE_TEAM.items() if tt == t)}区由「{t}」施工"
            for t in dict.fromkeys(ZONE_TEAM.values()))
        L.append(f"口径说明：{_zone_desc}，各队分区排期不同（部分分区尚未到开工期）；"
                 "「到期履约率」=已完成÷计划完成日已到期的任务数，是各队可比的效率口径；"
                 "「施工强度中位数」=已完成任务的面积÷实际日历天取中位数。")
        for t, s in team_stats.items():
            fr = '%.0f%%' % (s['fulfill'] * 100) if s['fulfill'] is not None else '无到期任务'
            sp = '%.0f㎡/天' % s['speed_med'] if s['speed_med'] else '-'
            ad = '%+.1f天' % s['avg_dev'] if s['avg_dev'] is not None else '-'
            L.append(f"- {t}（{'/'.join(s['zones'])}区）：任务{s['n']}项 = 已完成{s['done']} + 在建{s['doing']} + 未开工{s['todo']}"
                     f"（其中{s['future']}项计划完成日尚未到期，属正常排期）；到期履约率{fr}（{s['due_done']}/{s['due']}）；"
                     f"平均完成偏差{ad}；施工强度中位数{sp}")
            for it in s['late_list'][:5]:
                L.append(f"    · 到期未完成：{it['name']}（{it['category']}，计划{it['plan_end']}，"
                         f"完成{it['pct']*100:.0f}%，已逾期{it['over']}天）")

    return '\n'.join(L)


def build_llm_prompt(ctx_text):
    """系统提示：把它约束成项目工期管控工程师，只准基于给定数据分析"""
    sys_prompt = (
        "你是一名大型建筑工程项目部的生产经理兼进度管控工程师，擅长施工工序逻辑、流水段划分、"
        "人材机资源配置与工期纠偏。你会收到一份项目日报的结构化数据，需要输出给项目领导看的分析意见。\n\n"
        "硬性要求：\n"
        "1. 只能基于给定数据推理，严禁编造数据中没有的任务名、日期、人数、工程量；数据缺失就直说缺失。\n"
        "2. 分析原因要落到施工逻辑上（如工序衔接、工作面移交、劳动力/机械配置不足、材料供应、"
        "天气、前置条件未完成、验收检测等待、流水段穿插冲突等），并说明判断依据。\n"
        "3. 建议必须具体可执行，写明「谁去做、做什么、什么时间节点、达到什么效果」，避免空话套话。\n"
        "4. 分清轻重缓急，先讲影响关键线路的，再讲次要的。\n"
        "5. 用词专业、简洁，符合施工企业内部生产例会汇报口吻，不要用感叹号堆砌情绪。\n"
        "6. 严格按下面要求的 JSON 格式输出，不要输出 JSON 以外的任何文字、解释或 markdown 代码块标记。"
    )
    user_prompt = f"""以下是项目当日数据：

{ctx_text}

请输出 JSON，格式严格如下（字符串数组，每项都是一句完整的话）：
{{
  "overall": "总体研判，2-4句，先给结论再说主要矛盾",
  "problems": [
    {{"issue": "问题现象（带数据）", "cause": "可能原因分析（结合施工逻辑与资源配置）", "impact": "对工期/后续工序的影响"}}
  ],
  "suggestions": [
    {{"priority": "高|中|低", "action": "具体建议措施", "owner": "建议责任方/岗位", "deadline": "建议完成时限"}}
  ],
  "resources": ["人材机资源配置方面的观察与建议，1-4条"],
  "risks": ["未来一周至半个月的风险预判（结合即将启动任务与关键节点），1-3条"]
}}

要求：problems 3-6条（按严重程度排序），suggestions 4-6条，resources 1-4条，risks 1-3条。
如果某类数据确实没有可分析的内容，对应数组给1条说明"数据不足"的条目，不要硬凑。"""
    return sys_prompt, user_prompt


def _extract_json(text):
    """从模型回复里稳健地抠出 JSON 对象"""
    if not text:
        return None
    t = text.strip()
    # 去掉 ```json ... ``` 包裹
    t = re.sub(r'^```(?:json)?\s*', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\s*```$', '', t)
    t = t.strip()
    try:
        return json.loads(t)
    except Exception:
        pass
    # 退而求其次：截取第一个 { 到最后一个 }
    s, e = t.find('{'), t.rfind('}')
    if s >= 0 and e > s:
        try:
            return json.loads(t[s:e + 1])
        except Exception:
            return None
    return None


def call_llm(sys_prompt, user_prompt, max_retries=2):
    """调用 OpenAI 兼容的大模型接口；带重试与诊断。
    任何异常或空返回都会重试，最终失败返回 None（由调用方回退规则分析）。"""
    import urllib.request
    import urllib.error
    last_reason = ''
    for attempt in range(1, max_retries + 1):
        payload = {
            'model': LLM_MODEL,
            'messages': [
                {'role': 'system', 'content': sys_prompt},
                {'role': 'user', 'content': user_prompt},
            ],
            'temperature': LLM_TEMPERATURE,
            'max_tokens': LLM_MAX_TOKENS,
            'stream': False,
        }
        req = urllib.request.Request(
            LLM_API_URL,
            data=json.dumps(payload).encode('utf-8'),
            headers={
                'Content-Type': 'application/json',
                'Authorization': f'Bearer {LLM_API_KEY}',
            },
            method='POST',
        )
        try:
            with urllib.request.urlopen(req, timeout=LLM_TIMEOUT) as resp:
                body = json.loads(resp.read().decode('utf-8'))
            ch = body.get('choices', [{}])[0]
            content = (ch.get('message', {}) or {}).get('content') or ''
            finish = ch.get('finish_reason', '')
            usage = body.get('usage', {}) or {}
            comp = usage.get('completion_tokens')
            rt = (usage.get('completion_tokens_details') or {}).get('reasoning_tokens')
            if content.strip():
                if finish == 'length':
                    print(f'  ⚠️ 大模型回复被截断（finish=length），建议调大 LLM_MAX_TOKENS')
                print(f'  ℹ️ 大模型用量: 完成{comp}token（含思考{rt}）')
                return content
            # 空内容：多半是推理模型把 token 额度耗尽，正文没写出来
            last_reason = f'返回正文为空（finish={finish}, 完成{comp}token/思考{rt}）'
            print(f'  ⚠️ 第{attempt}次调用{last_reason}')
        except urllib.error.HTTPError as e:
            detail = e.read().decode('utf-8', errors='ignore')[:200]
            last_reason = f'HTTP {e.code}: {detail}'
            # 429/5xx 可重试，4xx 鉴权类不重试
            if e.code not in (429, 500, 502, 503, 504):
                print(f'  ❌ 大模型API返回错误 {e.code}: {detail}')
                return None
            print(f'  ⚠️ 第{attempt}次调用 {last_reason}')
        except Exception as e:
            last_reason = f'{type(e).__name__}: {e}'
            print(f'  ⚠️ 第{attempt}次调用失败 {last_reason}')
        if attempt < max_retries:
            wait = 3 * attempt
            print(f'  ⏳ {wait}秒后重试...')
            time.sleep(wait)
    print(f'  ❌ 大模型多次调用失败，最后一次原因：{last_reason}')
    return None


def render_zone_team_html(team_stats, cat_rows, teams, insights, today):
    """渲染「施工队分区进度对比」模块（放在人材机资源配置分析之后）"""
    if not team_stats or len(teams) < 2:
        return ''
    esc = lambda s: (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                     if s is not None else '')
    h = ['<p style="margin: 18px 0 6px 0; font-size: 14px; color: #222;">'
         '<b>🚧 施工队分区进度对比（' + ' vs '.join(
             '%s·%s区' % (esc(t), '/'.join(team_stats[t]['zones'])) for t in teams) + '）</b></p>']

    # --- 核心指标对照表 ---
    h.append('<table style="width: 100%; border-collapse: collapse; font-size: 12.5px;">'
             '<thead><tr style="background-color: #1a2740; color: #fff;">'
             '<th style="padding: 6px 8px; text-align: left; width: 30%;">对比指标</th>')
    for t in teams:
        h.append('<th style="padding: 6px 8px; text-align: center; color: %s;">%s（%s区）</th>'
                 % ('#fff', esc(t), '/'.join(team_stats[t]['zones'])))
    h.append('</tr></thead><tbody>')

    def _pct(v):
        return '%.0f%%' % (v * 100) if v is not None else '-'

    def _num(v, suf='', sign=False):
        if v is None:
            return '-'
        return ('%+.1f' % v if sign else '%.0f' % v) + suf

    # 每行：(标签, 取值函数, 谁更优 'high'|'low'|None)
    rows_def = [
        ('到期任务履约率', lambda s: '%s（%d/%d）' % (_pct(s['fulfill']), s['due_done'], s['due']),
         lambda s: s['fulfill'], 'high'),
        ('平均完成偏差', lambda s: _num(s['avg_dev'], '天', sign=True),
         lambda s: s['avg_dev'], 'low'),
        ('施工强度中位数', lambda s: _num(s['speed_med'], '㎡/天'),
         lambda s: s['speed_med'], 'high'),
        ('到期未完成', lambda s: '%d 项' % s['due_late'],
         lambda s: s['due_late'], 'low'),
        ('已完成 / 在建 / 未开工', lambda s: '%d / %d / %d' % (s['done'], s['doing'], s['todo']),
         None, None),
        ('任务总数（含未到开工期）', lambda s: '%d 项' % s['n'], None, None),
    ]
    for label, fmt, keyf, better in rows_def:
        vals = {t: keyf(team_stats[t]) if keyf else None for t in teams}
        # 标出优势方（数值可比且不相等时）
        best = None
        if better and all(v is not None for v in vals.values()):
            if better == 'high':
                best = max(vals, key=lambda t: vals[t])
            else:
                best = min(vals, key=lambda t: vals[t])
            if len(set(vals.values())) == 1:
                best = None
        h.append('<tr><td style="padding: 6px 8px; color: #444; border-bottom: 1px solid #eee;">%s</td>' % esc(label))
        for t in teams:
            strong = (t == best)
            h.append('<td style="padding: 6px 8px; text-align: center; border-bottom: 1px solid #eee; '
                     'color: %s; %s">%s</td>'
                     % (ZONE_TEAM_COLORS.get(t, '#333') if strong else '#333',
                        'font-weight: bold;' if strong else '', esc(fmt(team_stats[t]))))
        h.append('</tr>')
    h.append('</tbody></table>')

    # --- 分工序履约率 ---
    if cat_rows:
        h.append('<table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 8px;">'
                 '<thead><tr style="background-color: #eef1f6; color: #333;">'
                 '<th style="padding: 5px 8px; text-align: left;">分工序到期履约率</th>')
        for t in teams:
            h.append('<th style="padding: 5px 8px; text-align: center;">%s</th>' % esc(t))
        h.append('</tr></thead><tbody>')
        for row in cat_rows:
            h.append('<tr><td style="padding: 5px 8px; color: #444; border-bottom: 1px solid #f0f0f0;">%s</td>'
                     % esc(row['category']))
            for t in teams:
                s = row.get(t) or {}
                if s.get('due'):
                    txt = '%d/%d（%.0f%%）' % (s['due_done'], s['due'], s['fulfill'] * 100)
                    color = '#d92d20' if s['fulfill'] < 0.7 else ('#dc8a00' if s['fulfill'] < 1.0 else '#00875a')
                else:
                    # due=0：该工序下所有任务计划完成日均未到期（可能已在建），不能写成"未开工"
                    txt = '暂未到期（%d项）' % s.get('n', 0)
                    color = '#999'
                h.append('<td style="padding: 5px 8px; text-align: center; color: %s; '
                         'border-bottom: 1px solid #f0f0f0;">%s</td>' % (color, esc(txt)))
            h.append('</tr>')
        h.append('</tbody></table>')

    # --- 逾期明细（最多列4条）---
    lates = []
    for t in teams:
        for it in team_stats[t]['late_list'][:4]:
            lates.append((t, it))
    if lates:
        lates.sort(key=lambda x: -x[1]['over'])
        h.append('<div style="margin-top: 8px; padding: 8px 10px; background-color: #fff5f6; '
                 'border-left: 3px solid #d92d20; border-radius: 3px; font-size: 12px; color: #555;">'
                 '<b>到期未完成明细：</b>')
        seg = []
        for t, it in lates[:6]:
            seg.append('%s %s（%s，计划%s，已完成%.0f%%，逾期%d天）'
                       % (esc(t), esc(it['name'].split('（')[0]), esc(it['category']),
                          it['plan_end'].strftime('%m-%d') if it['plan_end'] else '-',
                          it['pct'] * 100, it['over']))
        h.append('；'.join(seg) + '</div>')

    # --- 结论 ---
    if insights:
        h.append('<ul style="margin: 8px 0 4px 0; padding-left: 22px; font-size: 13px; color: #444; line-height: 1.75;">')
        for it in insights:
            h.append('<li style="margin: 4px 0;">%s</li>' % esc(it))
        h.append('</ul>')
    h.append('<p style="margin: 6px 0 0 0; font-size: 11px; color: #999;">'
             '口径：履约率=已完成÷计划完成日已到期的任务数（不含尚未到开工期的分区，避免排期差异干扰）；'
             '施工强度=已完成任务面积÷实际日历天，取中位数。</p>')
    return '\n'.join(h)


def render_llm_html(data, zone_html=''):
    """把大模型返回的 JSON 渲染成邮件 HTML 片段；zone_html 插在「人材机资源配置分析」之后"""
    h = []
    esc = lambda s: (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                     if s is not None else '')

    overall = esc(data.get('overall') or '').strip()
    if overall:
        h.append(f'<p style="margin: 5px 0; font-size: 14px; color: #222; line-height: 1.7;"><b>📊 总体研判：</b>{overall}</p>')

    problems = data.get('problems') or []
    if problems:
        h.append('<p style="margin: 14px 0 6px 0; font-size: 14px; color: #222;"><b>🔍 问题定位与原因分析</b></p>')
        h.append('<table style="width: 100%; border-collapse: collapse; font-size: 12.5px;">')
        h.append('<thead><tr style="background-color: #1a2740; color: #fff;">'
                 '<th style="padding: 6px 8px; text-align: left; width: 26%;">问题现象</th>'
                 '<th style="padding: 6px 8px; text-align: left; width: 42%;">可能原因</th>'
                 '<th style="padding: 6px 8px; text-align: left;">工期影响</th></tr></thead><tbody>')
        for i, p in enumerate(problems, 1):
            bg = '#fff5f6' if i <= 2 else '#ffffff'
            h.append(f'<tr style="background-color: {bg};">'
                     f'<td style="padding: 7px 8px; color: #333; vertical-align: top; border-bottom: 1px solid #eee;">{i}. {esc(p.get("issue"))}</td>'
                     f'<td style="padding: 7px 8px; color: #333; vertical-align: top; border-bottom: 1px solid #eee;">{esc(p.get("cause"))}</td>'
                     f'<td style="padding: 7px 8px; color: #333; vertical-align: top; border-bottom: 1px solid #eee;">{esc(p.get("impact"))}</td>'
                     f'</tr>')
        h.append('</tbody></table>')

    suggestions = data.get('suggestions') or []
    if suggestions:
        h.append('<p style="margin: 16px 0 6px 0; font-size: 14px; color: #222;"><b>💡 纠偏措施建议</b></p>')
        h.append('<table style="width: 100%; border-collapse: collapse; font-size: 12.5px;">')
        h.append('<thead><tr style="background-color: #1a2740; color: #fff;">'
                 '<th style="padding: 6px 8px; text-align: center; width: 8%;">优先级</th>'
                 '<th style="padding: 6px 8px; text-align: left;">措施</th>'
                 '<th style="padding: 6px 8px; text-align: center; width: 18%;">建议责任方</th>'
                 '<th style="padding: 6px 8px; text-align: center; width: 16%;">完成时限</th></tr></thead><tbody>')
        pcolor = {'高': '#d92d20', '中': '#dc8a00', '低': '#00875a'}
        for s in suggestions:
            pri = str(s.get('priority') or '中').strip()
            pc = pcolor.get(pri, '#666')
            h.append(f'<tr>'
                     f'<td style="padding: 7px 8px; text-align: center; color: {pc}; font-weight: bold; border-bottom: 1px solid #eee;">{esc(pri)}</td>'
                     f'<td style="padding: 7px 8px; color: #333; border-bottom: 1px solid #eee;">{esc(s.get("action"))}</td>'
                     f'<td style="padding: 7px 8px; text-align: center; color: #333; border-bottom: 1px solid #eee;">{esc(s.get("owner"))}</td>'
                     f'<td style="padding: 7px 8px; text-align: center; color: #333; border-bottom: 1px solid #eee;">{esc(s.get("deadline"))}</td>'
                     f'</tr>')
        h.append('</tbody></table>')

    for key, title in (('resources', '🏗️ 人材机资源配置分析'), ('risks', '⚠️ 后续风险预判')):
        arr = data.get(key) or []
        if arr:
            h.append(f'<p style="margin: 16px 0 6px 0; font-size: 14px; color: #222;"><b>{title}</b></p>')
            h.append('<ul style="margin: 4px 0; padding-left: 22px; font-size: 13px; color: #444; line-height: 1.75;">')
            for it in arr:
                h.append(f'<li style="margin: 4px 0;">{esc(it)}</li>')
            h.append('</ul>')
        # 施工队分区对比紧跟在「人材机资源配置分析」之后
        if key == 'resources' and zone_html:
            h.append(zone_html)
    return '\n'.join(h)


_LLM_CACHE = {}   # 同一次运行内缓存大模型结果，避免重复调用（重复调用易触发限流导致静默降级）


def generate_llm_analysis(summary_results, subtasks, resources, upcoming, upcoming_keynodes, today,
                          fallback_result):
    """调用大模型生成分析；失败则返回 None（调用方使用规则版 fallback_result）"""
    # 缓存命中：直接复用本次运行已生成的结果，不再重复请求 API
    cache_key = (today.strftime('%Y-%m-%d'), len(summary_results), len(subtasks))
    if cache_key in _LLM_CACHE:
        cached = _LLM_CACHE[cache_key]
        if cached is not None:
            print('  ♻️ 复用本次运行的大模型分析结果（已缓存）')
            return cached
        print('  ♻️ 本次运行大模型已失败过，直接用规则版')
        return None

    ctx = build_llm_context(summary_results, subtasks, resources, upcoming, upcoming_keynodes, today)
    if not LLM_ENABLED:
        print('  ℹ️ 未配置 LLM_API_KEY，使用规则版分析')
    if os.environ.get('DUMP_CTX', '').strip().lower() in ('1', 'true', 'yes'):
        _p = os.path.join(OUTPUT_DIR, f'大模型输入数据_{today.strftime("%Y-%m-%d")}.txt')
        try:
            with open(_p, 'w', encoding='utf-8') as _f:
                _f.write(ctx)
            print(f'  🧾 大模型输入数据已导出: {_p}')
        except Exception as _e:
            print(f'  ⚠️ 导出上下文失败: {_e}')
    if not LLM_ENABLED:
        return None
    sys_p, user_p = build_llm_prompt(ctx)
    print(f'  🤖 调用大模型分析（{LLM_MODEL}）... 上下文 {len(ctx)} 字符')
    raw = call_llm(sys_p, user_p)
    if not raw:
        print('  ⚠️ 大模型无响应，回退规则版分析')
        _LLM_CACHE[cache_key] = None
        return None
    data = _extract_json(raw)
    if not data:
        print(f'  ⚠️ 大模型返回非预期JSON（前200字：{raw[:200]}），回退规则版分析')
        _LLM_CACHE[cache_key] = None
        return None
    print(f'  ✅ 大模型分析完成：问题{len(data.get("problems") or [])}条 / 建议{len(data.get("suggestions") or [])}条')
    _LLM_CACHE[cache_key] = data
    return data


def generate_ai_analysis(summary_results, subtasks, resources=None):
    """基于预警数据+人材机数据生成AI分析和建议"""
    red_subs = [st for st in subtasks if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_RED and st['pct'] < 1.0]
    paused_subs = [st for st in subtasks if st['status'] == '工序暂停']
    worst_sub = None
    for st in subtasks:
        if st['end_delay'] is not None and st['pct'] < 1.0:
            if worst_sub is None or st['end_delay'] > worst_sub['end_delay']:
                worst_sub = st
    red_summary = [r for r in summary_results if r['level'] == '红色']
    if red_subs or red_summary:
        analysis = f"当前工期整体承压：主任务红色预警{len(red_summary)}项、子任务严重滞后（延迟≥7天）{len(red_subs)}项，"
        if worst_sub:
            analysis += f"其中[{worst_sub['section']}]{worst_sub['name']}延迟最久（+{worst_sub['end_delay']}天），"
        analysis += "需重点攻坚。"
    elif any(r['level'] == '黄色' for r in summary_results):
        analysis = "当前工期存在局部滞后风险，建议关注黄色预警任务的推进节奏。"
    else:
        analysis = "当前工期整体可控，各主任务偏差在可接受范围内，建议保持现有节奏。"
    suggestions = []
    if worst_sub and worst_sub['end_delay'] is not None and worst_sub['end_delay'] >= TIME_WARN_RED:
        suggestions.append(f"优先攻坚[{worst_sub['section']}]{worst_sub['name']}（延迟{worst_sub['end_delay']}天、完成{worst_sub['pct']*100:.0f}%），制定专项赶工方案。")
    if paused_subs:
        names = '、'.join(f"[{st['section']}]{st['name']}" for st in paused_subs[:2])
        suggestions.append(f"协调{names}等暂停工序复工，明确责任人与时间节点。")
    if summary_results:
        evaluable = [r for r in summary_results if r['deviation'] is not None]
        if evaluable:
            worst = min(evaluable, key=lambda x: x['deviation'])
            if worst['deviation'] < 0:
                suggestions.append(f"关注{worst['daily_name']}整体进度（落后{abs(worst['dev_days'])}天），核查资源投入。")
    # 人材机资源洞察
    res_observations, res_suggestions = resource_insights(resources or {}, subtasks)
    if res_observations:
        analysis += ' 资源面：' + res_observations[0]
    suggestions.extend(res_suggestions)
    for ds in ["每日跟踪关键路径任务完成情况，提前识别偏差并干预。",
               "加强各参建单位协同，确保材料、机械、人员等要素及时到位。",
               "对滞后任务建立日清日结机制，及时纠偏。"]:
        if len(suggestions) >= 4:
            break
        if ds not in suggestions:
            suggestions.append(ds)
    return analysis, suggestions[:4], res_observations


# 前置检查：输入文件是否就绪（克隆仓库后为空目录时给出友好提示，而非抛异常）
_missing = []
if not os.path.exists(PLAN_FILE):
    _missing.append(f'总进度计划：{PLAN_FILE}\n'
                    f'     → 把计划表 xlsx 放进「master_plan/」目录，或用环境变量 PLAN_FILE 指定路径')
if not DAILY_FILE or not os.path.exists(DAILY_FILE):
    _missing.append(f'工程日报：目录 {DAILY_DIR}\n'
                    f'     → 把日报 xlsx 放进「daily_reports/」，文件名需含关键词：{"/".join(DAILY_FILE_KEYWORDS)}')
if _missing:
    print('❌ 缺少输入文件，无法运行：')
    for _m in _missing:
        print(f'   - {_m}')
    print('\n提示：首次使用请先阅读 README.md 的「输入表格格式约定」章节。')
    raise SystemExit(1)

# ============================================================
# 步骤1：读取进度计划
# ============================================================
print('📋 读取进度计划...')
plan_list = load_plan(PLAN_FILE)
print(f'  进度计划共 {len(plan_list)} 个任务')


# ============================================================
# 步骤1b：读取关键节点汇总
# ============================================================
print('📋 读取关键节点汇总...')
key_nodes = load_key_nodes(KEYNODE_FILE)
upcoming_keynodes = get_upcoming_keynodes(key_nodes, TODAY, days=15)
print(f'  关键节点共 {len(key_nodes)} 个，未来半个月内 {len(upcoming_keynodes)} 个')
for n in upcoming_keynodes:
    print(f'  ⭐ {n["date"].strftime("%m-%d")} [{n["category"]}] {n["name"]}')


# ============================================================
# 步骤2：读取日报（主任务总体行 + 1.2~1.6子任务）
# ============================================================
print('📋 读取日报...')
wb_daily = openpyxl.load_workbook(DAILY_FILE, data_only=True)
if DAILY_SHEET_NAME in wb_daily.sheetnames:
    ws_daily = wb_daily[DAILY_SHEET_NAME]
else:
    ws_daily = wb_daily.worksheets[0]
print(f'  使用工作表: {ws_daily.title}')

summary_rows, subtasks = parse_daily(ws_daily)

print(f'  主任务（总体行）共 {len(summary_rows)} 个:')
for sm in summary_rows:
    pct_str = f'{sm["pct"]*100:.1f}%' if sm['pct'] is not None else '-'
    print(f'    行{sm["row"]}: {sm["name"]} | {sm["status"]} | {pct_str}')

for st in subtasks:
    ed_str = f'{st["end_delay"]:+d}天' if st['end_delay'] is not None else '-'
    print(f'    [{st["section"]}][{st["block"]}] {st["name"]}: {st["status"]} | 完成{st["pct"]*100:.0f}% | 结束偏差:{ed_str}')
print(f'  子任务共 {len(subtasks)} 个')


# ============================================================
# 步骤2c：读取人材机数据
# ============================================================
print('📋 读取人材机数据...')
resources = parse_resources(ws_daily)
for sec_name, res in resources.items():
    labor_total = res['labor_total'] or sum(c for _, c in res['labor'])
    mach_total = sum(c for _, c in res['machinery'])
    print(f'  [{sec_name}] 人员 {labor_total}人（特殊工种{res["special"]}人）| 机械 {mach_total}台 | 材料 {len(res["materials"])}项')
if not resources:
    print('  ⚠️ 未识别到人材机数据')


# ============================================================
# 步骤3A：主任务汇总比对（总体行 vs 进度计划）
# ============================================================
print('📋 计算主任务汇总偏差（vs进度计划）...')
summary_results = []

def make_neutral_result(sm, level='无数据', source='-'):
    """构造不参与量化比对的总体行（无完成比例/无计划日期/尚未开始），保证总体行一个不漏"""
    return {
        'daily_name': sm['name'],
        'plan_name': source,
        'plan_start': sm.get('plan_start'),
        'plan_end': sm.get('plan_end'),
        'total_days': None,
        'elapsed_days': None,
        'plan_pct': None,
        'actual_pct': sm['pct'],
        'deviation': None,
        'dev_days': None,
        'level': level,
        'status': sm['status'],
    }


for sm in summary_rows:
    # 无量化完成比例的行（如手续办理/静载试验）：保留输出，标记为无数据
    if sm['pct'] is None:
        print(f'  ⚪ {sm["name"]}: 无完成比例数据，保留输出（不参与百分比比对）')
        summary_results.append(make_neutral_result(sm))
        continue

    plan_task = match_plan_task(sm['name'], plan_list)
    if plan_task:
        ref_start, ref_end = plan_task['start'], plan_task['end']
        total_days = plan_task['duration']
        source = plan_task['name']
    else:
        # 回退：用日报自身计划日期
        if sm['plan_start'] and sm['plan_end']:
            ref_start, ref_end = sm['plan_start'], sm['plan_end']
            total_days = max((ref_end - ref_start).days, 1)
            source = '日报计划日期'
        else:
            print(f'  ⚪ {sm["name"]}: 未匹配进度计划且日报无计划日期，保留输出')
            summary_results.append(make_neutral_result(sm))
            continue

    if total_days <= 0:
        summary_results.append(make_neutral_result(sm))
        continue
    elapsed = (TODAY - ref_start).days
    if elapsed < 0:
        print(f'  ⚪ {sm["name"]}: 参照任务尚未开始（{source}），保留输出')
        summary_results.append(make_neutral_result(sm, level='未开始', source=source))
        continue

    plan_pct = min(elapsed / total_days, 1.0)
    actual_pct = sm['pct']
    deviation = actual_pct - plan_pct
    dev_days = round(deviation * total_days)
    level = warning_level(deviation)

    sr = {
        'daily_name': sm['name'],
        'plan_name': source,
        'plan_start': ref_start,
        'plan_end': ref_end,
        'total_days': total_days,
        'elapsed_days': elapsed,
        'plan_pct': plan_pct,
        'actual_pct': actual_pct,
        'deviation': deviation,
        'dev_days': dev_days,
        'level': level,
        'status': sm['status'],
    }
    summary_results.append(sr)

    icon = {'红色': '🔴', '黄色': '🟡', '绿色': '🟢'}[level]
    print(f'  {icon} {sm["name"]} → {source}')
    print(f'     计划: {plan_pct*100:.0f}% | 实际: {actual_pct*100:.0f}% | 偏差: {deviation*100:+.1f}% ({dev_days:+d}天)')

summary_results.sort(key=lambda x: (x['deviation'] is None, x['deviation'] if x['deviation'] is not None else 0))


# ============================================================
# 步骤3B：子任务时间偏差统计
# ============================================================
print('📋 统计子任务时间偏差...')

subtask_groups = {}
for st in subtasks:
    subtask_groups.setdefault(st['section'], []).append(st)

sub_red = sum(1 for st in subtasks if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_RED and st['pct'] < 1.0)
sub_yellow = sum(1 for st in subtasks if st['end_delay'] is not None and TIME_WARN_YELLOW <= st['end_delay'] < TIME_WARN_RED and st['pct'] < 1.0)
sub_green = sum(1 for st in subtasks if st['end_delay'] is not None and st['end_delay'] < TIME_WARN_YELLOW)
sub_none = len(subtasks) - sub_red - sub_yellow - sub_green

for sec, items in sorted(subtask_groups.items()):
    print(f'\n  --- {sec} {items[0]["category"]} ({len(items)}子任务) ---')
    for st in items:
        sd_str = f'{st["start_delay"]:+d}天' if st['start_delay'] is not None else 'N/A'
        ed_str = f'{st["end_delay"]:+d}天' if st['end_delay'] is not None else 'N/A'
        icon = '🔴' if (st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_RED and st['pct'] < 1.0) else \
               '🟡' if (st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_YELLOW and st['pct'] < 1.0) else '🟢'
        print(f'    {icon} [{st["block"]}] {st["name"]}: 状态={st["status"]}, 开始偏差={sd_str}, 结束偏差={ed_str}, 完成{st["pct"]*100:.0f}%')


# ============================================================
# 步骤4：生成Excel预警报表
# ============================================================
print('\n📋 生成预警报表...')

wb_out = openpyxl.Workbook()

# --- 样式 ---
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
    'none': sum(1 for r in summary_results if r['level'] not in ('红色', '黄色', '绿色')),
}


# ============================================================
# Sheet1: 主任务汇总比对（总体行 vs进度计划）
# ============================================================
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
        r['plan_start'].strftime('%Y-%m-%d') if r['plan_start'] else '-',
        r['plan_end'].strftime('%Y-%m-%d') if r['plan_end'] else '-',
        r['total_days'] if r['total_days'] is not None else '-',
        r['elapsed_days'] if r['elapsed_days'] is not None else '-',
        f"{r['plan_pct']*100:.1f}%" if r['plan_pct'] is not None else '-',
        f"{r['actual_pct']*100:.1f}%" if r['actual_pct'] is not None else '-',
        f"{r['deviation']*100:+.1f}%" if r['deviation'] is not None else '-',
        f"{r['dev_days']:+d}天" if r['dev_days'] is not None else '-',
        r['level'],
    ]
    ws1.append(row_data)
    row_idx = ws1.max_row

    if r['level'] == '红色':
        fill, font = red_fill, red_font
    elif r['level'] == '黄色':
        fill, font = yellow_fill, yellow_font
    elif r['level'] == '绿色':
        fill, font = green_fill, green_font
    else:
        fill, font = PatternFill(start_color='2A3A50', end_color='2A3A50', fill_type='solid'), gray_font

    for col in range(1, len(row_data) + 1):
        cell = ws1.cell(row=row_idx, column=col)
        cell.fill = fill
        cell.border = thin_border
        cell.alignment = center_align
        cell.font = data_font if col in [1, 2] else (font if col == 11 else num_font)

col_widths = [16, 28, 12, 12, 10, 10, 10, 10, 10, 10, 10]
for i, w in enumerate(col_widths, 1):
    ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


# ============================================================
# Sheet2: 子任务时间偏差明细（1.2~1.6 全部）
# ============================================================
ws2 = wb_out.create_sheet('子任务时间偏差')
ws2.sheet_properties.tabColor = 'FF8800'

sub_cols = ['分节', '分类', '区段', '子任务名称', '单位', '总量',
            '计划开始', '计划结束', '实际开始', '实际结束', '状态',
            '开始偏差(天)', '结束偏差(天)', '完成%', '时间预警']
ws2.append(sub_cols)
for col in range(1, len(sub_cols) + 1):
    cell = ws2.cell(row=1, column=col)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = thin_border

# 按分节（1.2→1.6）分组输出
last_section = ''
for st in sorted(subtasks, key=lambda x: (x['section'], x['row'])):
    # 分节分组标题行
    if st['section'] != last_section:
        last_section = st['section']
        ws2.append([f'▎ {st["section"]} {st["category"]}'] + [''] * (len(sub_cols) - 1))
        row_idx = ws2.max_row
        ws2.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=len(sub_cols))
        cell = ws2.cell(row=row_idx, column=1)
        cell.font = Font(name='Microsoft YaHei', bold=True, size=10, color='FFFFFF')
        cell.fill = sub_header_fill
        cell.border = thin_border
        for c in range(2, len(sub_cols) + 1):
            ws2.cell(row=row_idx, column=c).fill = sub_header_fill
            ws2.cell(row=row_idx, column=c).border = thin_border

    # 时间预警等级
    ed = st['end_delay']
    if ed is not None and st['pct'] < 1.0:
        t_level = time_warning_level(ed)
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
        st['section'], st['category'], st['block'], st['name'], st['unit'],
        st['total_qty'] if st['total_qty'] else '-',
        start_plan_str, end_plan_str, actual_start_str, actual_end_str,
        st['status'], sd_str, ed_str, pct_str, t_level,
    ]
    ws2.append(row_data)
    row_idx = ws2.max_row

    # 着色
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
        if col in (4, 5, 11):
            cell.font = white_font if fill == PatternFill() else data_font
        elif col in (12, 13, 15):
            cell.font = font if fill != PatternFill() else white_font
        else:
            cell.font = white_font if fill == PatternFill() else num_font
        cell.alignment = center_align if col not in (2, 4) else wrap_align

sub_col_widths = [6, 12, 10, 20, 6, 8, 10, 10, 10, 10, 10, 12, 12, 8, 10]
for i, w in enumerate(sub_col_widths, 1):
    ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


# ============================================================
# Sheet3: 统计摘要
# ============================================================
ws3 = wb_out.create_sheet('统计摘要')
ws3.sheet_properties.tabColor = '00AA88'

summary_data = [
    ['工期预警统计摘要', '', '', ''],
    ['统计日期', TODAY.strftime('%Y-%m-%d'), '', ''],
    ['数据来源（进度计划）', PLAN_FILE.split('\\')[-1], '', ''],
    ['数据来源（日报）', os.path.basename(DAILY_FILE), '', ''],
    ['', '', '', ''],
    ['一、主任务汇总预警（总体行 vs进度计划）', '', '', ''],
    ['预警等级', '数量', '任务', '偏差天数'],
    ['🔴 红色', summary_counts['red'], '', ''],
    ['🟡 黄色', summary_counts['yellow'], '', ''],
    ['🟢 绿色', summary_counts['green'], '', ''],
    ['⚪ 无数据/未开始', summary_counts['none'], '', ''],
    ['合计', len(summary_results), '', ''],
]
for r in summary_results:
    dev_str = f'{r["dev_days"]:+d}天' if r['dev_days'] is not None else '-'
    summary_data.append([f'  {r["level"]}', '', f'{r["daily_name"]} → {r["plan_name"]}', dev_str])

summary_data.append(['', '', '', ''])
summary_data.append(['二、子任务时间偏差统计（1.2~1.6，日报内部比对）', '', '', ''])
summary_data.append(['时间预警', '数量', '', ''])
summary_data.append(['🔴 延迟≥7天', sub_red, '', ''])
summary_data.append(['🟡 延迟3-6天', sub_yellow, '', ''])
summary_data.append(['🟢 正常', sub_green, '', ''])
summary_data.append(['- 已完成/未开始', sub_none, '', ''])
summary_data.append(['合计子任务', len(subtasks), '', ''])

# 列出红色预警的子任务
red_subtasks = [st for st in subtasks if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_RED and st['pct'] < 1.0]
if red_subtasks:
    summary_data.append(['', '', '', ''])
    summary_data.append(['🔴 严重滞后子任务明细', '', '', ''])
    for st in red_subtasks:
        summary_data.append(['', f'[{st["section"]}/{st["block"]}] {st["name"]}',
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

# 保存（文件被占用时自动改用带序号的文件名）
try:
    wb_out.save(OUTPUT_FILE)
except PermissionError:
    base, ext = os.path.splitext(OUTPUT_FILE)
    OUTPUT_FILE = f'{base}_v2{ext}'
    print(f'  ⚠️ 原报表文件被占用，改存为: {os.path.basename(OUTPUT_FILE)}')
    wb_out.save(OUTPUT_FILE)
print(f'\n✅ 预警报表已生成: {OUTPUT_FILE}')
print(f'   主任务: {len(summary_results)} 个 | 🔴{summary_counts["red"]} 🟡{summary_counts["yellow"]} 🟢{summary_counts["green"]} ⚪{summary_counts["none"]}')
print(f'   子任务: {len(subtasks)} 个 | 🔴{sub_red} 🟡{sub_yellow} 🟢{sub_green} (-{sub_none})')


# ============================================================
# 步骤5：发送邮件预警
# ============================================================
def generate_email_html(summary_results, subtasks_all, today, resources=None, upcoming_keynodes=None):
    red_count = sum(1 for r in summary_results if r['level'] == '红色')
    yellow_count = sum(1 for r in summary_results if r['level'] == '黄色')
    green_count = sum(1 for r in summary_results if r['level'] == '绿色')
    none_count = len(summary_results) - red_count - yellow_count - green_count

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
                <p style="margin: 5px 0; font-size: 14px; color: #666;">统计日期：{today.strftime('%Y-%m-%d')} | 日报：{os.path.basename(DAILY_FILE)}</p>
                <p style="margin: 5px 0; font-size: 14px;">
                    <b>主任务汇总（{len(summary_results)}项）：</b>
                    <span style="color: #ff4060;">🔴 {red_count}</span> |
                    <span style="color: #ffaa00;">🟡 {yellow_count}</span> |
                    <span style="color: #00aa66;">🟢 {green_count}</span>
                    <span style="color: #8899aa;">| ⚪ 无数据/未开始 {none_count}</span>
                </p>
            </div>
    """

    # 一、主任务汇总表
    html += """
            <h3 style="color: #1a2740;">一、主任务汇总比对（总体行 vs进度计划）</h3>
            <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px;">
                <thead>
                    <tr style="background-color: #1a2740; color: white;">
                        <th style="padding: 10px 16px; text-align: left;">任务</th>
                        <th style="padding: 10px 16px; text-align: center;">计划%</th>
                        <th style="padding: 10px 16px; text-align: center;">实际%</th>
                        <th style="padding: 10px 16px; text-align: center;">偏差%</th>
                        <th style="padding: 10px 16px; text-align: center;">偏差天数</th>
                        <th style="padding: 10px 16px; text-align: center;">等级</th>
                    </tr>
                </thead>
                <tbody>
    """
    for r in summary_results:
        level_map = {'红色': ('#3d1520', '#ff4060'), '黄色': ('#3d3015', '#ffaa00'), '绿色': ('#153d20', '#00aa66')}
        bg, tc = level_map.get(r['level'], ('#2a3a50', '#8899aa'))
        plan_pct_s = f"{r['plan_pct']*100:.0f}%" if r['plan_pct'] is not None else '-'
        actual_pct_s = f"{r['actual_pct']*100:.0f}%" if r['actual_pct'] is not None else '-'
        dev_pct_s = f"{r['deviation']*100:+.1f}%" if r['deviation'] is not None else '-'
        dev_days_s = f"{r['dev_days']:+d}天" if r['dev_days'] is not None else '-'
        html += f"""
                    <tr style="background-color: {bg};">
                        <td style="padding: 8px 16px; color: white;">{r['daily_name']}</td>
                        <td style="padding: 8px 16px; text-align: center; color: white;">{plan_pct_s}</td>
                        <td style="padding: 8px 16px; text-align: center; color: white;">{actual_pct_s}</td>
                        <td style="padding: 8px 16px; text-align: center; color: {tc}; font-weight: bold;">{dev_pct_s}</td>
                        <td style="padding: 8px 16px; text-align: center; color: {tc}; font-weight: bold;">{dev_days_s}</td>
                        <td style="padding: 8px 16px; text-align: center; color: {tc}; font-weight: bold;">{r['level']}</td>
                    </tr>
        """
    html += """
                </tbody>
            </table>
    """

    # 二、子任务时间偏差预警（只列红色/黄色）
    delayed_subs = [st for st in subtasks_all
                    if st['end_delay'] is not None and st['end_delay'] >= TIME_WARN_YELLOW and st['pct'] < 1.0]
    html += f"""
            <h3 style="color: #000000; margin-top: 30px;">二、子任务时间偏差预警（总任务{len(subtasks_all)} / 🔴红色{sub_red} / 🟡黄色{sub_yellow} / 🟢绿色{sub_green} / 已完成或未开始{sub_none}）</h3>
            <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 12px;">
                <thead>
                    <tr style="background-color: #1a2740; color: white;">
                        <th style="padding: 8px 12px; text-align: left;">分节</th>
                        <th style="padding: 8px 12px; text-align: left;">子任务</th>
                        <th style="padding: 8px 12px; text-align: center;">计划结束</th>
                        <th style="padding: 8px 12px; text-align: center;">延迟天数</th>
                        <th style="padding: 8px 12px; text-align: center;">完成%</th>
                        <th style="padding: 8px 12px; text-align: center;">状态</th>
                        <th style="padding: 8px 12px; text-align: center;">预警</th>
                    </tr>
                </thead>
                <tbody>
    """
    level_style = {'红色': ('#3d1520', '#ff4060'), '黄色': ('#3d3015', '#ffaa00')}
    if delayed_subs:
        for st in sorted(delayed_subs, key=lambda x: (x['section'], -(x['end_delay'] or 0))):
            t_level = time_warning_level(st['end_delay'])
            bg, tc = level_style[t_level]
            plan_end_str = st['plan_end'].strftime('%m-%d') if st['plan_end'] else '-'
            html += f"""
                    <tr style="background-color: {bg};">
                        <td style="padding: 6px 12px; color: white;">{st['section']}</td>
                        <td style="padding: 6px 12px; color: white;">{st['name']}</td>
                        <td style="padding: 6px 12px; text-align: center; color: white;">{plan_end_str}</td>
                        <td style="padding: 6px 12px; text-align: center; color: {tc}; font-weight: bold;">+{st['end_delay']}天</td>
                        <td style="padding: 6px 12px; text-align: center; color: white;">{st['pct']*100:.0f}%</td>
                        <td style="padding: 6px 12px; text-align: center; color: white;">{st['status'] or '-'}</td>
                        <td style="padding: 6px 12px; text-align: center; color: {tc}; font-weight: bold;">{t_level}</td>
                    </tr>
            """
    else:
        html += """
                    <tr>
                        <td colspan="7" style="padding: 10px; text-align: center; color: #888;">当前无黄色及以上预警子任务</td>
                    </tr>
        """
    html += """
                </tbody>
            </table>
    """

    # 未来一周计划启动任务（提前计算，供大模型分析与后续板块共用）
    upcoming = get_upcoming_starts(plan_list, today, days=7)

    # 施工队分区进度对比（按 ZONE_TEAM 配置的分区划分）
    zone_html = ''
    try:
        team_stats = compare_zone_teams(subtasks_all, today)
        if team_stats:
            cat_rows, cat_teams = zone_team_by_category(subtasks_all, today)
            zone_insights = zone_team_insights(team_stats, today)
            zone_html = render_zone_team_html(team_stats, cat_rows, cat_teams, zone_insights, today)
            for t, s in team_stats.items():
                fr = '%.0f%%' % (s['fulfill'] * 100) if s['fulfill'] is not None else '-'
                sp = '%.0f㎡/天' % s['speed_med'] if s['speed_med'] else '-'
                print(f'  🚧 {t}（{"/".join(s["zones"])}区）：到期履约率 {fr}（{s["due_done"]}/{s["due"]}）'
                      f' | 强度中位 {sp} | 到期未完成 {s["due_late"]} 项')
    except Exception as e:
        print(f'  ⚠️ 施工队分区对比生成失败（不影响其余报表）：{e}')
        zone_html = ''

    # 三、AI分析与建议（优先调用大模型，失败自动回退规则版）
    ai_analysis, ai_suggestions, res_observations = generate_ai_analysis(summary_results, subtasks_all, resources)
    llm_data = None
    try:
        llm_data = generate_llm_analysis(summary_results, subtasks_all, resources,
                                         upcoming, upcoming_keynodes, today,
                                         fallback_result=(ai_analysis, ai_suggestions, res_observations))
    except Exception as e:
        print(f'  ⚠️ 大模型分析异常，回退规则版：{e}')
        llm_data = None

    used_llm = bool(llm_data)
    src_tag = ('大模型深度分析 · ' + LLM_MODEL) if used_llm else '规则引擎分析'
    html += f"""
            <h3 style="color: #1a2740; margin-top: 30px;">三、AI分析与建议
                <span style="font-size: 11px; font-weight: normal; color: #888; background-color: #f0f2f5;
                             padding: 2px 8px; border-radius: 10px; vertical-align: middle;">{src_tag}</span>
            </h3>
            <div style="background-color: #f7f9fc; padding: 16px 18px; border-radius: 6px; margin: 15px 0;
                        border-left: 4px solid #4a6cf7;">
    """
    if used_llm:
        html += render_llm_html(llm_data, zone_html)
    else:
        html += f"""
                <p style="margin: 5px 0; font-size: 14px; color: #333; line-height: 1.7;"><b>📊 分析：</b>{ai_analysis}</p>
        """
        if res_observations:
            html += ('<p style="margin: 10px 0 5px 0; font-size: 14px; color: #333;"><b>🏗️ 人材机资源观察：</b></p>'
                     '<ul style="margin: 5px 0; padding-left: 20px; font-size: 13px; color: #555; line-height: 1.7;">')
            for obs in res_observations:
                html += f'<li style="margin: 5px 0;">{obs}</li>'
            html += '</ul>'
        if zone_html:
            html += zone_html
        html += ('<p style="margin: 10px 0 5px 0; font-size: 14px; color: #333;"><b>💡 建议：</b></p>'
                 '<ul style="margin: 5px 0; padding-left: 20px; font-size: 13px; color: #555; line-height: 1.7;">')
        for sug in ai_suggestions:
            html += f'<li style="margin: 5px 0;">{sug}</li>'
        html += '</ul>'
    html += """
            </div>
    """

    # 未来一周计划启动任务（基于总进度计划）
    week_end = today + timedelta(days=7)
    html += f"""
            <h4 style="color: #1a2740; margin: 20px 0 8px 0;">📅 未来一周计划启动任务（{today.strftime('%m-%d')} ~ {week_end.strftime('%m-%d')}）</h4>
    """
    if upcoming:
        html += f"""
            <div style="background-color: #f0fff6; padding: 15px; border-radius: 5px; margin: 10px 0; border-left: 4px solid #00aa66;">
                <p style="margin: 5px 0 10px 0; font-size: 13px; color: #555;">共 {len(upcoming)} 项任务将在未来一周内按计划启动，请提前落实资源与前置条件：</p>
                <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                    <thead>
                        <tr style="background-color: #1a2740; color: white;">
                            <th style="padding: 6px; text-align: left;">任务名称</th>
                            <th style="padding: 6px; text-align: center;">计划开始</th>
                            <th style="padding: 6px; text-align: center;">计划完成</th>
                            <th style="padding: 6px; text-align: center;">工期(天)</th>
                            <th style="padding: 6px; text-align: center;">距今</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        for p in upcoming:
            days_left = (p['start'] - today).days
            start_tag = '今天启动' if days_left == 0 else f'{days_left}天后'
            bg = '#e8f5ee' if days_left <= 1 else '#ffffff'
            html += f"""
                        <tr style="background-color: {bg};">
                            <td style="padding: 6px; color: #333;">{p['name']}</td>
                            <td style="padding: 6px; text-align: center; color: #333;">{p['start'].strftime('%m-%d')}</td>
                            <td style="padding: 6px; text-align: center; color: #333;">{p['end'].strftime('%m-%d')}</td>
                            <td style="padding: 6px; text-align: center; color: #333;">{p['duration']}</td>
                            <td style="padding: 6px; text-align: center; color: #00aa66; font-weight: bold;">{start_tag}</td>
                        </tr>
            """
        html += """
                    </tbody>
                </table>
            </div>
        """
    else:
        html += """
            <div style="background-color: #f8f9fa; padding: 12px; border-radius: 5px; margin: 10px 0; font-size: 13px; color: #888;">
                未来一周内总进度计划中没有新的任务启动安排。
            </div>
        """

    # 未来半个月关键节点提醒（基于关键节点汇总表）
    halfmonth_end = today + timedelta(days=15)
    html += f"""
            <h4 style="color: #1a2740; margin: 20px 0 8px 0;">⭐ 未来半个月关键节点提醒（{today.strftime('%m-%d')} ~ {halfmonth_end.strftime('%m-%d')}）</h4>
    """
    if upcoming_keynodes:
        html += f"""
            <div style="background-color: #fff8f0; padding: 15px; border-radius: 5px; margin: 10px 0; border-left: 4px solid #ff8800;">
                <p style="margin: 5px 0 10px 0; font-size: 13px; color: #555;">共 {len(upcoming_keynodes)} 个关键节点将在未来半个月内到期，请提前安排落实：</p>
                <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                    <thead>
                        <tr style="background-color: #1a2740; color: white;">
                            <th style="padding: 6px; text-align: left;">关键节点名称</th>
                            <th style="padding: 6px; text-align: center;">类别</th>
                            <th style="padding: 6px; text-align: center;">节点日期</th>
                            <th style="padding: 6px; text-align: center;">距今</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        for n in upcoming_keynodes:
            days_left = (n['date'] - today).days
            node_tag = '今天到期' if days_left == 0 else f'{days_left}天后'
            tag_color = '#ff4060' if days_left <= 1 else '#ff8800'
            bg = '#ffe8e8' if days_left <= 1 else '#ffffff'
            html += f"""
                        <tr style="background-color: {bg};">
                            <td style="padding: 6px; color: #333;">{n['name']}</td>
                            <td style="padding: 6px; text-align: center; color: #666;">{n['category']}</td>
                            <td style="padding: 6px; text-align: center; color: #333;">{n['date'].strftime('%m-%d')}</td>
                            <td style="padding: 6px; text-align: center; color: {tag_color}; font-weight: bold;">{node_tag}</td>
                        </tr>
            """
        html += """
                    </tbody>
                </table>
            </div>
        """
    else:
        html += """
            <div style="background-color: #f8f9fa; padding: 12px; border-radius: 5px; margin: 10px 0; font-size: 13px; color: #888;">
                未来半个月内没有关键节点到期。
            </div>
        """

    html += f"""
            <div style="margin-top: 30px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; font-size: 12px; color: #999;">
                <p style="margin: 5px 0;">此邮件由工期预警管控系统自动生成</p>
                <p style="margin: 5px 0;">数据来源：{os.path.basename(DAILY_FILE)}</p>
                <p style="margin: 5px 0;">主任务汇总比对 vs 进度计划 | 子任务时间偏差仅列预警项</p>
                <p style="margin: 5px 0;">详细报表（含全部子任务）请查看附件Excel</p>
            </div>
        </div>
    </body>
    </html>
    """
    return html, subject_prefix


def send_warning_email(summary_results, subtasks_list, today, output_file, resources=None,
                       upcoming_keynodes=None, prebuilt_html=None):
    try:
        # 若已生成过 HTML（如导出预览时），直接复用，避免重复调用大模型
        if prebuilt_html is not None:
            html_body, subject_prefix = prebuilt_html
        else:
            html_body, subject_prefix = generate_email_html(summary_results, subtasks_list, today,
                                                            resources, upcoming_keynodes)

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

        print(f'\n📧 正在发送邮件...')
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
        server.quit()
        print(f'✅ 邮件已发送至: {", ".join(EMAIL_RECIPIENTS)}')
        return True

    except Exception as e:
        print(f'❌ 邮件发送失败: {str(e)}')
        return False


# 邮件HTML只生成一次（含大模型调用），预览与发送共用，避免重复调用
_prebuilt_html = None
if summary_results and (SEND_EMAIL or os.environ.get('DUMP_HTML', '').strip().lower() in ('1', 'true', 'yes')):
    try:
        _prebuilt_html = generate_email_html(summary_results, subtasks, TODAY, resources, upcoming_keynodes)
    except Exception as _e:
        print(f'\n⚠️ 生成邮件HTML失败: {_e}')
        _prebuilt_html = None

# 调试：导出邮件HTML预览文件（设环境变量 DUMP_HTML=1 开启）
if os.environ.get('DUMP_HTML', '').strip().lower() in ('1', 'true', 'yes') and summary_results:
    try:
        if _prebuilt_html is not None:
            _dump = os.path.join(OUTPUT_DIR, f'邮件预览_{TODAY.strftime("%Y-%m-%d")}.html')
            with open(_dump, 'w', encoding='utf-8') as _f:
                _f.write(_prebuilt_html[0])
            print(f'\n🧾 邮件HTML预览已导出: {_dump}')
        else:
            print('\n⚠️ HTML生成失败，跳过预览导出')
    except Exception as _e:
        print(f'\n⚠️ 导出HTML预览失败: {_e}')

if SEND_EMAIL and summary_results:
    email_success = send_warning_email(summary_results, subtasks, TODAY, OUTPUT_FILE, resources,
                                       upcoming_keynodes, prebuilt_html=_prebuilt_html)
else:
    if not SEND_EMAIL:
        print('\n⚠️ 邮件发送已禁用 (SEND_EMAIL=False)')
    elif not summary_results:
        print('\n⚠️ 无预警数据，跳过邮件发送')

print(f'\n🎉 工期预警任务完成！')
