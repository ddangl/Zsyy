# 新区域打卡数据处理逻辑（衍生版）
#
# 与一区版 processor.py 的区别：
# - 上班/下班时间都是完整日期时间，时长直接相减，无需模糊时间解析
# - 统计目标日 = 文件名第一天 D1，从所有日期 Sheet 取 上班时间.date == D1 的行
# - 同工号多条打卡保留最新一条；未参与 / 工号=无 的行不显示
# - 只分两组：在班（按时长升序）、接班（不标色）
# - 处理是纯过滤，不回写 Excel（一区版会把合并结果写回 Sheet1）

import os
import re
import sys
from datetime import datetime, timedelta

import openpyxl
import pandas as pd
from jinja2 import Template
from tabulate import tabulate

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

import config2 as config

COL_CHECKIN = config.COL_CHECKIN
COL_NAME = config.COL_NAME
COL_ID = config.COL_ID
COL_START = config.COL_START
COL_END = config.COL_END
COL_LOCATION = config.COL_LOCATION
COL_HANDOVER = config.COL_HANDOVER
COL_SHIFT_TODAY = config.COL_SHIFT_TODAY
COL_TOMORROW = config.COL_TOMORROW

CATEGORY_ORDER = config.CATEGORY_ORDER

DATE_SHEET_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')


# ============ 基础解析 ============

def _target_date_from_filename(file_path):
    """从文件名取统计目标日：`2026-08-20至...` → 2026-08-20"""
    base = os.path.splitext(os.path.basename(file_path))[0]
    date_part = base.split('_')[0]
    if '至' in date_part:
        date_str = date_part.split('至')[0]
    else:
        m = re.search(r'(\d{4}-\d{2}-\d{2})', date_part)
        date_str = m.group(1) if m else None
    if date_str:
        try:
            return datetime.strptime(date_str, '%Y-%m-%d').date()
        except ValueError:
            return None
    return None


def _resolve_target_date(file_path):
    """统计目标日：优先文件名，回退到第一个日期 Sheet 的日期"""
    date = _target_date_from_filename(file_path)
    if date is None:
        sheets = _date_sheets(file_path)
        if sheets:
            date = datetime.strptime(str(sheets[0]).strip(), '%Y-%m-%d').date()
    return date


def _date_sheets(file_path):
    """找出所有 YYYY-MM-DD 命名的数据 Sheet（自动跳过汇总 Sheet）"""
    xl = pd.ExcelFile(file_path)
    return [s for s in xl.sheet_names if DATE_SHEET_RE.match(str(s).strip())]


def _clean_checkin_time(value):
    """打卡时间：去掉（已修改填写信息）等括号后缀后解析，用于去重排序"""
    if pd.isna(value):
        return pd.NaT
    text = str(value).split('（')[0].split('(')[0].strip()
    if not text or text == '未参与':
        return pd.NaT
    return pd.to_datetime(text, errors='coerce')


def _parse_dt(value, default_date=None):
    """解析上班/下班时间：完整日期时间、datetime 对象、或 'HH:MM'（补 default_date 当天）"""
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value)
    text = str(value).strip()
    if not text or text == '无':
        return pd.NaT
    if default_date is not None and re.match(r'^\d{1,2}:\d{2}$', text):
        text = f"{default_date.strftime('%Y-%m-%d')} {text}"
    return pd.to_datetime(text, errors='coerce')


def _norm_id(value):
    """工号统一为字符串（Excel 数值型工号转成整数字符串）"""
    if pd.isna(value):
        return ''
    if isinstance(value, float) and float(value).is_integer():
        return str(int(value))
    return str(value).strip()


def _shift_info(text):
    """解析新表单的单选项字符串。

    返回 (类型, 房间)：
      '白班：（请补充工作地点手术室）：12' → ('白班', '12')
      '白班（请补充工作地点手术室）：13副麻二' → ('白班', '13副麻二')
      '16点接班' → ('接班', '')
      '值班' → ('值班', '')
      空/无/未参与 → (None, '')
    """
    t = '' if text is None or pd.isna(text) else str(text).strip()
    if not t or t in ('无', '未参与'):
        return (None, '')
    if t.startswith('白班'):
        room = re.split(r'[：:]', t)[-1].strip()
        return ('白班', room)
    if '值班' in t:
        return ('值班', '')
    if '接班' in t:
        return ('接班', '')
    return (None, t)


def get_time_category(work_hours):
    if isinstance(work_hours, (int, float)):
        for category, (low, high) in config.WORK_HOURS_THRESHOLDS.items():
            if low <= work_hours < high:
                return category
    return 'normal'


# ============ 读取与合并 ============

def read_excel_file(file_path):
    try:
        sheets = _date_sheets(file_path)
        if not sheets:
            print("未找到日期命名的 Sheet（YYYY-MM-DD）")
            return None

        target_date = _resolve_target_date(file_path)
        if target_date is None:
            print("无法确定统计目标日期（文件名和 Sheet 名都没有日期）")
            return None
        print(f"统计目标日期: {target_date}（数据Sheet: {sheets}）")

        frames = []
        is_v2 = False
        for s in sheets:
            df = pd.read_excel(file_path, sheet_name=s)
            # 新表单列名归一化：上班时间（24小时制）→ 上班时间
            if config.COL_START_V2 in df.columns:
                df = df.rename(columns={config.COL_START_V2: COL_START})
            if config.COL_END_V2 in df.columns:
                df = df.rename(columns={config.COL_END_V2: COL_END})
            if COL_TOMORROW in df.columns:
                is_v2 = True
            if COL_ID not in df.columns or COL_START not in df.columns:
                continue
            sheet_date = datetime.strptime(str(s).strip(), '%Y-%m-%d').date()
            df = df.copy()
            df['_sheet_date'] = sheet_date
            frames.append(df)
        if not frames:
            print(f"Sheet 中缺少 {COL_ID} / {COL_START} 列")
            return None
        print(f"表单格式: {'新版（今日班次+明日工作地点）' if is_v2 else '旧版（是否4p副麻接班）'}")

        raw = pd.concat(frames, ignore_index=True)
        print(f"读取 {len(raw)} 行原始数据")

        # 剔除未参与 / 无工号 行（不显示）
        raw['_id'] = raw[COL_ID].apply(_norm_id)
        raw = raw[(raw['_id'] != '') & (raw['_id'].str.lower() != '无')]
        if COL_CHECKIN in raw.columns:
            raw = raw[~raw[COL_CHECKIN].astype(str).str.contains('未参与', na=False)]
        print(f"剔除未参与/无工号行后: {len(raw)} 行")

        # 解析上班/下班时间（'HH:MM' 形式按所在 Sheet 日期补全）
        raw['_start'] = pd.to_datetime(
            raw.apply(lambda r: _parse_dt(r[COL_START], r['_sheet_date']), axis=1), errors='coerce')
        raw['_end'] = pd.to_datetime(
            raw.apply(lambda r: _parse_dt(r[COL_END], r['_sheet_date']), axis=1), errors='coerce')

        # 只保留上班时间在目标日的班次（其余 Sheet 里的前一天补卡行自动排除）
        raw = raw[raw['_start'].notna() & (raw['_start'].dt.date == target_date)]
        print(f"上班时间在 {target_date} 的行: {len(raw)} 行")

        # 打了上班卡但下班时间缺失/非法 → 收集错误，提示用「修改表格」修正
        bad_end = raw[raw['_end'].isna()]
        if not bad_end.empty:
            bad = bad_end[[COL_NAME, COL_ID, COL_START, COL_END]].copy()
            bad[COL_START] = bad[COL_START].astype(str)
            bad[COL_END] = bad[COL_END].astype(str)
            msg = "\n发现下班时间缺失或格式不合法！\n" + "=" * 50 + "\n以下记录需通过「修改表格」修正：\n"
            msg += tabulate(bad, headers='keys', tablefmt='grid', showindex=False)
            raise ValueError(msg)

        # 同一工号多条 → 保留打卡时间最新的一条
        if COL_CHECKIN in raw.columns:
            raw['_checkin'] = raw[COL_CHECKIN].apply(_clean_checkin_time)
        else:
            raw['_checkin'] = pd.NaT
        raw = raw.sort_values('_checkin').drop_duplicates(subset='_id', keep='last')
        print(f"按工号去重后: {len(raw)} 行")

        return raw.reset_index(drop=True).copy()

    except ValueError:
        raise
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None


# ============ 统计 ============

def process_work_data(df, file_path, show_missing=True):
    if df.empty:
        return {}

    target_date = _resolve_target_date(file_path)
    date_str = target_date.strftime('%Y-%m-%d') if target_date else datetime.now().strftime('%Y-%m-%d')

    records = []
    is_v2 = COL_TOMORROW in df.columns
    for _, row in df.iterrows():
        start, end = row['_start'], row['_end']
        diff = end - start
        if diff.total_seconds() < 0:
            diff = diff + timedelta(days=1)
        hours = round(diff.total_seconds() / 3600, 2)

        if is_v2:
            # 新格式：分组按明日安排（接班组=明日16点接班，值班组=明日值班，与一区版按次日分类一致）
            # 次日列=明日工作地点，当日列=今日班次/房间
            today_type, today_room = _shift_info(row.get(COL_SHIFT_TODAY))
            tom_type, tom_room = _shift_info(row.get(COL_TOMORROW))
            if tom_type in ('接班', '值班'):
                next_loc = tom_type
            elif tom_type == '白班':
                next_loc = tom_room
            else:
                next_loc = ''
            day_loc = '接班' if today_type == '接班' else today_room
            is_duty = tom_type == '值班'
            is_handover = (not is_duty) and (tom_type == '接班')
            is_today_handover = today_type == '接班'
        else:
            # 旧格式：单一工作地点列=当日房间，无明日数据，接班按是否4p副麻接班（当日）
            day_loc = '' if pd.isna(row.get(COL_LOCATION)) else str(row[COL_LOCATION]).strip()
            next_loc = ''
            is_duty = '值班' in day_loc
            is_handover = (not is_duty) and ((str(row.get(COL_HANDOVER, '')).strip() == '是') or ('接班' in day_loc))
            is_today_handover = is_handover

        cross_day = end.date() > start.date()

        records.append({
            '姓名': str(row[COL_NAME]).strip(),
            '工号': row['_id'],
            '次日': next_loc,
            '当日': day_loc,
            '上班时间': start.strftime('%H:%M'),
            '下班时间': ('次日' if cross_day else '') + end.strftime('%H:%M'),
            '上班时长(小时)': hours,
            '_is_handover': is_handover,
            '_is_duty': is_duty,
            '_is_today_handover': is_today_handover,
            '_end_dt': end,
        })

    buckets = {cat: [] for cat in CATEGORY_ORDER}
    for r in records:
        if r['_is_duty']:
            buckets[config.CATEGORY_DUTY].append(r)
        elif r['_is_handover']:
            buckets[config.CATEGORY_HANDOVER].append(r)
        else:
            buckets[config.CATEGORY_NORMAL].append(r)

    buckets[config.CATEGORY_NORMAL].sort(key=lambda r: r['上班时长(小时)'])
    buckets[config.CATEGORY_HANDOVER].sort(key=lambda r: r['_end_dt'])
    buckets[config.CATEGORY_DUTY].sort(key=lambda r: r['_end_dt'])

    categories = [(cat, buckets[cat]) for cat in CATEGORY_ORDER]
    return {'date': date_str, 'categories': categories}


# ============ 渲染与截图 ============

def generate_html_table(data):
    template = Template(config.HTML_TEMPLATE)

    for _, data_list in data['categories']:
        for record in data_list:
            if record['_is_handover'] or record['_is_duty'] or record.get('_is_today_handover'):
                record['css_class'] = ''
            else:
                time_category = get_time_category(record['上班时长(小时)'])
                record['css_class'] = time_category if time_category != 'normal' else ''

    date_str = data['date']
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    title = f"{date_obj.year}-{date_obj.month:02d}-{date_obj.day:02d}(星期{config.WEEK_MAP[date_obj.weekday()]})"

    return template.render(title=title, categories=data['categories'])


def create_table_image_from_html(html_file_path, output_filename):
    if not PLAYWRIGHT_AVAILABLE:
        print("未安装 playwright，跳过截图")
        return None
    try:
        print("正在使用playwright转换HTML为图片...")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(f"file:///{os.path.abspath(html_file_path).replace(os.sep, '/')}")
            page.wait_for_load_state('networkidle')
            page.screenshot(path=output_filename, full_page=True)
            browser.close()
        print(f"图片已保存到: {output_filename}")
        return output_filename
    except Exception as e:
        print(f"playwright转换失败: {e}")
        return None


def main(file_path=None, show_missing=True):
    if file_path is None:
        print("未指定文件")
        return

    if not os.path.exists(file_path):
        print(f"文件 {file_path} 不存在")
        return

    print(f"正在读取文件：{file_path}")
    df = read_excel_file(file_path)
    if df is None or df.empty:
        print("没有找到有效的数据")
        return

    print("处理工作数据...")
    data = process_work_data(df, file_path, show_missing=show_missing)
    if not data:
        print("没有有效的工作数据")
        return

    print("生成HTML表格...")
    html_content = generate_html_table(data)
    html_filename = config.HTML_OUTPUT_FILENAME
    with open(html_filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"HTML文件已保存: {html_filename}")

    print("生成表格图片...")
    target_date = _resolve_target_date(file_path)
    date_str = target_date.strftime('%Y%m%d') if target_date else datetime.now().strftime('%Y%m%d')
    image_filename = create_table_image_from_html(html_filename, f"{date_str}{config.IMAGE_SUFFIX}")

    print("处理完成！")
    return html_filename, image_filename


# ============ 表格编辑（overview / save / patch） ============

def _overview_sheet(file_path):
    sheets = _date_sheets(file_path)
    return sheets[0] if sheets else None


def _fmt_dt_cell(value):
    if pd.isna(value):
        return ''
    ts = pd.to_datetime(value, errors='coerce')
    return ts.strftime('%Y-%m-%d %H:%M') if pd.notna(ts) else str(value).strip()


def _editable_cols(df):
    """按文件实际列返回可编辑列（新/旧表单格式自适应）"""
    if COL_TOMORROW in df.columns:
        ordered = [COL_NAME, COL_ID, config.COL_START_V2, config.COL_END_V2, COL_SHIFT_TODAY, COL_TOMORROW]
    else:
        ordered = [COL_NAME, COL_ID, COL_START, COL_END, COL_LOCATION, COL_HANDOVER]
    return [c for c in ordered if c in df.columns]


def read_excel_overview(file_path):
    """读取第一个日期 Sheet，供前端编辑弹窗展示"""
    sheet = _overview_sheet(file_path)
    if sheet is None:
        raise ValueError("未找到日期命名的 Sheet")
    df = pd.read_excel(file_path, sheet_name=sheet)

    existing_cols = _editable_cols(df)
    result = df[existing_cols].copy()
    for c in (config.COL_START_V2, config.COL_END_V2, COL_START, COL_END):
        if c in result.columns:
            result[c] = result[c].apply(_fmt_dt_cell)
    result = result.where(result.notna(), '')
    result['_sheet'] = sheet
    result['_orig_idx'] = range(len(result))
    return result, existing_cols


def _update_cells(file_path, sheet_name, updates):
    """updates: [(excel行号, 列名, 新值), ...]"""
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]
    for row_idx, col_name, value in updates:
        for c_idx in range(1, ws.max_column + 1):
            if ws.cell(row=1, column=c_idx).value == col_name:
                ws.cell(row=row_idx, column=c_idx, value=value if not pd.isna(value) else '')
                break
    wb.save(file_path)


def _append_row(file_path, sheet_name, row_data):
    wb = openpyxl.load_workbook(file_path)
    ws = wb[sheet_name]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    new_row = [row_data.get(h, '') if h in row_data else '' for h in headers]
    ws.append(new_row)
    wb.save(file_path)


def save_excel_overview(file_path, changes, overview_df, log=print):
    """按行号回写被编辑的单元格到原 Sheet"""
    sheet = overview_df['_sheet'].iloc[0] if len(overview_df) else _overview_sheet(file_path)
    updates = []
    for idx_str, fields in changes.items():
        idx = int(idx_str)
        orig_idx = int(overview_df.iloc[idx].get('_orig_idx', idx))
        excel_row = orig_idx + 2
        for col, val in fields.items():
            updates.append((excel_row, col, val))
    _update_cells(file_path, sheet, updates)
    log(f"已保存 {len(changes)} 条修改到: {file_path}")


def patch_excel_data(file_path, row_data, log=print):
    """新增/修改一行：工号已存在则更新，否则追加到第一个日期 Sheet（新/旧列名自适应）"""
    sheet = _overview_sheet(file_path)
    if sheet is None:
        raise ValueError("未找到日期命名的 Sheet")

    target_date = _target_date_from_filename(file_path)
    if target_date is None:
        target_date = datetime.strptime(sheet, '%Y-%m-%d').date()

    def _full_time(text, default_hm):
        text = str(text or '').strip()
        if not text:
            text = f"{target_date.strftime('%Y-%m-%d')} {default_hm}" if default_hm else ''
        elif re.match(r'^\d{1,2}:\d{2}$', text):
            text = f"{target_date.strftime('%Y-%m-%d')} {text}"
        return text

    df = pd.read_excel(file_path, sheet_name=sheet)
    cols = _editable_cols(df)
    # 上/下班列名（新表单是 24小时制 后缀版）
    start_col = config.COL_START_V2 if config.COL_START_V2 in cols else COL_START
    end_col = config.COL_END_V2 if config.COL_END_V2 in cols else COL_END

    name = str(row_data.get(COL_NAME, '')).strip()
    eid = str(row_data.get(COL_ID, '')).strip()
    start = _full_time(row_data.get(start_col, ''), '07:30')
    # 下班时间不设默认值：留空会在处理时报"下班时间缺失"，提示补填
    end = _full_time(row_data.get(end_col, ''), None)

    values = {COL_NAME: name, COL_ID: eid, start_col: start, end_col: end}
    for c in cols:
        if c in (COL_NAME, COL_ID, start_col, end_col):
            continue
        v = str(row_data.get(c, '') or '').strip()
        if v:
            values[c] = v

    if eid and COL_ID in df.columns:
        mask = df[COL_ID].apply(_norm_id) == eid
    else:
        mask = pd.Series([False] * len(df))

    if mask.any():
        idx = mask.idxmax()
        _update_cells(file_path, sheet, [(idx + 2, c, v) for c, v in values.items()])
        log(f"已修改第 {idx + 2} 行: {name}({eid}) 上班={start} 下班={end}")
    else:
        new_row = dict(values)
        if COL_CHECKIN in df.columns:
            new_row[COL_CHECKIN] = ''
        if '序号' in df.columns:
            new_row['序号'] = len(df) + 1
        _append_row(file_path, sheet, new_row)
        log(f"已新增行: {name}({eid}) 上班={start} 下班={end}")
    log(f"文件已保存: {file_path}")


if __name__ == "__main__":
    file_path = None
    show_missing = True
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    if len(sys.argv) > 2:
        show_missing = sys.argv[2].lower() in ['true', '1', 'yes']
    main(file_path, show_missing)
