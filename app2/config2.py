# 新区域打卡统计配置（衍生版，独立于一区版 config.py）
#
# 数据格式（每个日期 Sheet，Sheet 名为 YYYY-MM-DD）：
#   序号 / 昵称 / 打卡时间 / 姓名 / 工号 / 地理位置 / 上班时间 / 下班时间 /
#   工作地点（几号手术室） / 是否4p副麻接班 / 评级 / 点评数 / 打卡天数 / 打卡次数
#
# 上班时间、下班时间均为完整日期时间（如 2026-08-20 16:00），
# 时长 = 下班时间 - 上班时间，直接相减。

# ---- 数据列名 ----
COL_CHECKIN = '打卡时间'
COL_NAME = '姓名'
COL_ID = '工号'
COL_START = '上班时间'
COL_END = '下班时间'
COL_LOCATION = '工作地点（几号手术室）'
COL_HANDOVER = '是否4p副麻接班'

# 2026-08-26 起表单改版后的新列名（与旧列共存兼容，按文件实际列自动选择）
COL_START_V2 = '上班时间（24小时制）'
COL_END_V2 = '下班时间（24小时制）'
COL_SHIFT_TODAY = '今日是白班还是16点接班'    # 值如 白班：（请补充工作地点手术室）：12 / 16点接班
COL_TOMORROW = '明日工作地点手术室'          # 值如 白班（请补充工作地点手术室）：19 / 16点接班 / 值班

# ---- 分组 ----
# 三组：白班（按时长升序）、接班（是否4p副麻接班=是 或 工作地点含"接班"，不标色）、值班（工作地点含"值班"，放最后，不标色）
CATEGORY_NORMAL = '白班'
CATEGORY_HANDOVER = '接班'
CATEGORY_DUTY = '值班'
CATEGORY_ORDER = [CATEGORY_NORMAL, CATEGORY_HANDOVER, CATEGORY_DUTY]

# ---- 加班标色阈值（沿用原版）----
WORK_HOURS_THRESHOLDS = {
    'light-red': (13.5, 14.5),
    'medium-red': (14.5, 15.5),
    'dark-red': (15.5, float('inf'))
}

# ---- 输出 ----
HTML_OUTPUT_FILENAME = "work_statistics_v2.html"
IMAGE_SUFFIX = "下班统计.png"

WEEK_MAP = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}

UI_TITLE = '打卡数据整理助手'

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{{ title }}</title>
    <style>
        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            margin: 2px;
            background-color: #FEFEFE;
            font-size: 24px;
        }
        .container {
            max-width: 100%;
            margin: 0 auto;
            background-color: white;
            padding: 6px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        h1 {
            text-align: center;
            color: #2C3E50;
            margin-bottom: 15px;
            font-size: 28px;
            line-height: 1.2;
            text-shadow: 0 1px 2px rgba(0,0,0,0.1);
            font-weight: 600;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 6px 0;
            font-size: 24px;
            table-layout: fixed;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        }
        th, td {
            border: 1px solid #E9E9E9;
            padding: 16px 3px;
            text-align: center;
            word-wrap: break-word;
            overflow-wrap: break-word;
            line-height: 1.4;
        }
        th {
            background: linear-gradient(135deg, #EFFEF8 0%, #E8F8F0 100%);
            font-weight: bold;
            color: #2C3E50;
            font-size: 24px;
            text-shadow: 0 1px 1px rgba(255,255,255,0.8);
        }
        .category-row {
            background: linear-gradient(135deg, #5C8DC7 0%, #4A7BB7 100%);
            color: white;
            font-weight: bold;
            font-size: 26px;
            text-shadow: 0 1px 2px rgba(0,0,0,0.2);
        }
        .light-red {
            background-color: #FFE6E6;
        }
        .medium-red {
            background-color: #FFB3B3;
        }
        .dark-red {
            background-color: #FF8080;
        }
        .gray-column {
            color: #A6A6A6;
        }
        .zebra {
            background-color: #F8F9FA;
        }
        tr:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            transition: all 0.2s ease;
        }
        tr {
            transition: all 0.2s ease;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>{{ title }}</h1>
        <table>
            <tr>
                <th>姓名(工号)</th>
                <th>次日</th>
                <th>下班时间</th>
                <th class="gray-column">当日</th>
            </tr>
            {% for category, data_list in categories %}
                {% if data_list %}
                <tr class="category-row">
                    <td colspan="4">{{ category }}</td>
                </tr>
                {% for record in data_list %}
                <tr{% if record.get('css_class') %} class="{{ record.css_class }}"{% elif loop.index is even %} class="zebra"{% endif %}>
                    <td>{{ record.姓名 }}({{ record.工号 }})</td>
                    <td>{{ record.get('次日', '') }}</td>
                    <td>{{ record.get('下班时间', '') }}</td>
                    <td class="gray-column">{{ record.get('当日', '') }}</td>
                </tr>
                {% endfor %}
                {% endif %}
            {% endfor %}
        </table>
    </div>
</body>
</html>
"""
