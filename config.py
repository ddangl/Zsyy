

EXCEL_COLUMN_RENAME = {

    '房间号/工作点': '当日',

    '第二天房间号/工作点（周五打卡填周一工作点）': '次日'

}



TIME_OFF_WORK_COL = '下班时间 24小时制 英文半角:！'



DEFAULT_START_TIME = '07:30'
HANDOVER_START_TIME = '16:00'



WORK_HOURS_THRESHOLDS = {
    'light-red': (13.5, 14.5),
    'medium-red': (14.5, 15.5),
    'dark-red': (15.5, float('inf'))
}



ROOM_RANGES = {

    "21号楼": [50, 51, 52, 53, 55, 56],

    "肝科": list(range(21, 30)),

    "10号楼三楼": [74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84],

    "10号楼四楼": [61, 62, 63, 64, 65, 66, 68, 69, 70, 71, 72],

    "8号楼四楼": list(range(802, 826)),

    "8号楼三楼": list(range(826, 846))

}



OUTSIDE_KEYWORDS = [

    '气管', '镜', 'VIP', 'vip', '门诊', '胃', '肠', '体检', '佘山', '专项'

]



CATEGORY_ORDER = [

    "肝科", "21号楼", "10号楼三楼", "10号楼四楼",

    "8号楼四楼", "8号楼三楼", "外围", "其他",

    "接班", "值班", "休息"

]



HTML_OUTPUT_FILENAME = "work_statistics_v2.html"

IMAGE_SUFFIX = "下班统计.png"



WEEK_MAP = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}



UI_TITLE = '打卡数据整理助手'

UI_WINDOW_SIZE = '900x700'

UI_BG_COLOR = '#f0f0f0'

UI_COPYRIGHT = 'V2.1'



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

        .missing-row {

            background: linear-gradient(135deg, #F4E2EE 0%, #F0D8E8 100%);

            font-weight: bold;

            font-size: 24px;

            color: #6B4C7A;

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

                    <td>{{ record.get('下班时间 24小时制 英文半角:！', '') }}</td>

                    <td class="gray-column">{{ record.get('当日', '') }}</td>

                </tr>

                {% endfor %}

                {% if missing_rooms %}

                    {% if category in missing_rooms %}

                    <tr class="missing-row">

                        <td colspan="4">未打卡: {{ missing_rooms[category] }}</td>

                    </tr>

                    {% elif category in ["21号楼", "肝科", "10号楼三楼", "10号楼四楼", "8号楼四楼", "8号楼三楼"] %}

                    <tr class="missing-row">

                        <td colspan="4">均已打卡</td>

                    </tr>

                    {% endif %}

                {% endif %}

                {% endif %}

            {% endfor %}

        </table>

    </div>

</body>

</html>

"""

