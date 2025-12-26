"""
康恩贝内部行业信息简报 - 高级数据看板 v4.0
双主题设计：优雅深色 / 清新浅色
专业级配色方案
"""

import streamlit as st
import json
import re
from datetime import datetime, timedelta
from collections import Counter
import imaplib
import email
from email.header import decode_header
import email.utils

# 尝试导入图表库
try:
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

try:
    from wordcloud import WordCloud
    import matplotlib.pyplot as plt
    import io
    import base64
    HAS_WORDCLOUD = True
except ImportError:
    HAS_WORDCLOUD = False

try:
    from streamlit_autorefresh import st_autorefresh
    HAS_AUTOREFRESH = True
except ImportError:
    HAS_AUTOREFRESH = False

# ========================== 配置 ==========================
QQ_EMAIL = "2420778484@qq.com"
AUTH_CODE = "ulhzlajcvkpsebjh"
TARGET_SUBJECT = "康恩贝内部行业信息简报"
STORAGE_FILE = "email_data.json"
PAGE_SIZE = 10
FETCH_DAYS = 30

# ========================== 高级配色方案 ==========================
# 深色主题 - 科技感蓝紫色调
DARK_COLORS = {
    'bg_primary': '#0F0F1A',
    'bg_secondary': '#1A1A2E',
    'bg_tertiary': '#252542',
    'border': '#2D2D4A',
    'text_primary': '#EAEAFF',
    'text_secondary': '#9090B0',
    'accent_1': '#6C5CE7',
    'accent_2': '#00CEC9',
    'accent_3': '#FD79A8',
    'accent_4': '#FDCB6E',
    'accent_5': '#74B9FF',
    'success': '#00B894',
    'chart_colors': ['#6C5CE7', '#00CEC9', '#FD79A8', '#FDCB6E', '#74B9FF', '#A29BFE', '#55EFC4', '#FF7675']
}

# 浅色主题 - 清新莫兰迪色调
LIGHT_COLORS = {
    'bg_primary': '#F7F8FC',
    'bg_secondary': '#FFFFFF',
    'bg_tertiary': '#EEF0F7',
    'border': '#E0E4ED',
    'text_primary': '#2D3748',
    'text_secondary': '#718096',
    'accent_1': '#5B6AD0',
    'accent_2': '#38A89D',
    'accent_3': '#D4708A',
    'accent_4': '#D69E2E',
    'accent_5': '#4299E1',
    'success': '#38A169',
    'chart_colors': ['#5B6AD0', '#38A89D', '#D4708A', '#D69E2E', '#4299E1', '#9F7AEA', '#48BB78', '#ED8936']
}

def get_theme_css(dark_mode):
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    
    return f"""
<style>
    .stApp {{
        background: {c['bg_primary']};
    }}
    
    #MainMenu, footer, header {{visibility: hidden;}}
    .block-container {{padding: 1.5rem 2.5rem 2rem 2.5rem; max-width: 1500px;}}
    
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {c['bg_secondary']} 0%, {c['bg_primary']} 100%);
        border-right: 1px solid {c['border']};
        min-width: 300px !important;
        width: 300px !important;
        transform: translateX(0) !important;
    }}
    
    /* 强制侧边栏始终显示，禁止收起 */
    [data-testid="collapsedControl"] {{
        display: none !important;
    }}
    
    [data-testid="stSidebar"][aria-expanded="false"] {{
        min-width: 300px !important;
        width: 300px !important;
        margin-left: 0 !important;
        transform: translateX(0) !important;
    }}
    
    section[data-testid="stSidebar"] > div {{
        width: 300px !important;
    }}
    [data-testid="stSidebar"] .stMarkdown {{color: {c['text_primary']};}}
    [data-testid="stSidebar"] .stMarkdown h1,
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3 {{color: {c['text_primary']};}}
    
    .main-title {{
        font-size: 2.2rem;
        font-weight: 800;
        background: linear-gradient(135deg, {c['accent_1']} 0%, {c['accent_2']} 50%, {c['accent_5']} 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 0.3rem;
        letter-spacing: -0.5px;
    }}
    .sub-title {{
        color: {c['text_secondary']};
        font-size: 1rem;
        margin-bottom: 1.8rem;
        font-weight: 400;
    }}
    
    .status-bar {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 1rem 1.5rem;
        background: linear-gradient(135deg, {c['bg_secondary']}80, {c['bg_tertiary']}60);
        border: 1px solid {c['border']};
        border-radius: 14px;
        margin-bottom: 2rem;
        backdrop-filter: blur(10px);
    }}
    .status-left {{display: flex; align-items: center; gap: 0.8rem;}}
    .status-dot {{
        width: 10px; height: 10px;
        background: {c['success']};
        border-radius: 50%;
        animation: pulse 2s infinite;
        box-shadow: 0 0 12px {c['success']}80;
    }}
    @keyframes pulse {{
        0%, 100% {{opacity: 1; transform: scale(1);}}
        50% {{opacity: 0.6; transform: scale(0.85);}}
    }}
    .status-text {{color: {c['success']}; font-size: 0.9rem; font-weight: 500;}}
    .status-time {{color: {c['text_secondary']}; font-size: 0.85rem;}}
    
    .kpi-container {{
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1.5rem;
        margin-bottom: 2rem;
    }}
    .kpi-card {{
        background: linear-gradient(145deg, {c['bg_secondary']}, {c['bg_tertiary']}40);
        border: 1px solid {c['border']};
        border-radius: 18px;
        padding: 1.8rem 1.5rem;
        text-align: center;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative;
        overflow: hidden;
    }}
    .kpi-card::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 4px;
        background: linear-gradient(90deg, {c['accent_1']}, {c['accent_2']});
    }}
    .kpi-card:nth-child(2)::before {{background: linear-gradient(90deg, {c['accent_2']}, {c['accent_5']});}}
    .kpi-card:nth-child(3)::before {{background: linear-gradient(90deg, {c['accent_5']}, {c['accent_3']});}}
    .kpi-card:nth-child(4)::before {{background: linear-gradient(90deg, {c['accent_4']}, {c['accent_3']});}}
    .kpi-card:hover {{
        transform: translateY(-6px);
        box-shadow: 0 20px 40px {c['accent_1']}15;
        border-color: {c['accent_1']}50;
    }}
    .kpi-icon {{font-size: 2rem; margin-bottom: 0.6rem; filter: grayscale(0.2);}}
    .kpi-value {{
        font-size: 2.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, {c['accent_1']}, {c['accent_2']});
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        line-height: 1.2;
    }}
    .kpi-card:nth-child(2) .kpi-value {{background: linear-gradient(135deg, {c['accent_2']}, {c['accent_5']}); -webkit-background-clip: text;}}
    .kpi-card:nth-child(3) .kpi-value {{background: linear-gradient(135deg, {c['accent_5']}, {c['accent_3']}); -webkit-background-clip: text;}}
    .kpi-card:nth-child(4) .kpi-value {{background: linear-gradient(135deg, {c['accent_4']}, {c['accent_3']}); -webkit-background-clip: text;}}
    .kpi-label {{color: {c['text_secondary']}; font-size: 0.9rem; margin-top: 0.6rem; font-weight: 500;}}
    .kpi-change {{font-size: 0.8rem; margin-top: 0.5rem; font-weight: 600;}}
    .kpi-change.up {{color: {c['success']};}}
    
    .chart-container {{
        background: {c['bg_secondary']};
        border: 1px solid {c['border']};
        border-radius: 18px;
        padding: 1.8rem;
        margin-bottom: 1.5rem;
        transition: all 0.3s ease;
    }}
    .chart-container:hover {{
        box-shadow: 0 10px 30px {c['bg_primary']}40;
    }}
    .chart-title {{
        color: {c['text_primary']};
        font-size: 1.1rem;
        font-weight: 700;
        margin-bottom: 1.2rem;
        display: flex;
        align-items: center;
        gap: 0.8rem;
    }}
    .chart-title-icon {{
        width: 4px;
        height: 20px;
        background: linear-gradient(180deg, {c['accent_1']}, {c['accent_2']});
        border-radius: 2px;
    }}
    
    .list-item {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 1rem 0;
        border-bottom: 1px solid {c['border']}50;
    }}
    .list-item:last-child {{border-bottom: none;}}
    .list-item-label {{color: {c['text_primary']}; font-size: 0.95rem; font-weight: 500;}}
    .list-item-value {{
        color: {c['accent_1']};
        font-weight: 700;
        background: linear-gradient(135deg, {c['accent_1']}15, {c['accent_2']}15);
        padding: 0.4rem 1rem;
        border-radius: 25px;
        font-size: 0.9rem;
        border: 1px solid {c['accent_1']}30;
    }}
    
    .progress-bar {{
        height: 8px;
        background: {c['bg_tertiary']};
        border-radius: 4px;
        overflow: hidden;
        margin-top: 0.5rem;
    }}
    .progress-fill {{
        height: 100%;
        border-radius: 4px;
        transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1);
    }}
    
    .tag-list {{display: flex; flex-wrap: wrap; gap: 0.6rem; margin-top: 0.8rem;}}
    .tag-item {{
        padding: 0.35rem 0.85rem;
        border-radius: 25px;
        font-size: 0.75rem;
        font-weight: 600;
        letter-spacing: 0.3px;
        border: 1px solid;
    }}
    .tag-research {{background: {c['accent_1']}15; color: {c['accent_1']}; border-color: {c['accent_1']}30;}}
    .tag-policy {{background: {c['accent_2']}15; color: {c['accent_2']}; border-color: {c['accent_2']}30;}}
    .tag-market {{background: {c['accent_3']}15; color: {c['accent_3']}; border-color: {c['accent_3']}30;}}
    .tag-ai {{background: {c['accent_5']}15; color: {c['accent_5']}; border-color: {c['accent_5']}30;}}
    .tag-default {{background: {c['text_secondary']}15; color: {c['text_secondary']}; border-color: {c['text_secondary']}30;}}
    
    .content-card {{
        background: {c['bg_secondary']};
        border: 1px solid {c['border']};
        border-radius: 18px;
        padding: 1.5rem 1.8rem;
        margin-bottom: 1.2rem;
        border-left: 4px solid {c['accent_1']};
        transition: all 0.3s ease;
    }}
    .content-card:hover {{
        transform: translateX(6px);
        box-shadow: 0 10px 30px {c['accent_1']}10;
        border-left-color: {c['accent_2']};
    }}
    .card-header {{
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 1.2rem;
        flex-wrap: wrap;
    }}
    .card-title {{
        font-size: 1.05rem;
        font-weight: 600;
        color: {c['text_primary']};
        flex: 1;
        min-width: 200px;
        line-height: 1.6;
    }}
    .card-date {{
        font-size: 0.8rem;
        color: white;
        background: linear-gradient(135deg, {c['accent_1']}, {c['accent_2']});
        padding: 0.5rem 1.2rem;
        border-radius: 25px;
        font-weight: 600;
        white-space: nowrap;
    }}
    
    .empty-state {{
        text-align: center;
        padding: 5rem 2rem;
        color: {c['text_secondary']};
        background: {c['bg_secondary']};
        border-radius: 18px;
        border: 1px solid {c['border']};
    }}
    .empty-icon {{font-size: 4.5rem; margin-bottom: 1.5rem; opacity: 0.5;}}
    .empty-state h3 {{color: {c['text_primary']}; margin-bottom: 0.6rem; font-size: 1.3rem;}}
    .empty-state p {{color: {c['text_secondary']};}}
    
    .pagination-info {{
        text-align: center;
        color: {c['text_secondary']};
        font-size: 0.9rem;
        padding: 1.2rem;
        background: {c['bg_secondary']};
        border-radius: 14px;
        margin: 1.5rem 0;
        border: 1px solid {c['border']};
    }}
    
    .stSelectbox > div > div {{
        background: {c['bg_secondary']};
        border: 1px solid {c['border']};
        border-radius: 12px;
        color: {c['text_primary']};
    }}
    .stTextInput > div > div > input {{
        background: {c['bg_secondary']} !important;
        border: 1px solid {c['border']} !important;
        border-radius: 12px !important;
        color: {c['text_primary']} !important;
    }}
    .stTextInput > div > div > input:focus {{
        border-color: {c['accent_1']} !important;
        box-shadow: 0 0 0 3px {c['accent_1']}20 !important;
    }}
    .stButton > button {{
        background: linear-gradient(135deg, {c['accent_1']}, {c['accent_2']});
        border: none;
        color: white;
        border-radius: 12px;
        font-weight: 600;
        padding: 0.6rem 1.5rem;
        transition: all 0.3s ease;
    }}
    .stButton > button:hover {{
        transform: translateY(-2px);
        box-shadow: 0 8px 20px {c['accent_1']}40;
    }}
    div[data-testid="stExpander"] {{
        background: {c['bg_secondary']};
        border: 1px solid {c['border']};
        border-radius: 14px;
        overflow: hidden;
    }}
    div[data-testid="stExpander"] details {{
        background: {c['bg_secondary']};
    }}
    div[data-testid="stExpander"] summary {{
        color: {c['text_primary']} !important;
        font-weight: 600 !important;
        background: {c['bg_tertiary']}50;
    }}
    div[data-testid="stExpander"] summary span {{
        color: {c['text_primary']} !important;
    }}
    div[data-testid="stExpander"] div[data-testid="stMarkdownContainer"] {{
        color: {c['text_primary']} !important;
    }}
    div[data-testid="stExpander"] div[data-testid="stMarkdownContainer"] p {{
        color: {c['text_primary']} !important;
    }}
    .streamlit-expanderHeader {{
        color: {c['text_primary']} !important;
        font-weight: 600 !important;
        background: {c['bg_tertiary']}50;
    }}
    .streamlit-expanderContent {{
        color: {c['text_primary']} !important;
        padding: 1.2rem;
        background: {c['bg_tertiary']}30;
    }}
    .streamlit-expanderContent p {{
        color: {c['text_primary']} !important;
    }}
    
    /* 确保所有文本在深色模式下可见 */
    .stMarkdown, .stMarkdown p, .stText {{
        color: {c['text_primary']} !important;
    }}
</style>
"""

# ========================== 邮件处理函数 ==========================
def decode_chinese(s):
    if not s:
        return ""
    if isinstance(s, bytes):
        try:
            s = s.decode("utf-8")
        except UnicodeDecodeError:
            s = str(s)
    decoded = decode_header(s)
    result = []
    for part, encoding in decoded:
        if isinstance(part, bytes):
            for enc in [encoding, "utf-8", "gbk", "gb2312"]:
                if enc:
                    try:
                        result.append(part.decode(enc))
                        break
                    except UnicodeDecodeError:
                        continue
            else:
                result.append(str(part))
        else:
            result.append(str(part))
    return "".join(result)

def fetch_emails():
    today = datetime.now().date()
    start_date = today - timedelta(days=FETCH_DAYS)
    tomorrow = today + timedelta(days=1)

    try:
        mail = imaplib.IMAP4_SSL("imap.qq.com", 993)
        mail.login(QQ_EMAIL, AUTH_CODE)
    except Exception as e:
        return [], str(e)

    select_status, _ = mail.select("INBOX")
    if select_status != "OK":
        mail.logout()
        return [], "无法访问收件箱"

    start_date_str = start_date.strftime("%d-%b-%Y")
    tomorrow_str = tomorrow.strftime("%d-%b-%Y")
    status, data = mail.search(None, f"SINCE {start_date_str} BEFORE {tomorrow_str}")
    
    if status != "OK":
        mail.close()
        mail.logout()
        return [], "无法获取邮件列表"
    
    email_ids = data[0].split()
    existing_ids = set()
    try:
        with open(STORAGE_FILE, "r", encoding="utf-8") as f:
            stored_data = json.load(f)
            existing_ids = {item["email_id"] for item in stored_data}
    except:
        pass

    new_emails = []
    for email_id in reversed(email_ids):
        email_id_str = email_id.decode()
        if email_id_str in existing_ids:
            continue

        status, msg_data = mail.fetch(email_id, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])

        subject = decode_chinese(msg.get("Subject", ""))
        if TARGET_SUBJECT not in subject:
            continue

        content = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        content = decode_chinese(payload)
                    break
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                content = decode_chinese(payload)

        send_time = "未知"
        date_str = msg.get("Date")
        if date_str:
            try:
                send_time = email.utils.parsedate_to_datetime(date_str).strftime("%Y-%m-%d %H:%M:%S")
            except:
                pass

        new_emails.append({
            "email_id": email_id_str,
            "send_time": send_time,
            "subject": subject,
            "content": content.strip()
        })

    mail.close()
    mail.logout()
    return new_emails, None

def save_emails(new_emails):
    if not new_emails:
        return 0
    try:
        with open(STORAGE_FILE, "r", encoding="utf-8") as f:
            all_data = json.load(f)
    except:
        all_data = []

    all_data.extend(new_emails)
    unique = {item["email_id"]: item for item in all_data}
    all_data = sorted(unique.values(), key=lambda x: x.get("send_time", ""), reverse=True)
    
    with open(STORAGE_FILE, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    return len(new_emails)

def load_data():
    try:
        with open(STORAGE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []

# ========================== 数据分析函数 ==========================
def extract_date(subject):
    match = re.search(r"\d{4}-\d{2}-\d{2}", subject)
    if match:
        try:
            return datetime.strptime(match.group(), "%Y-%m-%d")
        except:
            pass
    return datetime(1970, 1, 1)

def extract_sources(subject):
    match = re.search(r"（来源：(.+?)）", subject)
    if match:
        return [s.strip() for s in re.split(r'[/、]', match.group(1))]
    return []

def extract_keywords(content):
    keywords = []
    patterns = [
        (r'研发', '研发', 'research'),
        (r'合规', '合规', 'policy'),
        (r'政策', '政策', 'policy'),
        (r'生产', '生产', 'default'),
        (r'销售|市场', '市场', 'market'),
        (r'创新药', '创新药', 'research'),
        (r'中药', '中药', 'default'),
        (r'AI|人工智能', 'AI', 'ai'),
        (r'临床', '临床', 'research'),
        (r'审评|审批', '审批', 'policy'),
    ]
    seen = set()
    for pattern, tag, tag_type in patterns:
        if re.search(pattern, content) and tag not in seen:
            keywords.append((tag, tag_type))
            seen.add(tag)
    return keywords[:6]

def extract_departments(content):
    depts = []
    patterns = [
        (r'研发部', '研发部'),
        (r'市场部', '市场部'),
        (r'合规部', '合规部'),
        (r'生产部', '生产部'),
        (r'销售部', '销售部'),
        (r'战略部', '战略部'),
        (r'投资部', '投资部'),
        (r'供应链', '供应链'),
        (r'财务部', '财务部'),
        (r'采购部', '采购部'),
    ]
    for pattern, dept in patterns:
        if re.search(pattern, content):
            depts.append(dept)
    return depts

def extract_hot_topics(content):
    """提取热点话题"""
    topics = []
    patterns = [
        (r'AI|人工智能', 'AI药物研发'),
        (r'创新药', '创新药研发'),
        (r'中药', '中药监管'),
        (r'医疗器械', '医疗器械'),
        (r'临床试验', '临床试验'),
        (r'合规|监管', '政策合规'),
        (r'合作|战略', '战略合作'),
        (r'数字化|数智化', '数字化转型'),
        (r'疫苗', '疫苗研发'),
        (r'肺癌|抗癌', '抗癌药物'),
        (r'化妆品', '化妆品监管'),
        (r'原料药', '原料药'),
        (r'质量', '质量管理'),
        (r'出口|进口', '药品进出口'),
    ]
    seen = set()
    for pattern, topic in patterns:
        if re.search(pattern, content) and topic not in seen:
            topics.append(topic)
            seen.add(topic)
    return topics

def extract_competitors(subject):
    """从标题提取提及的竞品"""
    competitors = []
    patterns = [
        (r'仁和', '仁和'),
        (r'同仁堂', '同仁堂'),
        (r'阿斯利康', '阿斯利康'),
        (r'石药集团', '石药集团'),
    ]
    for pattern, name in patterns:
        if re.search(pattern, subject):
            competitors.append(name)
    return competitors

def get_analytics(data):
    if not data:
        return {}
    
    total = len(data)
    today = datetime.now().date()
    
    # 本周一（weekday: 0=周一, 1=周二, ..., 6=周日）
    monday = today - timedelta(days=today.weekday())
    # 本月1号
    first_day_of_month = today.replace(day=1)
    
    this_week = sum(1 for item in data if extract_date(item["subject"]).date() >= monday)
    this_month = sum(1 for item in data if extract_date(item["subject"]).date() >= first_day_of_month)
    
    all_sources = []
    for item in data:
        all_sources.extend(extract_sources(item["subject"]))
    source_counter = Counter(all_sources)
    
    all_keywords = []
    for item in data:
        all_keywords.extend([k[0] for k in extract_keywords(item["content"])])
    keyword_counter = Counter(all_keywords)
    
    all_depts = []
    for item in data:
        all_depts.extend(extract_departments(item["content"]))
    dept_counter = Counter(all_depts)
    
    # 热点话题统计（近7期）
    recent_data = sorted(data, key=lambda x: extract_date(x["subject"]), reverse=True)[:7]
    all_topics = []
    for item in recent_data:
        all_topics.extend(extract_hot_topics(item["content"]))
    topic_counter = Counter(all_topics)
    
    # 竞品动态统计
    all_competitors = []
    for item in data:
        all_competitors.extend(extract_competitors(item["subject"]))
    competitor_counter = Counter(all_competitors)
    
    date_counter = Counter()
    for item in data:
        d = extract_date(item["subject"])
        if d.year > 1970:
            date_counter[d.strftime("%m-%d")] += 1
    
    dates = [extract_date(item["subject"]) for item in data]
    valid_dates = [d for d in dates if d.year > 1970]
    latest = max(valid_dates) if valid_dates else None
    
    return {
        "total": total,
        "this_week": this_week,
        "this_month": this_month,
        "sources": source_counter.most_common(10),
        "keywords": keyword_counter.most_common(10),
        "departments": dept_counter.most_common(10),
        "hot_topics": topic_counter.most_common(8),
        "competitors": competitor_counter.most_common(4),
        "daily_trend": sorted(date_counter.items()),
        "latest_date": latest.strftime("%Y-%m-%d") if latest else "无"
    }

# ========================== 图表函数 ==========================
def get_chart_colors(dark_mode):
    return DARK_COLORS['chart_colors'] if dark_mode else LIGHT_COLORS['chart_colors']

def create_source_pie(sources, dark_mode):
    if not sources or not HAS_PLOTLY:
        return None
    
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    labels = [s[0][:10] for s in sources[:6]]
    values = [s[1] for s in sources[:6]]
    colors = get_chart_colors(dark_mode)
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.6,
        marker_colors=colors[:len(labels)],
        textinfo='percent',
        textfont_size=12,
        textfont_color=c['text_primary'],
        hovertemplate="<b>%{label}</b><br>%{value} 次<br>%{percent}<extra></extra>"
    )])
    
    fig.update_layout(
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.25,
            xanchor="center",
            x=0.5,
            font=dict(size=11, color=c['text_secondary'])
        ),
        margin=dict(l=20, r=20, t=30, b=70),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=300
    )
    return fig

def create_keyword_bar(keywords, dark_mode):
    if not keywords or not HAS_PLOTLY:
        return None
    
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    colors = get_chart_colors(dark_mode)
    labels = [k[0] for k in keywords[:8]][::-1]
    values = [k[1] for k in keywords[:8]][::-1]
    bar_colors = [colors[i % len(colors)] for i in range(len(values))][::-1]
    
    fig = go.Figure(data=[go.Bar(
        x=values,
        y=labels,
        orientation='h',
        marker_color=bar_colors,
        marker_line_width=0,
        text=values,
        textposition='outside',
        textfont=dict(color=c['text_primary'], size=12),
        hovertemplate="<b>%{y}</b>: %{x} 次<extra></extra>"
    )])
    
    fig.update_layout(
        xaxis=dict(
            showgrid=True,
            gridcolor=c['border'],
            tickfont=dict(color=c['text_secondary'], size=11),
            zeroline=False
        ),
        yaxis=dict(
            tickfont=dict(color=c['text_primary'], size=12)
        ),
        margin=dict(l=90, r=60, t=20, b=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=300
    )
    return fig

def hex_to_rgba(hex_color, alpha=1.0):
    """将十六进制颜色转换为rgba格式"""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f'rgba({r},{g},{b},{alpha})'

def create_trend_chart(daily_trend, dark_mode):
    if not daily_trend or not HAS_PLOTLY:
        return None
    
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    dates = [d[0] for d in daily_trend[-14:]]
    values = [d[1] for d in daily_trend[-14:]]
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=dates,
        y=values,
        fill='tozeroy',
        fillcolor=hex_to_rgba(c['accent_1'], 0.15),
        line=dict(color=c['accent_1'], width=3),
        mode='lines+markers',
        marker=dict(size=10, color=c['accent_1'], line=dict(width=3, color=c['bg_secondary'])),
        hovertemplate="<b>%{x}</b><br>%{y} 条简报<extra></extra>"
    ))
    
    fig.update_layout(
        xaxis=dict(
            showgrid=False,
            tickfont=dict(color=c['text_secondary'], size=11),
            tickangle=-45
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor=c['border'],
            tickfont=dict(color=c['text_secondary'], size=11),
            zeroline=False
        ),
        margin=dict(l=50, r=30, t=20, b=70),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=220,
        hovermode='x unified'
    )
    return fig

def create_dept_pie(departments, dark_mode):
    if not departments or not HAS_PLOTLY:
        return None
    
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    colors = get_chart_colors(dark_mode)
    labels = [d[0] for d in departments[:6]]
    values = [d[1] for d in departments[:6]]
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        marker_colors=colors[:len(labels)],
        textinfo='label+percent',
        textfont_size=11,
        textfont_color=c['text_primary'],
        hovertemplate="<b>%{label}</b><br>%{value} 次<br>%{percent}<extra></extra>"
    )])
    
    fig.update_layout(
        showlegend=False,
        margin=dict(l=20, r=20, t=20, b=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=250
    )
    return fig

def create_hot_topics_chart(topics, dark_mode):
    """创建近期热点话题图表"""
    if not topics or not HAS_PLOTLY:
        return None
    
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    colors = get_chart_colors(dark_mode)
    
    labels = [t[0] for t in topics[:6]][::-1]
    values = [t[1] for t in topics[:6]][::-1]
    bar_colors = [colors[i % len(colors)] for i in range(len(values))][::-1]
    
    fig = go.Figure(data=[go.Bar(
        x=values,
        y=labels,
        orientation='h',
        marker_color=bar_colors,
        marker_line_width=0,
        text=values,
        textposition='outside',
        textfont=dict(color=c['text_primary'], size=12),
        hovertemplate="<b>%{y}</b>: 近7期出现 %{x} 次<extra></extra>"
    )])
    
    fig.update_layout(
        xaxis=dict(
            showgrid=True,
            gridcolor=c['border'],
            tickfont=dict(color=c['text_secondary'], size=11),
            zeroline=False,
            dtick=1
        ),
        yaxis=dict(
            tickfont=dict(color=c['text_primary'], size=11)
        ),
        margin=dict(l=100, r=50, t=20, b=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=220
    )
    return fig

def create_dept_radar(departments, dark_mode):
    """创建部门任务雷达图"""
    if not departments or not HAS_PLOTLY or len(departments) < 3:
        return None
    
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    
    labels = [d[0] for d in departments[:8]]
    values = [d[1] for d in departments[:8]]
    
    # 雷达图需要闭合，复制第一个值到最后
    labels_closed = labels + [labels[0]]
    values_closed = values + [values[0]]
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatterpolar(
        r=values_closed,
        theta=labels_closed,
        fill='toself',
        fillcolor=hex_to_rgba(c['accent_1'], 0.2),
        line=dict(color=c['accent_1'], width=2),
        marker=dict(size=8, color=c['accent_1']),
        hovertemplate="<b>%{theta}</b><br>%{r} 次行动建议<extra></extra>"
    ))
    
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                showticklabels=True,
                tickfont=dict(color=c['text_secondary'], size=10),
                gridcolor=c['border'],
                linecolor=c['border']
            ),
            angularaxis=dict(
                tickfont=dict(color=c['text_primary'], size=11),
                gridcolor=c['border'],
                linecolor=c['border']
            ),
            bgcolor='rgba(0,0,0,0)'
        ),
        margin=dict(l=60, r=60, t=30, b=30),
        paper_bgcolor='rgba(0,0,0,0)',
        height=280,
        showlegend=False
    )
    return fig

def create_competitor_bar(competitors, dark_mode):
    """创建竞品动态统计图"""
    if not competitors or not HAS_PLOTLY:
        return None
    
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    colors = [c['accent_3'], c['accent_2'], c['accent_5'], c['accent_4']]
    
    labels = [comp[0] for comp in competitors]
    values = [comp[1] for comp in competitors]
    
    fig = go.Figure(data=[go.Bar(
        x=labels,
        y=values,
        marker_color=colors[:len(labels)],
        text=values,
        textposition='outside',
        textfont=dict(color=c['text_primary'], size=13, weight='bold'),
        hovertemplate="<b>%{x}</b><br>被提及 %{y} 次<extra></extra>"
    )])
    
    fig.update_layout(
        xaxis=dict(
            tickfont=dict(color=c['text_primary'], size=12),
            showgrid=False
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor=c['border'],
            tickfont=dict(color=c['text_secondary'], size=11),
            zeroline=False
        ),
        margin=dict(l=40, r=20, t=20, b=40),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=200
    )
    return fig

def create_wordcloud(data, dark_mode):
    """生成词云图"""
    if not HAS_WORDCLOUD or not data:
        return None
    
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    
    # 收集所有关键词和话题
    all_words = []
    for item in data:
        content = item.get("content", "")
        # 提取关键词
        keywords = extract_keywords(content)
        for kw, _ in keywords:
            all_words.extend([kw] * 3)  # 关键词权重高
        # 提取话题
        topics = extract_hot_topics(content)
        all_words.extend(topics)
        # 提取部门
        depts = extract_departments(content)
        all_words.extend(depts)
        # 提取来源
        sources = extract_sources(item.get("subject", ""))
        for s in sources:
            if len(s) <= 6:  # 只保留短名称
                all_words.append(s)
    
    if not all_words:
        return None
    
    # 统计词频
    word_freq = Counter(all_words)
    
    # 词云配色
    if dark_mode:
        bg_color = '#1A1A2E'
        color_list = ['#6C5CE7', '#00CEC9', '#FD79A8', '#FDCB6E', '#74B9FF', '#A29BFE', '#55EFC4', '#FF7675', '#81ECEC', '#FFEAA7']
    else:
        bg_color = '#FFFFFF'
        color_list = ['#5B6AD0', '#38A89D', '#D4708A', '#D69E2E', '#4299E1', '#9F7AEA', '#48BB78', '#ED8936', '#38B2AC', '#ECC94B']
    
    def color_func(word, font_size, position, orientation, random_state=None, **kwargs):
        import random
        return random.choice(color_list)
    
    # 生成词云
    wc = WordCloud(
        font_path=None,  # 使用默认字体，中文可能需要指定
        width=800,
        height=400,
        background_color=bg_color,
        max_words=50,
        max_font_size=120,
        min_font_size=16,
        color_func=color_func,
        prefer_horizontal=0.8,
        relative_scaling=0.5,
        margin=10
    )
    
    # 尝试使用中文字体
    try:
        # Windows 常见中文字体路径
        font_paths = [
            "C:/Windows/Fonts/msyh.ttc",  # 微软雅黑
            "C:/Windows/Fonts/simhei.ttf",  # 黑体
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",  # Linux
            "/System/Library/Fonts/PingFang.ttc",  # Mac
        ]
        for fp in font_paths:
            import os
            if os.path.exists(fp):
                wc = WordCloud(
                    font_path=fp,
                    width=800,
                    height=400,
                    background_color=bg_color,
                    max_words=50,
                    max_font_size=120,
                    min_font_size=16,
                    color_func=color_func,
                    prefer_horizontal=0.8,
                    relative_scaling=0.5,
                    margin=10
                )
                break
    except:
        pass
    
    wc.generate_from_frequencies(word_freq)
    
    # 转换为base64图片
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.imshow(wc, interpolation='bilinear')
    ax.axis('off')
    fig.patch.set_facecolor(bg_color)
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', facecolor=bg_color, edgecolor='none', dpi=150)
    buf.seek(0)
    plt.close(fig)
    
    return buf

# ========================== 页面组件 ==========================
def render_kpi_cards(analytics, dark_mode):
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    total = analytics.get("total", 0)
    this_week = analytics.get("this_week", 0)
    this_month = analytics.get("this_month", 0)
    latest = analytics.get("latest_date", "无")
    
    week_change = f"+{this_week}" if this_week > 0 else "0"
    
    st.markdown(f'''
        <div class="kpi-container">
            <div class="kpi-card">
                <div class="kpi-icon">📊</div>
                <div class="kpi-value">{total}</div>
                <div class="kpi-label">简报总数</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon">📅</div>
                <div class="kpi-value">{this_week}</div>
                <div class="kpi-label">本周新增</div>
                <div class="kpi-change up">↑ {week_change}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon">📈</div>
                <div class="kpi-value">{this_month}</div>
                <div class="kpi-label">本月累计</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon">🕐</div>
                <div class="kpi-value" style="font-size:1.4rem;">{latest}</div>
                <div class="kpi-label">最新日期</div>
            </div>
        </div>
    ''', unsafe_allow_html=True)

# ========================== 主函数 ==========================
def main():
    st.set_page_config(
        page_title="康恩贝行业简报看板",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    if "dark_mode" not in st.session_state:
        st.session_state.dark_mode = True
    if "current_page" not in st.session_state:
        st.session_state.current_page = 1
    if "last_refresh" not in st.session_state:
        st.session_state.last_refresh = datetime.now()
    if "view_mode" not in st.session_state:
        st.session_state.view_mode = "看板"
    
    dark_mode = st.session_state.dark_mode
    c = DARK_COLORS if dark_mode else LIGHT_COLORS
    
    if HAS_AUTOREFRESH:
        st_autorefresh(interval=30*60*1000, key="auto_refresh")
    
    st.markdown(get_theme_css(dark_mode), unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### ⚙️ 控制面板")
        st.markdown("---")
        
        theme_icon = "🌙" if dark_mode else "☀️"
        theme_text = "深色模式" if dark_mode else "浅色模式"
        if st.button(f"{theme_icon} {theme_text}", use_container_width=True, type="primary"):
            st.session_state.dark_mode = not dark_mode
            st.rerun()
        
        st.markdown("")
        
        view_mode = st.radio(
            "视图模式",
            ["📊 数据看板", "📋 列表视图"],
            index=0 if st.session_state.view_mode == "看板" else 1,
            label_visibility="collapsed"
        )
        st.session_state.view_mode = "看板" if "看板" in view_mode else "列表"
        
        st.markdown("---")
        
        if st.button("🔄 同步数据", use_container_width=True):
            with st.spinner("同步中..."):
                new_mails, error = fetch_emails()
                if error:
                    st.error(f"❌ {error}")
                else:
                    count = save_emails(new_mails)
                    st.session_state.last_refresh = datetime.now()
                    st.success(f"✅ 新增 {count} 条")
                    st.rerun()
        
        st.caption(f"上次同步: {st.session_state.last_refresh.strftime('%H:%M:%S')}")
        
        st.markdown("---")
        
        st.markdown("#### 🔍 筛选")
        search_keyword = st.text_input("关键词", placeholder="搜索...", label_visibility="collapsed")
        
        data = load_data()
        all_dates = sorted(list({
            extract_date(item["subject"]).strftime("%Y-%m-%d")
            for item in data
            if extract_date(item["subject"]).year > 1970
        }), reverse=True)
        
        selected_date = st.selectbox("日期筛选", ["全部"] + all_dates[:30], label_visibility="collapsed")
        
        st.markdown("---")
        
        if st.button("📥 导出数据", use_container_width=True):
            if data:
                st.download_button(
                    "下载JSON",
                    json.dumps(data, ensure_ascii=False, indent=2),
                    f"简报_{datetime.now().strftime('%Y%m%d')}.json",
                    "application/json"
                )
        
        st.markdown("---")
        st.caption("v4.1 词云增强版")
        plotly_status = "✅" if HAS_PLOTLY else "❌"
        auto_status = "✅" if HAS_AUTOREFRESH else "❌"
        wc_status = "✅" if HAS_WORDCLOUD else "❌"
        st.caption(f"Plotly: {plotly_status} | 词云: {wc_status}")
    
    st.markdown('<div class="main-title">📊 康恩贝内部行业信息简报</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">实时行业动态监控 · 智能数据分析看板</div>', unsafe_allow_html=True)
    
    refresh_text = "自动同步已启用" if HAS_AUTOREFRESH else "手动同步模式"
    last_time = st.session_state.last_refresh.strftime("%Y-%m-%d %H:%M:%S")
    st.markdown(f'''
        <div class="status-bar">
            <div class="status-left">
                <div class="status-dot"></div>
                <span class="status-text">{refresh_text} · 数据已加载</span>
            </div>
            <span class="status-time">最后更新: {last_time}</span>
        </div>
    ''', unsafe_allow_html=True)
    
    data = load_data()
    
    if search_keyword:
        data = [item for item in data if search_keyword.lower() in item["subject"].lower() or search_keyword.lower() in item["content"].lower()]
    if selected_date != "全部":
        data = [item for item in data if extract_date(item["subject"]).strftime("%Y-%m-%d") == selected_date]
    
    if not data:
        st.markdown('''
            <div class="empty-state">
                <div class="empty-icon">📭</div>
                <h3>暂无数据</h3>
                <p>请点击侧边栏的"同步数据"按钮获取简报</p>
            </div>
        ''', unsafe_allow_html=True)
        return
    
    data.sort(key=lambda x: extract_date(x["subject"]), reverse=True)
    analytics = get_analytics(data)
    
    if st.session_state.view_mode == "看板":
        render_kpi_cards(analytics, dark_mode)
        
        # 词云图 - 展示所有关键词的视觉分布
        st.markdown(f'''
            <div class="chart-container">
                <div class="chart-title">
                    <div class="chart-title-icon"></div>
                    ☁️ 行业热点词云
                </div>
            </div>
        ''', unsafe_allow_html=True)
        
        if HAS_WORDCLOUD and data:
            wordcloud_img = create_wordcloud(data, dark_mode)
            if wordcloud_img:
                st.image(wordcloud_img, use_container_width=True)
            else:
                st.info("词云生成中...")
        else:
            st.info("安装 wordcloud 以查看词云: pip install wordcloud matplotlib")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f'''
                <div class="chart-container">
                    <div class="chart-title">
                        <div class="chart-title-icon"></div>
                        🏢 信息来源分布
                    </div>
                </div>
            ''', unsafe_allow_html=True)
            
            if HAS_PLOTLY and analytics.get("sources"):
                fig = create_source_pie(analytics["sources"], dark_mode)
                if fig:
                    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
            
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            colors = get_chart_colors(dark_mode)
            for i, (source, count) in enumerate(analytics.get("sources", [])[:5]):
                pct = int(count / analytics["total"] * 100) if analytics["total"] > 0 else 0
                color = colors[i % len(colors)]
                st.markdown(f'''
                    <div class="list-item">
                        <span class="list-item-label">{source[:12]}</span>
                        <span class="list-item-value">{count} 次</span>
                    </div>
                    <div class="progress-bar">
                        <div class="progress-fill" style="width:{pct}%;background:linear-gradient(90deg, {color}, {color}90);"></div>
                    </div>
                ''', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown(f'''
                <div class="chart-container">
                    <div class="chart-title">
                        <div class="chart-title-icon"></div>
                        🏷️ 热点关键词
                    </div>
                </div>
            ''', unsafe_allow_html=True)
            
            if HAS_PLOTLY and analytics.get("keywords"):
                fig = create_keyword_bar(analytics["keywords"], dark_mode)
                if fig:
                    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
            
            st.markdown(f'''
                <div class="chart-container">
                    <div class="chart-title">
                        <div class="chart-title-icon"></div>
                        🎯 部门行动建议分布
                    </div>
                </div>
            ''', unsafe_allow_html=True)
            
            if HAS_PLOTLY and analytics.get("departments"):
                fig = create_dept_radar(analytics["departments"], dark_mode)
                if fig:
                    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
                else:
                    # 如果部门数据太少，用饼图代替
                    fig = create_dept_pie(analytics["departments"], dark_mode)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        
        st.markdown("---")
        st.markdown(f'''
            <div class="chart-container">
                <div class="chart-title">
                    <div class="chart-title-icon"></div>
                    📋 最新简报
                </div>
            </div>
        ''', unsafe_allow_html=True)
        
        for item in data[:5]:
            date = extract_date(item["subject"])
            date_str = date.strftime("%Y-%m-%d") if date.year > 1970 else "未知"
            title = item["subject"].split("（来源")[0] if "（来源" in item["subject"] else item["subject"]
            keywords = extract_keywords(item["content"])
            
            tags_html = " ".join([f'<span class="tag-item tag-{k[1]}">{k[0]}</span>' for k in keywords[:4]])
            
            with st.expander(f"📄 {title} ({date_str})"):
                st.markdown(f'<div class="tag-list">{tags_html}</div>', unsafe_allow_html=True)
                st.markdown("---")
                st.markdown(item["content"])
    
    else:
        total = len(data)
        total_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        current_page = min(st.session_state.current_page, total_pages)
        
        start = (current_page - 1) * PAGE_SIZE
        end = min(start + PAGE_SIZE, total)
        
        st.markdown(f"共 **{total}** 条记录 | 第 {current_page}/{total_pages} 页")
        st.markdown("---")
        
        for idx, item in enumerate(data[start:end], start + 1):
            date = extract_date(item["subject"])
            date_str = date.strftime("%Y-%m-%d") if date.year > 1970 else "未知"
            title = item["subject"].split("（来源")[0] if "（来源" in item["subject"] else item["subject"]
            keywords = extract_keywords(item["content"])
            
            tags_html = " ".join([f'<span class="tag-item tag-{k[1]}">{k[0]}</span>' for k in keywords[:4]])
            
            st.markdown(f'''
                <div class="content-card">
                    <div class="card-header">
                        <div class="card-title">【{idx}】{title}</div>
                        <div class="card-date">📅 {date_str}</div>
                    </div>
                    <div class="tag-list">{tags_html}</div>
                </div>
            ''', unsafe_allow_html=True)
            
            with st.expander("查看详情"):
                st.markdown(item["content"])
        
        st.markdown(f'<div class="pagination-info">📄 显示 {start + 1}-{end} 条 / 共 {total} 条</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col1:
            if st.button("⬅️ 上一页", disabled=current_page == 1, use_container_width=True):
                st.session_state.current_page -= 1
                st.rerun()
        with col2:
            new_page = st.number_input("跳转", min_value=1, max_value=total_pages, value=current_page, label_visibility="collapsed")
            if new_page != current_page:
                st.session_state.current_page = new_page
                st.rerun()
        with col3:
            if st.button("下一页 ➡️", disabled=current_page == total_pages, use_container_width=True):
                st.session_state.current_page += 1
                st.rerun()

if __name__ == "__main__":
    main()
