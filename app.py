# -*- coding: utf-8 -*-
"""
Document Automation System (DAS) - Flask Backend V3.0
Nhập liệu Tiếng Việt → Hệ thống quản lý và xuất hồ sơ thông minh.
"""
import os, uuid, re, unicodedata, json, base64, traceback, io, zipfile, requests
from datetime import date, datetime, timedelta, timezone
from flask import Flask, request, jsonify, send_file, render_template, Response, make_response
from flask_cors import CORS
from jinja2 import Template
from deep_translator import GoogleTranslator
from dotenv import load_dotenv
from urllib.parse import quote
from unicodedata import normalize
from flask_sqlalchemy import SQLAlchemy
from flask_basicauth import BasicAuth
from PIL import Image
from vietnamese_names_dict import get_vietnamese_name_in_chinese
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

load_dotenv()

app = Flask(__name__, static_folder='static', static_url_path='')
app.debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
CORS(app, resources={r"/*": {"origins": ["https://cv.fct.vn", "http://127.0.0.1:5000", "http://localhost:5000"]}})
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB limit

app.config['BASIC_AUTH_USERNAME'] = 'fctvt'
app.config['BASIC_AUTH_PASSWORD'] = '1503'
app.config['BASIC_AUTH_FORCE_PROMPT'] = True
basic_auth = BasicAuth(app)

def auth_required(f):
    if os.environ.get('RENDER'):
        return basic_auth.required(f)
    return f

# --- DATABASE CONFIG ---
db_url = os.environ.get('DATABASE_URL')
if db_url and db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)
if not db_url:
    db_url = 'sqlite:///' + os.path.join(os.path.dirname(os.path.abspath(__file__)), 'history.db')

app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

class FormHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ma_so = db.Column(db.String(50))
    ho_ten = db.Column(db.String(100))
    ten_file = db.Column(db.String(255))
    data_json = db.Column(db.Text)
    ngay_tao = db.Column(db.DateTime, default=datetime.utcnow)
    is_selected = db.Column(db.Boolean, default=False)  # Trúng tuyển
    is_deleted = db.Column(db.Boolean, default=False)   # Xóa mềm

with app.app_context():
    try:
        db.create_all()
        # Auto migration for is_deleted
        try:
            from sqlalchemy import text, inspect
            inspector = inspect(db.engine)
            # Lấy danh sách cột của bảng form_history (có thể viết thường hoặc viết hoa)
            columns = [c['name'].lower() for c in inspector.get_columns('form_history')]
            if 'is_deleted' not in columns:
                if 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']:
                    db.session.execute(text('ALTER TABLE form_history ADD COLUMN is_deleted BOOLEAN DEFAULT 0'))
                else:
                    db.session.execute(text('ALTER TABLE form_history ADD COLUMN is_deleted BOOLEAN DEFAULT FALSE'))
                db.session.commit()
        except Exception as ex:
            print(f"⚠️ Column migration skipped or failed: {ex}")
            db.session.rollback()
    except Exception as e:
        print(f"❌ Database initialization error: {e}")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- TRANSLATION MAPS ---
FIXED_TRANS = {
    'độc thân': '未婚', 'doc than': '未婚', 'đã kết hôn': '已婚', 'da ket hon': '已婚', 'có gia đình': '已婚',
    'ly hôn': '離婚', 'ly hon': '離婚', 'góa': '喪偶', 'goa': '喪偶',
    'tiểu học': '國小', 'tieu hoc': '國小', 'thcs': '國中', 'trung học cơ sở': '國中',
    'thpt': '高中', 'trung học phổ thông': '高中', 'trung cấp': '高職', 'trung cap': '高職',
    'cao đẳng': '專科', 'cao dang': '專科', 'đại học': '大學', 'dai hoc': '大學',
    'thạc sĩ': '碩士', 'thac si': '碩士', 'tiến sĩ': '博士', 'tien si': '博士',
    'việt nam': '越南', 'viet nam': '越南', 'đài loan': '台灣', 'nhật bản': '日本',
    'hàn quốc': '韓國', 'malaysia': '馬來西亞', 'macau': '澳門', 'thái lan': '泰國',
    'châu âu': '歐洲', 'nga': '俄羅斯',
    'nghệ an': '藝安', 'nghe an': '藝安',
    'may': '縫紉', 'thợ may': '縫紉', 'may mặc': '縫紉', 'hàn': '焊接', 'thợ hàn': '焊接',
    'điện': '電工', 'thợ điện': '電工', 'sơn': '噴漆', 'thợ sơn': '噴漆',
    'tiện': '車床', 'thợ tiện': '車床', 'phay': '銑床', 'thợ phay': '銑床',
    'bào': '刨床', 'thợ bào': '刨床', 'đúc': '鑄造', 'thợ đúc': '鑄造',
    'dệt': '紡織', 'thợ dệt': '紡織', 'mộc': '木工', 'thợ mộc': '木工',
    'in ấn': '印刷', 'in': '印刷', 'cơ khí': '機械', 'gia công': '加工',
    'xây dựng': '營造', 'xây': '營造', 'lắp ráp': '組裝', 'đóng gói': '包裝',
    'kiểm tra': '檢查', 'qc': '品檢', 'kho': '倉庫', 'thủ kho': '倉管',
    'nấu ăn': '烹飪', 'đầu bếp': '廚師', 'giúp việc': '幫傭', 'điều dưỡng': '護理',
    'chăm sóc người già': '照顧老人', 'nông nghiệp': '農業', 'nông': '農業',
    'chăn nuôi': '畜牧', 'thủy sản': '水產', 'nhựa': '塑膠',
    'lái xe': '駕駛', 'tài xế': '司機', 'xe nâng': '堆高機', 'lái xe nâng': '堆高機',
    'cnc': 'CNC', 'tig mig': 'Tig/Mig',
}


YELLOW_ALERTS_MAP = {
    "f05": "骨折", "f06": "手汗", "f11": "脊椎受傷", "f13": "伏地挺身 10~30 下",
    "f14": "搬重 20~40 kg", "f18": "肝炎", "f19": "斷指", "f20": "哮喘",
    "f21": "伏地挺身 50 下以上", "f22": "搬重 50kg 以上",
}

SKILL_MAPPING = {
    'f23': 'Hàn điện / 電焊', 'f24': 'Hàn Argon / 氬焊', 'f25': 'Hàn CO2 / 氣焊', 'f26': 'Tig/Mig',
    'f27': 'Đúc / 鑄造', 'f28': 'Dệt / 紡織', 'f29': 'May / 縫紉', 'f30': 'Lái xe nâng / 堆高機',
    'f31': 'Tiện / 車床', 'f32': 'Phay / 銑床', 'f33': 'Bào / 刨床', 'f34': 'CNC',
    'f35': 'Đột dập / 沖床', 'f36': 'In ấn / 印刷', 'f37': 'Thợ mộc / 木工', 'f38': 'Lái xe tải/khách / 卡車/客司機',
    'f39': 'Nhựa / 塑膠', 'f40': 'Xây dựng / 營造', 'f41': 'Sửa chữa máy / 機械維修', 'f42': 'Điều dưỡng / 護理工',
    'f43': 'Giúp việc / 幫傭', 'f44': 'Xe cẩu / 吊車', 'f45': 'Cẩu trục / 天車', 'f46': 'Máy xúc / 挖土機'
}

ZH_TO_VI = {
    # Học vấn
    "國中": "Cấp 2",
    "高中": "Cấp 3",
    "中專": "Trung cấp",
    "大專": "Cao đẳng",
    "大學": "Đại học",
    "國小": "Cấp 1",
    
    # Nơi ở / 63 Tỉnh Thành
    "萊州": "Lai Châu",
    "奠邊": "Điện Biên",
    "山羅": "Sơn La",
    "和平": "Hòa Bình",
    "安沛": "Yên Bái",
    "老街": "Lào Cai",
    "河江": "Hà Giang",
    "高平": "Cao Bằng",
    "北𣴓": "Bắc Kạn",
    "諒山": "Lạng Sơn",
    "宣光": "Tuyên Quang",
    "太原": "Thái Nguyên",
    "富壽": "Phú Thọ",
    "北江": "Bắc Giang",
    "廣寧": "Quảng Ninh",
    "北寧": "Bắc Ninh",
    "河南": "Hà Nam",
    "河內": "Hà Nội",
    "海陽": "Hải Dương",
    "海防": "Hải Phòng",
    "興安": "Hưng Yên",
    "南定": "Nam Định",
    "寧平": "Ninh Bình",
    "太平": "Thái Bình",
    "永福": "Vĩnh Phúc",
    "清化": "Thanh Hóa",
    "乂安": "Nghệ An",
    "河靜": "Hà Tĩnh",
    "廣平": "Quảng Bình",
    "廣治": "Quảng Trị",
    "承天順化": "Thừa Thiên Huế",
    "峴港": "Đà Nẵng",
    "廣南": "Quảng Nam",
    "廣義": "Quảng Ngãi",
    "平定": "Bình Định",
    "富安": "Phú Yên",
    "慶和": "Khánh Hòa",
    "寧順": "Ninh Thuận",
    "平順": "Bình Thuận",
    "崑嵩": "Kon Tum",
    "嘉萊": "Gia Lai",
    "多樂": "Đắk Lắk",
    "得農": "Đắk Nông",
    "林同": "Lâm Đồng",
    "平福": "Bình Phước",
    "平陽": "Bình Dương",
    "同奈": "Đồng Nai",
    "西寧": "Tây Ninh",
    "巴地頭頓": "Bà Rịa - Vũng Tàu",
    "胡志明市": "TP Hồ Chí Minh",
    "隆安": "Long An",
    "同塔": "Đồng Tháp",
    "前江": "Tiền Giang",
    "安江": "An Giang",
    "檳椥": "Bến Tre",
    "永隆": "Vĩnh Long",
    "茶榮": "Trà Vinh",
    "後江": "Hậu Giang",
    "堅江": "Kiên Giang",
    "朔莊": "Sóc Trăng",
    "薄遼": "Bạc Liêu",
    "金甌": "Cà Mau",
    "芹苴": "Cần Thơ",
}


# --- HELPERS ---
def is_chinese(text: str) -> bool:
    if not text: return False
    cjk_count = sum(1 for c in text if '\u4e00' <= c <= '\u9fff' or '\u3400' <= c <= '\u4dbf' or '\uf900' <= c <= '\ufaff')
    non_space = len(text.replace(' ', '').replace(',', '').replace('，', '').replace('、', '').replace('。', ''))
    return non_space > 0 and (cjk_count / non_space) >= 0.5

def translate_fixed(text: str) -> str:
    if not text: return text
    return FIXED_TRANS.get(text.strip().lower(), text)

def translate_name(text: str) -> str:
    """Dành riêng cho dịch Họ Tên: Ưu tiên từ điển tên để tránh nhầm với tiếng Anh"""
    if not text or not text.strip() or is_chinese(text): return text

    # 1. Thử dịch từ từ điển tên riêng
    # Hàm trả về None nếu có phần nào không tìm thấy → fallback Google Translate
    dict_result = get_vietnamese_name_in_chinese(text.strip())
    if dict_result is not None:
        return dict_result

    # 2. Nếu không có trong từ điển, dùng Google Translate
    try:
        result = GoogleTranslator(source='vi', target='zh-TW').translate(text.strip())
        return result if result else text
    except: return text

def translate_free(text: str) -> str:
    """Dành cho dịch nội dung tự do (công việc, địa chỉ): Ưu tiên thuật ngữ nghề nghiệp"""
    if not text or not text.strip() or is_chinese(text): return text
    
    # 1. Kiểm tra khớp chính xác trong FIXED_TRANS
    fixed = FIXED_TRANS.get(text.strip().lower())
    if fixed: return fixed
    
    # 2. Xử lý các từ khóa quan trọng TRONG câu (ví dụ: "may" -> "縫紉")
    # Để tránh Google dịch nhầm "may" thành "có lẽ/có thể" (possibly)
    processed_text = text
    # Các từ khóa cần bảo vệ (case insensitive)
    protected_terms = {
        'may': '縫紉',
        'thợ may': '縫紉',
        'may mặc': '縫紉',
        'thợ sơn': '噴漆',
        'thợ hàn': '焊接',
        'thợ điện': '電工',
        'lái xe': '駕駛',
        'nghệ an': '藝安',
    }
    
    # Nếu text là một từ đơn nằm trong danh sách bảo vệ, trả về luôn
    if text.strip().lower() in protected_terms:
        return protected_terms[text.strip().lower()]

    # 3. Dùng Google Translate cho đoạn văn
    try:
        result = GoogleTranslator(source='vi', target='zh-TW').translate(processed_text.strip())
        # Nếu Google dịch nhầm "may" -> "可能" hoặc "Nghệ An" -> "义安", ta sửa lại
        if '可能' in result and 'may' in text.lower():
            result = result.replace('可能', '縫紉')
        if '义安' in result and 'nghệ an' in text.lower():
            result = result.replace('义安', '藝安')
        return result if result else text
    except: return text

def sanitize_filename_master(name):
    if not name: return "UnNamed"
    s = normalize('NFKD', str(name)).encode('ascii', 'ignore').decode('ascii')
    s = re.sub(r'[^\w\s-]', '', s).strip()
    s = re.sub(r'[-\s]+', '_', s)
    return s

def fmt_date(d):
    if not d: return ""
    try: return datetime.strptime(str(d), "%Y-%m-%d").strftime("%d/%m/%Y")
    except: return str(d)

def calc_age(birth_str):
    if not birth_str: return ""
    try:
        b = datetime.strptime(str(birth_str), "%Y-%m-%d")
        today = date.today()
        return today.year - b.year - ((today.month, today.day) < (b.month, b.day))
    except: return ""

def chk(v):
    return v in (True, 'true', '1', 1, 'yes', 'on', 'checked')

def get_base64_image(path, max_size=None, quality=80):
    if not os.path.exists(path): return ""
    try:
        if max_size:
            from PIL import Image
            img = Image.open(path)
            img.thumbnail((max_size, max_size), Image.LANCZOS)
            if img.mode in ("RGBA", "P"): img = img.convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=quality)
            return f"data:image/jpeg;base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"
        with open(path, "rb") as f:
            ext = path.split('.')[-1].lower()
            return f"data:image/{ext};base64,{base64.b64encode(f.read()).decode('utf-8')}"
    except: return ""

# --- CACHE ---
_LOGO_B64_CACHE = None
_BG_B64_CACHE   = None
_TEMPLATE_CACHE = None

def _init_cache():
    global _LOGO_B64_CACHE, _BG_B64_CACHE, _TEMPLATE_CACHE
    _LOGO_B64_CACHE = get_base64_image(os.path.join(BASE_DIR, 'static', 'logo.png'))
    _BG_B64_CACHE   = get_base64_image(os.path.join(BASE_DIR, 'static', 'fct_bg.png'), max_size=400, quality=75)
    try:
        with open(os.path.join(BASE_DIR, 'templates', 'fct_template_v6.18.html'), 'r', encoding='utf-8') as f:
            _TEMPLATE_CACHE = f.read()
    except: pass

_init_cache()


def _fetch_r2_image_as_base64(url: str) -> str:
    """Tải ảnh từ URL R2 về và chuyển thành Base64 để nhúng vào HTML offline.
    Đây là cầu nối giữa R2 (lưu file thật) và file HTML (cần Base64 để offline được).
    """
    if not url or not url.startswith('http'):
        return ""
    try:
        resp = requests.get(url, timeout=10)
        if resp.status_code != 200:
            print(f"⚠️ R2 fetch failed ({resp.status_code}): {url}")
            return ""
        # Xác định mime type từ Content-Type header hoặc extension URL
        content_type = resp.headers.get('Content-Type', 'image/jpeg').split(';')[0].strip()
        b64 = base64.b64encode(resp.content).decode('utf-8')
        return f"data:{content_type};base64,{b64}"
    except Exception as e:
        print(f"⚠️ R2 fetch error: {e}")
        return ""

# --- LOGIC PREPARE ---
def prepare_render_data(raw_data: dict) -> dict:
    data = {}
    fields = [
        'Maso', 'Hoten', 'TentiengTrung', 'Ngaysinh', 'Tuoi', 'Chieucao', 'Cannang', 
        'Lienhe', 'Noio', 'HotenBo', 'TB', 'HotenMe', 'TM', 'VoChong', 'VC', 
        'Socon', 'Anhchiem', 'Xepthu', 'f48', 'N1', 'N2', 'N3', 'ndcv1', 'ndcv2', 'ndcv3', 
        'loi_binh_1', 'Honnhan', 'Hocvan', 'QG1', 'QG2', 'QG3', 'video_link_1', 'video_link_2'
    ]
    for f in fields:
        val = str(raw_data.get(f, '')).strip()
        if f in ['Honnhan', 'Hocvan']: data[f] = FIXED_TRANS.get(val.lower(), val)
        elif f in ['QG1', 'QG2', 'QG3']: data[f] = translate_fixed(val)
        else: data[f] = val

    data['Ngaysinh'] = fmt_date(data['Ngaysinh'])
    if not data['Tuoi'] and raw_data.get('Ngaysinh'): data['Tuoi'] = calc_age(raw_data.get('Ngaysinh'))

    fields_to_translate = ['Noio', 'ndcv1', 'ndcv2', 'ndcv3', 'loi_binh_1', 'N1', 'N2', 'N3']
    non_empty = {f: data[f] for f in fields_to_translate if data.get(f, '').strip()}
    if non_empty:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        with ThreadPoolExecutor(max_workers=min(len(non_empty), 4)) as executor:
            future_map = {executor.submit(translate_free, v): k for k, v in non_empty.items()}
            for future in as_completed(future_map):
                field = future_map[future]
                try: data[field] = future.result()
                except: pass

    yellow_alerts = []
    for i in range(1, 23):
        key = f'f{i:02d}'
        if chk(raw_data.get(key)) and key in YELLOW_ALERTS_MAP: yellow_alerts.append(YELLOW_ALERTS_MAP[key])
    if yellow_alerts:
        alert_str = "、".join(yellow_alerts)
        data['loi_binh_1'] = (data['loi_binh_1'] + "、" + alert_str) if data.get('loi_binh_1') else alert_str

    skills_html = []
    for key, name in SKILL_MAPPING.items():
        if chk(raw_data.get(key)):
            skills_html.append(f'<span class="skill-tag">{name}</span>')
    data['KyNangList_HTML'] = "".join(skills_html)

    tay = []
    if chk(raw_data.get('f01')): tay.append("右手")
    if chk(raw_data.get('f07')): tay.append("左手")
    data['TayThuan'] = " / ".join(tay) if tay else "右手"
    
    vision_map = [('f02', "右眼受損"), ('f08', "左眼受損"), ('f16', "散光"), ('f17', "色盲"), ('f15', "近視")]
    vision = [label for k, label in vision_map if chk(raw_data.get(k))]
    data['ThiLuc'] = " / ".join(vision) if vision else "正常"
    data['f12'] = "Có / 有" if chk(raw_data.get('f12')) else "無"

    for key in ('photo', 'qr_line'):
        path = raw_data.get(key, '')
        if isinstance(path, str) and path.startswith('data:image/'):
            # Đã là Base64 (upload từ Local hoặc cũ) → dùng luôn
            data[f'{key}_base64'] = path
        elif isinstance(path, str) and path.startswith('http'):
            # Là URL từ Cloudflare R2 → Kéo về và chuyển thành Base64 để nhúng offline
            print(f"🔄 Fetching {key} from R2 for offline HTML...")
            data[f'{key}_base64'] = _fetch_r2_image_as_base64(path)
        else:
            data[f'{key}_base64'] = ""

    # Xử lý ảnh tài liệu (giấy tờ) để render trang 2
    data['document_images'] = raw_data.get('document_images', [])

    return data

def _protect_html(html: str) -> str:
    html = re.sub(r'<!--(?!\[if).*?-->', '', html, flags=re.DOTALL)
    html = re.sub(r'>\s+<', '><', html)
    html = re.sub(r'\s{2,}', ' ', html).strip()
    anti_devtools = (
        '<script>'
        '(function(){'
        'document.addEventListener("contextmenu",function(e){e.preventDefault();},false);'
        'document.addEventListener("keydown",function(e){'
        'if(e.keyCode===123||(e.ctrlKey&&e.shiftKey&&(e.keyCode===73||e.keyCode===74||e.keyCode===67))||(e.ctrlKey&&e.keyCode===85)){'
        'e.preventDefault();e.stopPropagation();return false;}'
        '});'
        # Cải thiện devtools detection: chỉ kích hoạt khi THỰC SỰ có DevTools
        # Kiểm tra cả width/height diff VÀ tỷ lệ zoom để tránh false positive
        'var _t=function(){'
        'var wDiff=window.outerWidth-window.innerWidth;'
        'var hDiff=window.outerHeight-window.innerHeight;'
        'var zoom=window.devicePixelRatio||1;'
        # Chỉ trigger khi diff lớn VÀ zoom gần 1 (không phải user zoom)
        'if((wDiff>400||hDiff>400)&&zoom>=0.8&&zoom<=1.2){'
        'document.body.innerHTML="";'
        '}'
        '};'
        'setInterval(_t,1000);'
        '})();'
        '</script>'
    )
    return html.replace('</body>', anti_devtools + '</body>')

def generate_html_resume(form_data: dict, template_name='fct_template_v6.18.html') -> str:
    processed_data = prepare_render_data(form_data)
    processed_data['logo_base64'] = _LOGO_B64_CACHE or get_base64_image(os.path.join(BASE_DIR, 'static', 'logo.png'))
    processed_data['bg_base64']   = _BG_B64_CACHE   or get_base64_image(os.path.join(BASE_DIR, 'static', 'fct_bg.png'), max_size=400, quality=75)
    
    # Nhúng dữ liệu gốc (không chứa ảnh base64 nặng) để có thể nạp lại sau này
    raw_for_embed = {k: v for k, v in form_data.items() if k not in ('photo', 'qr_line', 'document_images')}
    # Đánh dấu nếu có ảnh
    if form_data.get('photo'): raw_for_embed['__has_photo'] = True
    if form_data.get('qr_line'): raw_for_embed['__has_qr'] = True
    raw_json_str = json.dumps(raw_for_embed, ensure_ascii=False)
    
    # Dùng placeholder để _protect_html không phá hủy JSON
    processed_data['raw_data_json'] = '___FCT_RAW_PLACEHOLDER___'
    
    if _TEMPLATE_CACHE: template = Template(_TEMPLATE_CACHE)
    else:
        with open(os.path.join(BASE_DIR, 'templates', template_name), 'r', encoding='utf-8') as f:
            template = Template(f.read())
    html = _protect_html(template.render(processed_data))
    # Thay placeholder bằng JSON thật SAU KHI minify xong
    return html.replace('___FCT_RAW_PLACEHOLDER___', raw_json_str)

def _resize_image_for_db(data_uri: str, max_px: int = 1200, quality: int = 85) -> str:
    if not data_uri or not data_uri.startswith('data:image/'): return data_uri
    try:
        from PIL import Image
        header, encoded = data_uri.split(',', 1)
        img_bytes = base64.b64decode(encoded)
        img = Image.open(io.BytesIO(img_bytes))
        if max(img.width, img.height) > max_px: img.thumbnail((max_px, max_px), Image.LANCZOS)
        buf = io.BytesIO()
        if img.mode in ('RGBA', 'P'): img = img.convert('RGB')
        img.save(buf, format='JPEG', quality=quality, optimize=True)
        return f"data:image/jpeg;base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"
    except: return data_uri

def _prepare_data_for_db(data: dict) -> dict:
    clean = dict(data)
    for key in ('photo', 'qr_line'):
        if clean.get(key): clean[key] = _resize_image_for_db(clean[key])
    # Resize ảnh tài liệu nếu có - Giảm sâu hơn để tránh lỗi 413 (900px, quality 60)
    if clean.get('document_images') and isinstance(clean['document_images'], list):
        clean['document_images'] = [_resize_image_for_db(img, max_px=900, quality=60) for img in clean['document_images'] if img]
    return clean

# --- API ROUTES ---
@app.route('/')
def user_form(): return render_template('user_form.html')

@app.route('/fct-1503')
@auth_required
def index():
    resp = make_response(render_template('index.html'))
    resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

@app.route('/api/health')
def health(): return jsonify({'ok': True, 'msg': 'DAS V3.0 running'})

def _process_form_data(request):
    if request.content_type and 'multipart/form-data' in request.content_type:
        data = json.loads(request.form.get('data', '{}'))
        for key in ('photo', 'qr_line'):
            file = request.files.get(key)
            if file:
                ext = os.path.splitext(file.filename)[1][1:].lower() or 'png'
                data[key] = f"data:image/{'jpeg' if ext=='jpg' else ext};base64,{base64.b64encode(file.read()).decode('utf-8')}"
    else: data = request.get_json() or {}
    return data

@app.route('/api/submit-only', methods=['POST'])
def api_submit_only():
    try:
        data = _process_form_data(request)
        record_id = data.get('_record_id')
        ma_so = str(data.get('Maso', '')).strip()
        ho_ten = str(data.get('Hoten', '')).strip()

        if ma_so and ma_so.upper() != 'CHO_DUYET':
            existing = FormHistory.query.filter_by(ma_so=ma_so).first()
            if existing and (not record_id or int(record_id) != existing.id):
                return jsonify({'success': False, 'error': f'Mã số "{ma_so}" đã tồn tại.'}), 400

        if record_id:
            record = FormHistory.query.get(int(record_id))
            if not record: return jsonify({'success': False, 'error': 'Not found'}), 404
            old_data = json.loads(record.data_json) if record.data_json else {}
            for key in ('photo', 'qr_line', 'document_images'):
                if key not in data and old_data.get(key): data[key] = old_data[key]
            record.ma_so = ma_so or 'CHO_DUYET'
            record.ho_ten = ho_ten
            record.data_json = json.dumps(_prepare_data_for_db(data), ensure_ascii=False)
            db.session.commit()
            msg = 'Đã cập nhật.'
        else:
            record = FormHistory(ma_so=ma_so or 'CHO_DUYET', ho_ten=ho_ten, data_json=json.dumps(_prepare_data_for_db(data), ensure_ascii=False))
            db.session.add(record)
            db.session.commit()
            msg = 'Đã nộp form.'

        return jsonify({'success': True, 'id': record.id, 'ma_so': record.ma_so, 'msg': msg})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/translate', methods=['POST'])
def api_translate():
    try:
        data = request.get_json() or {}
        # Sử dụng translate_name cho API này vì nó thường được gọi cho ô Họ Tên
        return jsonify({'success': True, 'translated': translate_name(data.get('text', ''))})
    except: return jsonify({'success': False}), 500

@app.route('/api/download-cv/<maso>', methods=['GET'])
@auth_required
def download_history(maso):
    try:
        record = FormHistory.query.filter_by(ma_so=maso).order_by(FormHistory.ngay_tao.desc()).first()
        if not record: return jsonify({"error": "Not found"}), 404
        html_content = generate_html_resume(json.loads(record.data_json))
        filename = f"{maso}_{sanitize_filename_master(record.ho_ten)}.html"
        return send_file(io.BytesIO(html_content.encode('utf-8')), mimetype='text/html', as_attachment=True, download_name=filename)
    except: return jsonify({"error": "Error"}), 400

@app.route('/resume-<int:record_id>.html')
@auth_required
def api_preview(record_id):
    try:
        record = FormHistory.query.get(record_id)
        if not record: return "Not found", 404
        html_content = generate_html_resume(json.loads(record.data_json))
        return Response(html_content, mimetype="text/html", headers={"Content-Type": "text/html; charset=utf-8"})
    except Exception as e: return str(e), 500

@app.route('/cv/<path:slug>', methods=['GET'])
def secure_web_view(slug):
    try:
        # Hỗ trợ 2 định dạng:
        # 1. /cv/92/MD13243_Nguyen_Van_Chuc (id/slug) -> Ưu tiên ID (chính xác nhất)
        # 2. /cv/MD13243_Nguyen_Van_Chuc (slug) -> Tìm theo Maso (tiện lợi)
        
        parts = slug.split('/')
        record = None
        
        if len(parts) >= 2:
            # Định dạng id/slug
            try:
                rid = int(parts[0])
                record = FormHistory.query.get(rid)
            except: pass
        
        if not record:
            # Định dạng slug (tìm theo Maso)
            # Tách lấy phần trước dấu _ đầu tiên hoặc dùng cả slug nếu không có _
            maso_part = slug.split('_')[0] if '_' in slug else slug
            # Trường hợp đặc biệt: Nếu slug là số nguyên, thử tìm theo ID
            if slug.isdigit():
                record = FormHistory.query.get(int(slug))
            
            if not record:
                # Tìm bản ghi mới nhất có mã số này
                record = FormHistory.query.filter_by(ma_so=maso_part).order_by(FormHistory.ngay_tao.desc()).first()

        if not record: return "Không tìm thấy hồ sơ / Not found", 404
        
        # Kiểm tra maso trong slug (nếu có id/maso) để đảm bảo tính bảo mật/nhất quán
        # (Nếu dùng Maso từ slug thì record đã khớp rồi)
        
        html_content = generate_html_resume(json.loads(record.data_json))
        # Tạo tên file đẹp cho trình duyệt
        clean_name = sanitize_filename_master(record.ho_ten)
        filename = f"{record.ma_so}_{clean_name}.html"
        encoded_filename = quote(filename)
        
        response = Response(html_content, mimetype="text/html")
        response.headers["Content-Type"] = "text/html; charset=utf-8"
        response.headers["Content-Disposition"] = f'inline; filename="{filename}"; filename*=UTF-8\'\'{encoded_filename}'
        return response
    except Exception as e:
        traceback.print_exc()
        return str(e), 500


@app.route('/api/history', methods=['GET'])
@auth_required
def api_history():
    try:
        # Tối ưu: Chỉ lấy metadata, không lấy data_json nặng nề
        records = FormHistory.query.filter_by(is_deleted=False).order_by(FormHistory.ngay_tao.desc()).limit(100).all()
        vietnam_tz = timezone(timedelta(hours=7))
        data = [{
            'id': r.id, 'ma_so': r.ma_so, 'ho_ten': r.ho_ten,
            'is_selected': getattr(r, 'is_selected', False),  # Fallback nếu field chưa tồn tại
            'ngay_tao': r.ngay_tao.replace(tzinfo=timezone.utc).astimezone(vietnam_tz).strftime("%d/%m/%Y %H:%M:%S") if r.ngay_tao else ''
        } for r in records]
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/history/<int:record_id>/data', methods=['GET'])
@auth_required
def api_history_data(record_id):
    try:
        record = FormHistory.query.get(record_id)
        if not record: return jsonify({'success': False, 'error': 'Not found'}), 404
        return jsonify({
            'success': True, 
            'data_json': json.loads(record.data_json) if record.data_json else {}
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/history/bulk-delete', methods=['POST'])
@auth_required
def api_bulk_delete_history():
    try:
        data = request.get_json() or {}
        ids = data.get('ids', [])
        if not ids: return jsonify({'success': False, 'error': 'No IDs'}), 400
        records = FormHistory.query.filter(FormHistory.id.in_([int(i) for i in ids])).all()
        for r in records: 
            r.is_deleted = True
            try:
                if r.data_json:
                    jd = json.loads(r.data_json)
                    jd.pop('photo', None)
                    jd.pop('qr_line', None)
                    jd.pop('document_images', None)
                    r.data_json = json.dumps(jd)
            except: pass
        db.session.commit()
        return jsonify({'success': True, 'deleted': len(records)})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/history/<int:record_id>', methods=['DELETE'])
@auth_required
def api_delete_history(record_id):
    try:
        r = FormHistory.query.get(record_id)
        if r:
            r.is_deleted = True
            try:
                if r.data_json:
                    jd = json.loads(r.data_json)
                    jd.pop('photo', None)
                    jd.pop('qr_line', None)
                    jd.pop('document_images', None)
                    r.data_json = json.dumps(jd)
            except: pass
            db.session.commit()
        return jsonify({'success': True})
    except: return jsonify({'success': False}), 500

@app.route('/api/history/hard-delete-year', methods=['POST'])
@auth_required
def api_hard_delete_year():
    try:
        data = request.get_json() or {}
        year_val = data.get('year')
        if not year_val or str(year_val) == 'ALL': return jsonify({'success': False, 'error': 'Invalid year'}), 400
        from sqlalchemy import extract
        records = FormHistory.query.filter(extract('year', FormHistory.ngay_tao) == int(year_val)).all()
        for r in records: db.session.delete(r)
        db.session.commit()
        return jsonify({'success': True, 'deleted': len(records)})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/history/bulk-download', methods=['POST'])
@auth_required
def api_bulk_download():
    try:
        data = request.get_json() or {}
        ids = data.get('ids', [])
        if not ids: return jsonify({'success': False, 'error': 'No IDs'}), 400
        records = FormHistory.query.filter(FormHistory.id.in_([int(i) for i in ids])).all()
        if not records: return jsonify({'success': False, 'error': 'No records found'}), 404
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for r in records:
                try:
                    form_data = json.loads(r.data_json)
                    html_content = generate_html_resume(form_data)
                    filename = f"{r.ma_so}_{sanitize_filename_master(r.ho_ten)}.html"
                    zf.writestr(filename, html_content)
                except: pass
        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name='FCT_HoSo_Export.zip')
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

SKILL_MAPPING = {
    'f23': 'Hàn điện', 'f31': 'Tiện', 'f39': 'Nhựa',
    'f24': 'Hàn argon', 'f32': 'Phay', 'f40': 'Xây dựng',
    'f25': 'Hàn CO2', 'f33': 'Bào', 'f41': 'Sữa chữa máy',
    'f26': 'Tig Mig', 'f34': 'CNC', 'f42': 'Điều dưỡng',
    'f27': 'Đúc', 'f35': 'Đột dập', 'f43': 'Giúp việc',
    'f28': 'Dệt', 'f36': 'In ấn', 'f44': 'Xe cẩu',
    'f29': 'May', 'f37': 'Thợ mộc', 'f45': 'Cẩu trục',
    'f30': 'Xe nâng', 'f38': 'Lái xe tải/khách', 'f46': 'Máy xúc'
}

@app.route('/api/history/export-excel', methods=['POST'])
@auth_required
def api_export_excel():
    try:
        data = request.get_json() or {}
        ids = data.get('ids', [])
        year_val = data.get('year')
        
        if ids:
            records = FormHistory.query.filter(FormHistory.id.in_([int(i) for i in ids])).all()
        elif year_val and str(year_val) != 'ALL':
            from sqlalchemy import extract
            records = FormHistory.query.filter(extract('year', FormHistory.ngay_tao) == int(year_val)).all()
        else:
            records = FormHistory.query.all()
            
        if not records:
            return jsonify({'success': False, 'error': 'No records found'}), 404
            
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Danh sách ứng viên"
        
        headers = ['Mã số', 'Họ tên', 'Ngày sinh', 'Trạng thái', 'Chiều cao (cm)', 'Cân nặng (kg)', 
                   'Trình độ văn hóa', 'Nơi ở', 'Tay nghề', 'Kinh nghiệm công việc', 'Người phụ trách']
        ws.append(headers)
        
        # Thống kê
        skills_count = {}
        edu_count = {}
        location_count = {}
        type_count = {'Nam (MD)': 0, 'Nữ (FD)': 0, 'Điều dưỡng (KD)': 0, 'Khác': 0}
        job_type_count = {'Nam công xưởng': 0, 'Nữ công xưởng': 0, 'Giúp việc & Khác': 0, 'Điều dưỡng': 0}
        
        for r in records:
            try:
                form_data = json.loads(r.data_json)
                ma_so = r.ma_so or ''
                ma_so_upper = ma_so.strip().upper()
                if ma_so_upper.startswith('MD'):
                    type_count['Nam (MD)'] += 1
                elif ma_so_upper.startswith('FD'):
                    type_count['Nữ (FD)'] += 1
                elif ma_so_upper.startswith('KD'):
                    type_count['Điều dưỡng (KD)'] += 1
                else:
                    type_count['Khác'] += 1
                    
                # Phân loại nhóm ngành nghề
                is_giup_viec = chk(form_data.get('f43')) # f43: Giúp việc
                if ma_so_upper.startswith('MD'):
                    job_type_count['Nam công xưởng'] += 1
                elif ma_so_upper.startswith('KD'):
                    job_type_count['Điều dưỡng'] += 1
                elif is_giup_viec:
                    job_type_count['Giúp việc & Khác'] += 1
                elif ma_so_upper.startswith('FD'):
                    job_type_count['Nữ công xưởng'] += 1
                else:
                    job_type_count['Giúp việc & Khác'] += 1
                    
                ho_ten = r.ho_ten or ''
                ngay_sinh = form_data.get('Ngaysinh', '')
                chieu_cao = form_data.get('Chieucao', '')
                can_nang = form_data.get('Cannang', '')
                hoc_van = form_data.get('Hocvan', '')
                noi_o = form_data.get('Noio', '')
                
                # Dịch học lực và nơi ở sang tiếng Việt cho báo cáo Excel
                hoc_van_vi = ZH_TO_VI.get(hoc_van.strip(), hoc_van) if hoc_van else ''
                noi_o_vi = ZH_TO_VI.get(noi_o.strip(), noi_o) if noi_o else ''
                
                nguoi_pt = form_data.get('f48', '')
                
                if hoc_van_vi:
                    edu_count[hoc_van_vi] = edu_count.get(hoc_van_vi, 0) + 1
                    
                if noi_o_vi:
                    noi_o_clean = noi_o_vi.strip()
                    if noi_o_clean:
                        location_count[noi_o_clean] = location_count.get(noi_o_clean, 0) + 1
                
                # Tay nghề
                skills = []
                for k, v in SKILL_MAPPING.items():
                    if form_data.get(k):
                        skills.append(v)
                        skills_count[v] = skills_count.get(v, 0) + 1
                tay_nghe = ", ".join(skills)
                
                # Kinh nghiệm
                kn = []
                for i in range(1, 4):
                    qg = form_data.get(f'QG{i}', '')
                    nam = form_data.get(f'N{i}', '')
                    cv = form_data.get(f'ndcv{i}', '')
                    if qg or cv:
                        parts = []
                        if qg: parts.append(qg)
                        if nam: parts.append(f"({nam})")
                        if cv: parts.append(f": {cv}")
                        kn.append(" ".join(parts))
                kinh_nghiem = "\n".join(kn)
                
                # Trạng thái
                trang_thai = '🎯 Trúng tuyển' if getattr(r, 'is_selected', False) else '📝 Gửi form'
                
                ws.append([ma_so, ho_ten, ngay_sinh, trang_thai, chieu_cao, can_nang, hoc_van_vi, noi_o_vi, tay_nghe, kinh_nghiem, nguoi_pt])
            except Exception as e:
                print("Error parsing record:", e)
                pass
                
        # Premium Styling
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # FCT Blue
        header_font = Font(name="Arial", color="FFFFFF", bold=True, size=11)
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # Slate-50
        body_font = Font(name="Arial", size=10)
        
        ws.row_dimensions[1].height = 25 # Header height
        ws.freeze_panes = 'A2'
        
        for col_num, cell in enumerate(ws[1], 1):
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            
        selected_row_fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid") # Xanh lá pastel (Emerald-100)
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=11), 2):
            is_even = (row_idx % 2 == 0)
            ws.row_dimensions[row_idx].height = 35 # Tăng chiều cao mặc định cho dòng
            
            is_selected_row = (row[3].value == '🎯 Trúng tuyển')
            row_fill = selected_row_fill if is_selected_row else (zebra_fill if is_even else None)
            
            for cell in row:
                cell.font = body_font
                cell.border = thin_border
                # Column Kinh nghiệm (cột 10) căn trái, các cột khác căn giữa
                if cell.column == 10:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if row_fill:
                    cell.fill = row_fill
                    
        # Auto-fit columns
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    lines = str(cell.value).split('\n')
                    for line in lines:
                        if len(line) > max_length:
                            max_length = len(line)
                except:
                    pass
            adjusted_width = min((max_length + 3), 60) # Max 60, thêm padding
            ws.column_dimensions[column].width = adjusted_width
            
        # Thêm AutoFilter
        ws.auto_filter.ref = ws.dimensions
            
        # Thêm Sheet Thống Kê
        ws_stat = wb.create_sheet(title="Thống Kê")
        
        # 1. Dashboard Title (Merged A1:N1)
        ws_stat.merge_cells("A1:N1")
        title_cell = ws_stat.cell(row=1, column=1, value="BÁO CÁO THỐNG KÊ HỒ SƠ ỨNG VIÊN")
        title_cell.font = Font(name="Arial", size=14, bold=True, color="1E3A8A")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_stat.row_dimensions[1].height = 35
        
        # 2. KPI Cards
        # Card 1: Tổng số ứng viên (A2:B4)
        ws_stat.merge_cells("A2:B4")
        card1 = ws_stat.cell(row=2, column=1, value=f"TỔNG SỐ ỨNG VIÊN\n\n{len(records)}")
        card1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Card 2: Nam (D2:E4)
        ws_stat.merge_cells("D2:E4")
        card2 = ws_stat.cell(row=2, column=4, value=f"NAM (MD)\n\n{type_count['Nam (MD)']}")
        card2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Card 3: Nữ (G2:H4)
        ws_stat.merge_cells("G2:H4")
        card3 = ws_stat.cell(row=2, column=7, value=f"NỮ (FD)\n\n{type_count['Nữ (FD)']}")
        card3.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Card 4: Điều dưỡng (J2:K4)
        ws_stat.merge_cells("J2:K4")
        card4 = ws_stat.cell(row=2, column=10, value=f"ĐIỀU DƯỠNG (KD)\n\n{type_count['Điều dưỡng (KD)']}")
        card4.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Card 5: Khác (M2:N4)
        ws_stat.merge_cells("M2:N4")
        card5 = ws_stat.cell(row=2, column=13, value=f"MÃ KHÁC\n\n{type_count['Khác']}")
        card5.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Style KPI Cards
        kpi_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") # Slate-100
        kpi_border = Border(left=Side(style='thin', color="CBD5E1"),
                            right=Side(style='thin', color="CBD5E1"),
                            top=Side(style='thin', color="CBD5E1"),
                            bottom=Side(style='thin', color="CBD5E1"))
        
        for col_start in [1, 4, 7, 10, 13]:
            for r in range(2, 5):
                for c in range(col_start, col_start + 2):
                    cell = ws_stat.cell(row=r, column=c)
                    cell.fill = kpi_fill
                    cell.border = kpi_border
            # Đặt font cho cell đầu tiên của mỗi card
            first_cell = ws_stat.cell(row=2, column=col_start)
            first_cell.font = Font(name="Arial", bold=True, size=11, color="1E3A8A")
            
        ws_stat.row_dimensions[2].height = 20
        ws_stat.row_dimensions[3].height = 20
        ws_stat.row_dimensions[4].height = 20
        
        # 3. Dynamic Tables in Columns A & B
        
        # Subsection I: Tay nghề
        row_skill = 6
        ws_stat.cell(row=row_skill - 1, column=1, value="I. THỐNG KÊ TAY NGHỀ & KỸ NĂNG").font = Font(name="Arial", bold=True, size=11, color="1E3A8A")
        ws_stat.cell(row=row_skill, column=1, value="Tay nghề")
        ws_stat.cell(row=row_skill, column=2, value="Số lượng")
        for idx, (skill, count) in enumerate(skills_count.items(), 1):
            ws_stat.cell(row=row_skill + idx, column=1, value=skill)
            ws_stat.cell(row=row_skill + idx, column=2, value=count)
            
        # Subsection II: Nhóm ngành nghề
        row_job = row_skill + len(skills_count) + 3
        ws_stat.cell(row=row_job - 1, column=1, value="II. PHÂN BỔ NHÓM NGÀNH NGHỀ").font = Font(name="Arial", bold=True, size=11, color="1E3A8A")
        ws_stat.cell(row=row_job, column=1, value="Nhóm ngành nghề")
        ws_stat.cell(row=row_job, column=2, value="Số lượng")
        for idx, (job, count) in enumerate(job_type_count.items(), 1):
            ws_stat.cell(row=row_job + idx, column=1, value=job)
            ws_stat.cell(row=row_job + idx, column=2, value=count)
            
        # Subsection III: Trình độ văn hóa
        row_edu = row_job + len(job_type_count) + 3
        ws_stat.cell(row=row_edu - 1, column=1, value="III. PHÂN BỔ TRÌNH ĐỘ VĂN HÓA").font = Font(name="Arial", bold=True, size=11, color="1E3A8A")
        ws_stat.cell(row=row_edu, column=1, value="Trình độ văn hóa")
        ws_stat.cell(row=row_edu, column=2, value="Số lượng")
        for idx, (edu, count) in enumerate(edu_count.items(), 1):
            ws_stat.cell(row=row_edu + idx, column=1, value=edu)
            ws_stat.cell(row=row_edu + idx, column=2, value=count)
            
        # Subsection IV: Nơi ở / Quê quán
        row_loc = row_edu + len(edu_count) + 3
        ws_stat.cell(row=row_loc - 1, column=1, value="IV. PHÂN BỔ THEO NƠI Ở / QUÊ QUÁN").font = Font(name="Arial", bold=True, size=11, color="1E3A8A")
        ws_stat.cell(row=row_loc, column=1, value="Nơi ở / Quê quán")
        ws_stat.cell(row=row_loc, column=2, value="Số lượng")
        sorted_locations = sorted(location_count.items(), key=lambda x: x[1], reverse=True)
        for idx, (loc, count) in enumerate(sorted_locations, 1):
            ws_stat.cell(row=row_loc + idx, column=1, value=loc)
            ws_stat.cell(row=row_loc + idx, column=2, value=count)

        # 4. Charts in Columns D to N (Never Overlapping with Columns A & B)
        
        # Chart 1: Tay nghề (Col BarChart)
        if skills_count:
            chart1 = BarChart()
            chart1.type = "col"
            chart1.style = 10
            chart1.title = "Thống kê Tay nghề (Kỹ năng)"
            chart1.y_axis.title = 'Số lượng ứng viên'
            chart1.x_axis.title = 'Tay nghề'
            chart1.width = 17
            chart1.height = 7.5
            
            data1 = Reference(ws_stat, min_col=2, min_row=row_skill, max_row=row_skill+len(skills_count), max_col=2)
            cats1 = Reference(ws_stat, min_col=1, min_row=row_skill+1, max_row=row_skill+len(skills_count))
            chart1.add_data(data1, titles_from_data=True)
            chart1.set_categories(cats1)
            chart1.legend = None
            chart1.shape = 4
            
            if chart1.series:
                series = chart1.series[0]
                series.dLbls = DataLabelList()
                series.dLbls.showVal = True
                
            ws_stat.add_chart(chart1, "D6")
            
        # Chart 2: Nhóm ngành nghề (PieChart)
        if job_type_count:
            chart_job = PieChart()
            chart_job.title = "Phân bổ Nhóm ngành nghề"
            chart_job.width = 17
            chart_job.height = 7.5
            chart_job.legend.position = "b"
            
            data_job = Reference(ws_stat, min_col=2, min_row=row_job, max_row=row_job+len(job_type_count))
            labels_job = Reference(ws_stat, min_col=1, min_row=row_job+1, max_row=row_job+len(job_type_count))
            chart_job.add_data(data_job, titles_from_data=True)
            chart_job.set_categories(labels_job)
            
            chart_job.dataLabels = DataLabelList()
            chart_job.dataLabels.showVal = True
            
            ws_stat.add_chart(chart_job, "J6")
            
        # Chart 3: Trình độ văn hóa (PieChart)
        if edu_count:
            pie = PieChart()
            pie.title = "Phân bổ Trình độ văn hóa"
            pie.width = 17
            pie.height = 7.5
            pie.legend.position = "b"
            
            data_edu = Reference(ws_stat, min_col=2, min_row=row_edu, max_row=row_edu+len(edu_count))
            labels_edu = Reference(ws_stat, min_col=1, min_row=row_edu+1, max_row=row_edu+len(edu_count))
            pie.add_data(data_edu, titles_from_data=True)
            pie.set_categories(labels_edu)
            
            pie.dataLabels = DataLabelList()
            pie.dataLabels.showVal = True
            
            ws_stat.add_chart(pie, "D21")
            
        # Chart 4: Nơi ở / Quê quán (PieChart)
        if location_count:
            chart_loc = PieChart()
            chart_loc.title = "Phân bổ theo Nơi ở / Quê quán"
            chart_loc.width = 24
            chart_loc.height = 11
            chart_loc.legend.position = "b"
            
            data_loc = Reference(ws_stat, min_col=2, min_row=row_loc, max_row=row_loc+len(location_count))
            cats_loc = Reference(ws_stat, min_col=1, min_row=row_loc+1, max_row=row_loc+len(location_count))
            chart_loc.add_data(data_loc, titles_from_data=True)
            chart_loc.set_categories(cats_loc)
            
            chart_loc.dataLabels = DataLabelList()
            chart_loc.dataLabels.showVal = True
            
            ws_stat.add_chart(chart_loc, f"D{row_loc}")

        # 5. Định dạng Style Premium cho Sheet Thống Kê
        thin_gray = Border(left=Side(style='thin', color="E2E8F0"), 
                           right=Side(style='thin', color="E2E8F0"), 
                           top=Side(style='thin', color="E2E8F0"), 
                           bottom=Side(style='thin', color="E2E8F0"))
        
        header_rows = {row_skill, row_job, row_edu, row_loc}
        section_rows = {row_skill - 1, row_job - 1, row_edu - 1, row_loc - 1}
        
        for r_idx in range(1, ws_stat.max_row + 1):
            if r_idx not in [1, 2, 3, 4]:
                if r_idx in section_rows:
                    ws_stat.row_dimensions[r_idx].height = 25
                else:
                    ws_stat.row_dimensions[r_idx].height = 20
                
            for c_idx in [1, 2]:
                cell = ws_stat.cell(row=r_idx, column=c_idx)
                if cell.value is not None:
                    if r_idx in header_rows:
                        cell.font = Font(name="Arial", color="FFFFFF", bold=True, size=10)
                        cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif r_idx >= 6 and r_idx not in section_rows:
                        cell.font = Font(name="Arial", size=10)
                        cell.border = thin_gray
                        if c_idx == 2:
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.alignment = Alignment(horizontal="left", vertical="center")

        # Cấu hình kích thước cột cho lưới
        ws_stat.column_dimensions['A'].width = 25
        ws_stat.column_dimensions['B'].width = 12
        ws_stat.column_dimensions['C'].width = 3
        
        for col_let in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
            ws_stat.column_dimensions[col_let].width = 12
            
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_prefix = f"Nam{year_val}_" if year_val and str(year_val) != 'ALL' else ""
        return send_file(excel_buffer, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                         as_attachment=True, download_name=f'FCT_UngVien_{filename_prefix}{timestamp}.xlsx')
    except Exception as e:
        print(traceback.format_exc())
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/history/bulk-print', methods=['POST'])
@auth_required
def api_bulk_print():
    try:
        data = request.get_json() or {}
        ids = data.get('ids', [])
        if not ids: return jsonify({'success': False, 'error': 'No IDs'}), 400
        records = FormHistory.query.filter(FormHistory.id.in_([int(i) for i in ids])).all()
        if not records: return jsonify({'success': False, 'error': 'No records found'}), 404
        
        # Render all individual HTMLs
        htmls = []
        for r in records:
            try:
                form_data = json.loads(r.data_json)
                htmls.append(generate_html_resume(form_data))
            except: pass
            
        if not htmls: return jsonify({'success': False, 'error': 'Failed to render records'}), 500
        
        base_html = htmls[0]
        base_html = base_html.replace('page-break-after: avoid !important;', 'page-break-after: always !important;')
        
        pages_to_append = []
        for html in htmls[1:]:
            start_idx = html.find('<div class="a4-page notranslate">')
            if start_idx != -1:
                end_idx = html.find('<script id="fct-raw-data"', start_idx)
                if end_idx == -1:
                    end_idx = html.find('<script>(function(){document.addEventListener', start_idx)
                if end_idx != -1:
                    pages_to_append.append(html[start_idx:end_idx])
        
        # Insert into base_html
        insert_idx = base_html.find('<script id="fct-raw-data"')
        if insert_idx == -1:
            insert_idx = base_html.find('<script>(function(){document.addEventListener')
            
        if insert_idx != -1:
            final_html = base_html[:insert_idx] + ''.join(pages_to_append) + base_html[insert_idx:]
        else:
            final_html = base_html
            
        return Response(final_html, mimetype="text/html", headers={"Content-Type": "text/html; charset=utf-8"})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ─── TRÚNG TUYỂN: Toggle trạng thái và gửi sang B ───────────────────
@app.route('/api/history/<int:record_id>/toggle-selected', methods=['POST'])
@auth_required
def api_toggle_selected(record_id):
    try:
        record = FormHistory.query.get(record_id)
        if not record:
            return jsonify({'success': False, 'error': 'Not found'}), 404
        
        # Nếu field is_selected chưa tồn tại, set default = False
        if not hasattr(record, 'is_selected') or record.is_selected is None:
            record.is_selected = False
        
        # Toggle trạng thái
        record.is_selected = not record.is_selected
        db.session.commit()
        
        # Nếu trúng tuyển (is_selected = True), gửi dữ liệu sang B
        if record.is_selected:
            try:
                form_data = json.loads(record.data_json) if record.data_json else {}
                
                # Mapping dữ liệu A → B
                worker_data = {
                    'id': record.ma_so or f'worker_{record.id}',
                    'full_name': record.ho_ten or '',
                    'date_of_birth': form_data.get('Ngaysinh') or None,
                    'phone_number': form_data.get('Lienhe') or None,
                    'hometown': form_data.get('Noio') or None,
                    'avatar_url': form_data.get('photo') or '',  # Base64 ảnh
                    'win_date': datetime.utcnow().strftime('%Y-%m-%d'),
                    'status': 'DRAFT',
                    'is_placed': False,
                    'passport_expiry': '',
                    'id_card_expiry': '',
                    'health_check_expiry': '',
                    'judicial_record_2_expiry': '',
                }
                
                # Gửi sang B (Firebase)
                b_api_url = os.environ.get('B_API_URL', 'http://localhost:3000')
                response = requests.post(
                    f'{b_api_url}/api/workers/sync-from-a',
                    json=worker_data,
                    timeout=10
                )
                
                if response.status_code != 200:
                    # Nếu gửi thất bại, rollback toggle
                    record.is_selected = False
                    db.session.commit()
                    return jsonify({
                        'success': False,
                        'error': f'Failed to sync to B: {response.text}'
                    }), 500
            except Exception as e:
                # Nếu có lỗi, rollback toggle
                record.is_selected = False
                db.session.commit()
                traceback.print_exc()
                return jsonify({
                    'success': False,
                    'error': f'Sync error: {str(e)}'
                }), 500
        
        return jsonify({
            'success': True,
            'is_selected': record.is_selected,
            'message': 'Đã trúng tuyển' if record.is_selected else 'Đã bỏ trúng tuyển'
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=app.debug, use_reloader=False)