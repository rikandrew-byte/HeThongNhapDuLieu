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
import boto3
from PIL import Image
from vietnamese_names_dict import get_vietnamese_name_in_chinese

load_dotenv()

app = Flask(__name__, static_folder='static', static_url_path='')
app.debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
CORS(app, resources={r"/*": {"origins": ["https://cv.fct.vn", "http://127.0.0.1:5000", "http://localhost:5000"]}})

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

with app.app_context():
    try:
        db.create_all()
        # Vá lỗi thiếu cột is_selected (Chạy trực tiếp cho cả SQLite/Postgres/Xata)
        try:
            from sqlalchemy import text
            # Thử thêm cột, nếu đã có rồi thì SQL sẽ báo lỗi và ta sẽ rollback/bỏ qua
            db.session.execute(text("ALTER TABLE form_history ADD COLUMN is_selected BOOLEAN DEFAULT FALSE"))
            db.session.commit()
            print("✅ Database migration: Column 'is_selected' added.")
        except Exception as e:
            db.session.rollback()
            # Lỗi thường là do cột đã tồn tại, ta im lặng bỏ qua
            if "already exists" not in str(e).lower() and "duplicate column" not in str(e).lower():
                print(f"⚠️ Migration notice: {e}")
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
    }
    
    # Nếu text là một từ đơn nằm trong danh sách bảo vệ, trả về luôn
    if text.strip().lower() in protected_terms:
        return protected_terms[text.strip().lower()]

    # 3. Dùng Google Translate cho đoạn văn
    try:
        result = GoogleTranslator(source='vi', target='zh-TW').translate(processed_text.strip())
        # Nếu Google dịch nhầm "may" -> "可能" (có lẽ) trong ngữ cảnh công việc, ta sửa lại
        if '可能' in result and 'may' in text.lower():
            # Đây là một heuristic đơn giản, có thể cải thiện thêm
            result = result.replace('可能', '縫紉')
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

# --- R2 UPLOAD HELPER ---
def init_r2_client():
    """Khởi tạo R2 client (Cloudflare)"""
    try:
        r2_client = boto3.client(
            's3',
            endpoint_url=f"https://{os.environ.get('R2_ACCOUNT_ID')}.r2.cloudflarestorage.com",
            aws_access_key_id=os.environ.get('R2_ACCESS_KEY'),
            aws_secret_access_key=os.environ.get('R2_SECRET_KEY'),
            region_name='auto'
        )
        return r2_client
    except Exception as e:
        print(f"⚠️ R2 client init failed: {str(e)}")
        return None

r2_client = init_r2_client()

def upload_to_r2(file_data, filename, folder='documents'):
    """Upload ảnh lên R2 (tự động compress & resize)"""
    if not r2_client:
        return None
    
    try:
        # Resize ảnh để tiết kiệm storage
        img = Image.open(io.BytesIO(file_data))
        
        # Giới hạn kích thước
        max_size = 1200
        if img.width > max_size or img.height > max_size:
            img.thumbnail((max_size, max_size), Image.LANCZOS)
        
        # Convert RGBA → RGB nếu cần
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        
        # Compress ảnh
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=75, optimize=True)
        buf.seek(0)
        
        # Tạo key duy nhất
        key = f"{folder}/{uuid.uuid4()}_{filename}"
        
        # Upload lên R2
        r2_client.upload_fileobj(
            buf,
            os.environ.get('R2_BUCKET_NAME'),
            key,
            ExtraArgs={'ContentType': 'image/jpeg'}
        )
        
        # Trả về public URL
        public_url = f"https://{os.environ.get('R2_PUBLIC_URL')}/{key}"
        print(f"✅ Uploaded to R2: {public_url}")
        return public_url
    except Exception as e:
        print(f"❌ R2 upload error: {str(e)}")
        traceback.print_exc()
        return None

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
            skills_html.append(f'<span class="px-2 py-1 bg-transparent border border-emerald-600 text-emerald-800 rounded text-[13px] font-bold uppercase">{name}</span>')
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
        'var _t=function(){if(window.outerWidth-window.innerWidth>160||window.outerHeight-window.innerHeight>160){document.body.innerHTML="";}};'
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
    raw_for_embed = {k: v for k, v in form_data.items() if k not in ('photo', 'qr_line')}
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
    return clean

# --- API ROUTES ---
@app.route('/')
def user_form(): return render_template('user_form.html')

@app.route('/fct-1503')
@auth_required
def index(): return render_template('index.html')

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
            for key in ('photo', 'qr_line'):
                if not data.get(key) and old_data.get(key): data[key] = old_data[key]
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

        # 🆕 Upload ảnh tài liệu lên R2
        document_urls = []
        if 'documents' in request.files:
            files = request.files.getlist('documents')
            for file in files:
                if file and file.filename:
                    try:
                        url = upload_to_r2(file.read(), file.filename, f'documents/{ma_so}')
                        if url:
                            document_urls.append(url)
                    except Exception as e:
                        print(f"⚠️ Failed to upload {file.filename}: {str(e)}")
            
            # Lưu URLs vào database
            if document_urls:
                data_to_save = json.loads(record.data_json) if record.data_json else {}
                data_to_save['document_urls'] = document_urls
                record.data_json = json.dumps(data_to_save, ensure_ascii=False)
                db.session.commit()

        return jsonify({'success': True, 'id': record.id, 'ma_so': record.ma_so, 'msg': msg, 'documents_uploaded': len(document_urls)})
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
        records = FormHistory.query.order_by(FormHistory.ngay_tao.desc()).limit(100).all()
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
        for r in records: db.session.delete(r)
        db.session.commit()
        return jsonify({'success': True, 'deleted': len(records)})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/history/<int:record_id>', methods=['DELETE'])
@auth_required
def api_delete_history(record_id):
    try:
        record = FormHistory.query.get(record_id)
        if record:
            db.session.delete(record)
            db.session.commit()
        return jsonify({'success': True})
    except: return jsonify({'success': False}), 500

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
        # Modify CSS for bulk print
        base_html = base_html.replace('height: 297mm !important; overflow: hidden !important;', 'height: auto !important; overflow: visible !important;', 1)
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