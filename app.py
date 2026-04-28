# -*- coding: utf-8 -*-
"""
Document Automation System (DAS) - Flask Backend V3.0
Nhập liệu Tiếng Việt → Xuất file Word với nội dung Tiếng Trung Phồn Thể
"""
import os, uuid, re, unicodedata, json, base64
from datetime import date, datetime
from flask import Flask, request, jsonify, send_file, render_template, Response
from flask_cors import CORS
from jinja2 import Template
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm, Pt
from deep_translator import GoogleTranslator
from dotenv import load_dotenv
from urllib.parse import quote
from unicodedata import normalize


load_dotenv()

app = Flask(__name__, static_folder='static', static_url_path='')
app.debug = True  # Lệnh cưỡng chế bật Debug
CORS(app, resources={r"/*": {"origins": ["https://cv.fct.vn", "http://127.0.0.1:5000", "http://localhost:5000"]}})
from flask_basicauth import BasicAuth

app.config['BASIC_AUTH_USERNAME'] = 'fctvt'  # Tên đăng nhập bạn chọn
app.config['BASIC_AUTH_PASSWORD'] = '1503'   # Mật khẩu bạn chọn
app.config['BASIC_AUTH_FORCE_PROMPT'] = True

basic_auth = BasicAuth(app)

# --- MIDDLEWARE BẢO MẬT TÙY BIẾN ---
def auth_required(f):
    """
    Nếu chạy trên Render (Cloud) -> Bắt buộc đăng nhập.
    Nếu chạy Local -> Tạm thời bỏ qua Guard để MASTER thử nghiệm.
    """
    if os.environ.get('RENDER'):
        return basic_auth.required(f)
    return f

from flask_sqlalchemy import SQLAlchemy

# --- CẤU HÌNH DATABASE (Linh hoạt Local & Cloud) ---
db_url = os.environ.get('DATABASE_URL')
if db_url:
    # Xử lý chuẩn hóa URL cho SQLAlchemy nếu dùng PostgreSQL
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
else:
    # Nếu chạy local (không có DATABASE_URL), tự động dùng SQLite
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

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
TMPL_DIR   = os.path.join(BASE_DIR, 'templates')
OUT_DIR    = os.path.join(BASE_DIR, 'output')
UPL_DIR    = os.path.join(BASE_DIR, 'uploads')
for d in (TMPL_DIR, OUT_DIR, UPL_DIR):
    os.makedirs(d, exist_ok=True)

from jinja2 import ChoiceLoader, FileSystemLoader
app.jinja_loader = ChoiceLoader([
    FileSystemLoader(TMPL_DIR),
    FileSystemLoader(BASE_DIR)
])

# Khởi tạo bảng trong Database nếu chưa có
with app.app_context():
    db.create_all()

# ─── Bảng dịch cố định Việt → Trung Phồn Thể ────────────────────────────────
FIXED_TRANS = {
    # Hôn nhân
    'độc thân': '未婚', 'doc than': '未婚',
    'đã kết hôn': '已婚', 'da ket hon': '已婚', 'có gia đình': '已婚',
    'ly hôn': '離婚', 'ly hon': '離婚',
    'góa': '喪偶', 'goa': '喪偶',
    # Học vấn
    'tiểu học': '國小', 'tieu hoc': '國小',
    'thcs': '國中', 'trung học cơ sở': '國中',
    'thpt': '高中', 'trung học phổ thông': '高中',
    'trung cấp': '高職', 'trung cap': '高職',
    'cao đẳng': '專科', 'cao dang': '專科',
    'đại học': '大學', 'dai hoc': '大學',
    'thạc sĩ': '碩士', 'thac si': '碩士',
    'tiến sĩ': '博士', 'tien si': '博士',
    # Quốc gia
    'việt nam': '越南', 'viet nam': '越南',
    'đài loan': '台灣',
    'nhật bản': '日本',
    'hàn quốc': '韓國',
    'malaysia': '馬來西亞',
    'macau': '澳門',
    'thái lan': '泰國',
    'châu âu': '歐洲',
    'nga': '俄羅斯',
}

TRANSLATION_MAP = {
    # Hôn nhân
    'độc thân': '未婚', 'doc than': '未婚',
    'đã kết hôn': '已婚', 'da ket hon': '已婚', 'có gia đình': '已婚',
    'ly hôn': '離婚', 'ly hon': '離婚',
    'góa': '喪偶', 'goa': '喪偶',
    # Học vấn
    'tiểu học': '國小', 'tieu hoc': '國小',
    'thcs': '國中', 'trung học cơ sở': '國中',
    'thpt': '高中', 'trung học phổ thông': '高中',
    'trung cấp': '高職', 'trung cap': '高職',
    'cao đẳng': '專科', 'cao dang': '專科',
    'đại học': '大學', 'dai hoc': '大學',
    'thạc sĩ': '碩士', 'thac si': '碩士',
    'tiến sĩ': '博士', 'tien si': '博士',
}

# BỘ TỪ ĐIỂN CÁC MỤC LƯU Ý ĐẶC BIỆT (MỤC TÔ VÀNG)
YELLOW_ALERTS_MAP = {
    "f05": "骨折",                 # Gãy xương
    "f06": "手汗",                 # Mồ hôi tay
    "f11": "脊椎受傷",             # Chấn thương cột sống
    "f13": "伏地挺身 10~30 下",    # Chống đẩy 10~30 cái
    "f14": "搬重 20~40 kg",        # Bê vác 20~40 kg
    "f18": "肝炎",                 # Viêm gan
    "f19": "斷指",                 # Cụt đốt ngón tay
    "f20": "哮喘",                 # Hen suyễn
    "f21": "伏地挺身 50 下以上",   # Chống đẩy >50 cái
    "f22": "搬重 50kg 以上",       # Bê vác >50kg
}

def translate_fixed(text: str) -> str:
    """Dịch các từ cố định theo bảng tra cứu"""
    if not text:
        return text
    key = text.strip().lower()
    return FIXED_TRANS.get(key, text)

def translate_free(text: str) -> str:
    """Dịch văn bản tự do sang Tiếng Trung Phồn Thể"""
    if not text or not text.strip():
        return text
    try:
        result = GoogleTranslator(source='vi', target='zh-TW').translate(text)
        return result if result else text
    except Exception:
        return text

def calc_age(dob_str: str) -> str:
    """Tính tuổi từ chuỗi ngày sinh YYYY-MM-DD"""
    if not dob_str:
        return ''
    try:
        p = dob_str.split('-')
        dob = date(int(p[0]), int(p[1]), int(p[2]))
        today = date.today()
        age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
        return str(age)
    except Exception:
        return ''

def sanitize_filename_master(text):
    # 1. Chuyển tiếng Việt có dấu thành không dấu
    text = normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
    # 2. Viết hoa mỗi chữ cái đầu (Title Case)
    text = text.title() 
    # 3. Xóa ký tự đặc biệt, thay khoảng trắng bằng dấu gạch dưới
    text = re.sub(r'[^\w\s-]', '', text).strip().replace(' ', '_')
    return text

def to_ascii(title: str) -> str:
    """Chuyển tiếng Việt có dấu sang không dấu, viết hoa mỗi chữ cái đầu"""
    if not title:
        return ''
    # Handle 'đ' and 'Đ' specifically as they are not handled by NFD normalization
    title = title.replace('đ', 'd').replace('Đ', 'D')
    nfd = unicodedata.normalize('NFD', title)
    ascii_str = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
    return ' '.join(word.capitalize() for word in ascii_str.split())

def fmt_date(d: str) -> str:
    """Chuyển YYYY-MM-DD → DD/MM/YYYY"""
    if d and '-' in d:
        p = d.split('-')
        if len(p) == 3:
            return f'{p[2]}/{p[1]}/{p[0]}'
    return d or ''

def chk(val):
    """Checkbox: True → ☑, False → □ (Định dạng MS Gothic size 11pt)"""
    if val in (True, 'true', '1', 1, 'yes', 'on', 'checked'):
        return RichText("☑", font='MS Gothic', size=22)
    return RichText("□", font='MS Gothic', size=22)

def get_base64_image(file_path):
    """Chuyển đổi file ảnh sang chuỗi Base64 để nhúng trực tiếp vào HTML"""
    if not file_path or not os.path.exists(file_path):
        return ""
    try:
        with open(file_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
            ext = os.path.splitext(file_path)[1][1:].lower()
            if ext == 'jpg': ext = 'jpeg'
            return f"data:image/{ext};base64,{encoded_string}"
    except Exception as e:
        print(f"Lỗi Base64: {e}")
        return ""

# --- EMERALD CORE V6.15: HTML EXPORT LOGIC ---

SKILL_MAPPING = {
    'f23': 'Hàn điện / 電焊', 'f24': 'Hàn Argon / 氬焊',
    'f25': 'Hàn CO2 / 氣焊', 'f26': 'Tig/Mig',
    'f27': 'Đúc / 鑄造', 'f28': 'Dệt / 紡織',
    'f29': 'May / 縫紉', 'f30': 'Lái xe nâng / 堆高機',
    'f31': 'Tiện / 車床', 'f32': 'Phay / 銑床',
    'f33': 'Bào / 刨床', 'f34': 'CNC',
    'f35': 'Đột dập / 沖床', 'f36': 'In ấn / 印刷',
    'f37': 'Thợ mộc / 木工', 'f38': 'Lái xe tải/khách / 卡車/客司機',
    'f39': 'Nhựa / 塑膠', 'f40': 'Xây dựng / 營造',
    'f41': 'Sửa chữa máy / 機械維修', 'f42': 'Điều dưỡng / 護理工',
    'f43': 'Giúp việc / 幫傭', 'f44': 'Xe cẩu / 吊車',
    'f45': 'Cẩu trục / 天車', 'f46': 'Máy xúc / 挖土機'
}

def prepare_html_data(raw_data: dict) -> dict:
    """Chuẩn hóa dữ liệu cho template HTML V6.24 (Cloud Optimized) - Không dịch tự động để tránh lỗi UTF-8"""
    # Tạo bản sao dữ liệu để xử lý
    data = {}
    fields = [
        'Maso', 'Hoten', 'TentiengTrung', 'Ngaysinh', 'Tuoi', 'Chieucao', 'Cannang', 
        'Lienhe', 'Noio', 'HotenBo', 'TB', 'HotenMe', 'TM', 'VoChong', 'VC', 
        'Socon', 'Anhchiem', 'Xepthu', 'f48', 'N1', 'N2', 'N3', 'ndcv1', 'ndcv2', 'ndcv3', 
        'loi_binh_1', 'Honnhan', 'Hocvan', 'QG1', 'QG2', 'QG3'
    ]
    for f in fields:
        val = str(raw_data.get(f, '')).strip()
        if f in ['Honnhan', 'Hocvan']:
            data[f] = TRANSLATION_MAP.get(val.lower(), val)
        elif f in ['QG1', 'QG2', 'QG3']:
            data[f] = translate_fixed(val)
        else:
            data[f] = val

    # Dịch tự do các trường cần thiết sang tiếng Trung Phồn Thể
    for f in ['Noio', 'ndcv1', 'ndcv2', 'ndcv3', 'loi_binh_1', 'N1', 'N2', 'N3']:
        data[f] = translate_free(data.get(f, ''))

    # Gom các trường tô vàng vào loi_binh_1 (HTML Ver 6.30)
    yellow_alerts = []
    for i in range(1, 23):
        key = f'f{i:02d}'
        if raw_data.get(key) in (True, 'true', '1', 1, 'yes', 'on', 'checked') and key in YELLOW_ALERTS_MAP:
            yellow_alerts.append(YELLOW_ALERTS_MAP[key])
            
    if yellow_alerts:
        alert_str = "、".join(yellow_alerts)
        if data.get('loi_binh_1'):
            data['loi_binh_1'] = data['loi_binh_1'] + "、" + alert_str
        else:
            data['loi_binh_1'] = alert_str

    # Kỹ năng
    skills_html = []
    for key, name in SKILL_MAPPING.items():
        if raw_data.get(key) in (True, 'true', '1', 1, 'yes', 'on', 'checked'):
            tag = f'<span class="px-2 py-1 bg-emerald-600 text-white rounded text-[9px] font-bold uppercase shadow-sm">{name}</span>'
            skills_html.append(tag)
    data['KyNangList_HTML'] = "".join(skills_html)

    # Các thông số đặc biệt
    data['f01'] = "右手" if raw_data.get('f01') in (True, 'true', 1) else ""
    data['f07'] = "Trái / 左手" if raw_data.get('f07') in (True, 'true', 1) else ""
    
    data['f02'] = "Hỏng mắt phải / 右眼受損" if raw_data.get('f02') in (True, 'true', 1) else ""
    data['f08'] = "Hỏng mắt trái / 左眼受損" if raw_data.get('f08') in (True, 'true', 1) else ""
    data['f16'] = "Loạn thị / 散光" if raw_data.get('f16') in (True, 'true', 1) else ""
    data['f17'] = "Mù màu / 色盲" if raw_data.get('f17') in (True, 'true', 1) else ""
    data['f15'] = "Cận / 近視" if raw_data.get('f15') in (True, 'true', 1) else ""
    
    if not any([data['f02'], data['f08'], data['f16'], data['f17'], data['f15']]):
        data['f15'] = "正常"

    data['f12'] = "Có / 有" if raw_data.get('f12') in (True, 'true', 1) else "Không / 無"

    # Chuyển ảnh đại diện sang Base64
    photo_path = raw_data.get('photo', '')
    if photo_path and isinstance(photo_path, str) and photo_path.startswith('data:image/'):
        data['photo_base64'] = photo_path
    elif photo_path and isinstance(photo_path, str) and os.path.exists(photo_path):
        data['photo_base64'] = get_base64_image(photo_path)
    else:
        data['photo_base64'] = ""
            
    return data

def generate_html_resume(form_data: dict, template_name='fct_template_v6.18.html') -> str:
    """Render HTML Resume (Bản 6.24: Nhúng Base64 toàn bộ ảnh)"""
    processed_data = prepare_html_data(form_data)
    
    # Nhúng Logo và Background vào context
    processed_data['logo_base64'] = get_base64_image(os.path.join(BASE_DIR, 'static', 'logo.png'))
    processed_data['bg_base64'] = get_base64_image(os.path.join(BASE_DIR, 'static', 'fct_bg.png'))
    
    TEMPLATE_PATH = os.path.join(BASE_DIR, 'templates', template_name)
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        template = Template(f.read())
        
    return template.render(processed_data)

def export_resume(form_data, export_type='html'):
    """Xuất hồ sơ sang định dạng HTML Minified (Bản 6.20 Emerald Forever)"""
    # 1. Render HTML từ template v6.20 (vẫn dùng chung tên file v6.18 để tránh đổi code nhiều nơi)
    rendered_html = generate_html_resume(form_data, 'fct_template_v6.18.html')
    
    # 2. Bảo vệ mã nguồn (Minification)
    minified_html = rendered_html
    
    # 3. QUY TẮC ĐẶT TÊN FILE (Filename Convention)
    ma_so = form_data.get('Maso', '').strip()
    ho_ten = form_data.get('Hoten', '').strip()
    clean_name = sanitize_filename_master(ho_ten)
    
    if ma_so and clean_name:
        base_name = f"{ma_so}_{clean_name}"
    elif ma_so:
        base_name = f"{ma_so}"
    elif clean_name:
        base_name = f"{clean_name}"
    else:
        base_name = f"resume_{uuid.uuid4().hex[:8]}"
        
    file_name = f"{base_name}.html"
    return minified_html, file_name

# ─── Chuẩn bị dữ liệu cho template ──────────────────────────────────────────
def prepare_data(raw: dict) -> dict:
    context = {}

    # 1. Trường văn bản và gia đình (Khớp 1:1 với name trong HTML)
    fields = [
        'Maso', 'Hoten', 'TentiengTrung', 'Ngaysinh', 'Tuoi', 'Chieucao', 'Cannang', 
        'Lienhe', 'Noio', 'HotenBo', 'TB', 'HotenMe', 'TM', 'VoChong', 'VC', 
        'Socon', 'Anhchiem', 'Xepthu', 'f48', 'N1', 'N2', 'N3', 'ndcv1', 'ndcv2', 'ndcv3', 'loi_binh_1'
    ]
    for f in fields:
        val = raw.get(f, '')
        context[f] = str(val).replace('\n', ' ').strip()

    # 2. Ngày sinh và tuổi
    context['Ngaysinh'] = fmt_date(context['Ngaysinh'])
    if raw.get('Ngaysinh') and not context['Tuoi']:
        context['Tuoi'] = calc_age(raw.get('Ngaysinh'))

    # 3. Dịch thuật
    for f in ['Honnhan', 'Hocvan', 'QG1', 'QG2', 'QG3']:
        context[f] = translate_fixed(raw.get(f, ''))
    
    for f in ['Noio', 'ndcv1', 'ndcv2', 'ndcv3', 'loi_binh_1', 'N1', 'N2', 'N3']:
        context[f] = translate_free(context.get(f, ''))

    # 4. Checkbox f01 -> f46
    for i in range(1, 47):
        key = f'f{i:02d}'
        context[key] = chk(raw.get(key, False))

    # Gom các trường tô vàng vào loi_binh_1 (Word & DB Sync)
    yellow_alerts = []
    for i in range(1, 23):
        key = f'f{i:02d}'
        if raw.get(key) in (True, 'true', '1', 1, 'yes', 'on', 'checked') and key in YELLOW_ALERTS_MAP:
            yellow_alerts.append(YELLOW_ALERTS_MAP[key])
            
    if yellow_alerts:
        alert_str = "、".join(yellow_alerts)
        if context.get('loi_binh_1'):
            context['loi_binh_1'] = context['loi_binh_1'] + "、" + alert_str
        else:
            context['loi_binh_1'] = alert_str

    context['photo'] = raw.get('photo', '')
    return context

def generate_word(form_data: dict, template_name='resume_template_chuan.docx') -> str:
    # 1. MỞ ĐÚNG FILE MẪU (Nằm cùng thư mục với app.py)
    TEMPLATE_PATH = os.path.join(BASE_DIR, template_name)
    
    pass

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f'Template không tồn tại: {TEMPLATE_PATH}')
    
    doc = DocxTemplate(TEMPLATE_PATH)
    
    # 2. XỬ LÝ ẢNH CHUYÊN SÂU
    photo_path = form_data.get('photo', '')
    if photo_path and isinstance(photo_path, str) and photo_path.startswith('data:image/'):
        try:
            import base64
            header, encoded = photo_path.split(",", 1)
            img_data = base64.b64decode(encoded)
            if not os.path.exists(UPL_DIR): os.makedirs(UPL_DIR)
            temp_path = os.path.join(UPL_DIR, f"temp_avatar_{uuid.uuid4().hex[:8]}.png")
            with open(temp_path, "wb") as fh:
                fh.write(img_data)
            form_data['photo'] = InlineImage(doc, temp_path, width=Mm(64), height=Mm(85))
        except Exception as e:
            print(f"Lỗi giải mã ảnh Base64 cho Word: {e}")
            form_data['photo'] = ""
    elif photo_path and isinstance(photo_path, str) and os.path.exists(photo_path):
        form_data['photo'] = InlineImage(doc, photo_path, width=Mm(64), height=Mm(85))
    else:
        form_data['photo'] = ""
    
    # 3. ĐỔ DỮ LIỆU VÀO TEMPLATE
    doc.render(form_data)
    
    # 4. LƯU FILE THEO QUY TẮC MASTER YÊU CẦU
    ma_so = form_data.get('Maso', '').strip()
    ho_ten = form_data.get('Hoten', '').strip()
    ten_khong_dau = to_ascii(ho_ten)
    if ma_so and ten_khong_dau:
        fname = f"{ma_so} {ten_khong_dau}.docx"
    elif ma_so:
        fname = f"{ma_so}.docx"
    elif ten_khong_dau:
        fname = f"{ten_khong_dau}.docx"
    else:
        fname = f"resume_{uuid.uuid4().hex[:8]}.docx"
    
    out = os.path.join(OUT_DIR, fname)
    doc.save(out)
    return out

# ─── API Routes ──────────────────────────────────────────────────────────────
@app.route('/')
def user_form():
    return render_template('user_form.html')

@app.route('/fct-1503')
@auth_required
def index():
    return render_template('index.html')

@app.route('/api/health')
def health():
    return jsonify({'ok': True, 'msg': 'DAS V3.0 running'})

@app.route('/api/generate', methods=['POST'])
@auth_required
def api_generate():
    try:
        # Handle both FormData and JSON
        if request.content_type and 'multipart/form-data' in request.content_type:
            data = json.loads(request.form.get('data', '{}'))
            photo_file = request.files.get('photo')
            if photo_file:
                file_bytes = photo_file.read()
                encoded_string = base64.b64encode(file_bytes).decode('utf-8')
                ext = os.path.splitext(photo_file.filename)[1][1:].lower()
                if not ext: ext = 'png'
                if ext == 'jpg': ext = 'jpeg'
                data['photo'] = f"data:image/{ext};base64,{encoded_string}"
        else:
            data = request.get_json() or {}
        
        # 1. Render HTML từ template v6.18
        form_data = prepare_data(data)
        html_content, fn = export_resume(form_data, 'html')
        
        # 2. Bỏ qua việc lưu file vật lý để không chiếm ổ cứng Render
        pass

        # 3. Lưu vào Database Lịch sử trước khi trả về file
        try:
            new_record = FormHistory(
                ma_so=form_data.get('Maso', ''),
                ho_ten=form_data.get('Hoten', ''),
                ten_file=fn,
                data_json=json.dumps(data, ensure_ascii=False)
            )
            db.session.add(new_record)
            db.session.commit()
        except Exception:
            db.session.rollback()

        # 4. Trả về Response để trình duyệt tải file
        return Response(
            html_content,
            mimetype="text/html",
            headers={
                "Content-Disposition": f'attachment; filename="{fn}"',
                "Content-Type": "text/html; charset=utf-8"
            }
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/api/submit-only', methods=['POST'])
@auth_required
def api_submit_only():
    try:
        # Handle both FormData and JSON
        if request.content_type and 'multipart/form-data' in request.content_type:
            data = json.loads(request.form.get('data', '{}'))
            photo_file = request.files.get('photo')
            if photo_file:
                file_bytes = photo_file.read()
                encoded_string = base64.b64encode(file_bytes).decode('utf-8')
                ext = os.path.splitext(photo_file.filename)[1][1:].lower()
                if not ext: ext = 'png'
                if ext == 'jpg': ext = 'jpeg'
                data['photo'] = f"data:image/{ext};base64,{encoded_string}"
        else:
            data = request.get_json() or {}
        
        form_data = prepare_data(data)
        
        # --- LƯU LỊCH SỬ VÀO DATABASE (KHÔNG XUẤT WORD) ---
        try:
            new_record = FormHistory(
                ma_so=form_data.get('Maso', '') or 'CHO_DUYET',
                ho_ten=form_data.get('Hoten', ''),
                ten_file='',
                data_json=json.dumps(data, ensure_ascii=False)
            )
            db.session.add(new_record)
            db.session.commit()
        except Exception:
            db.session.rollback()
            
        return jsonify({
            'success': True, 
            'id': new_record.id,
            'ma_so': new_record.ma_so,
            'msg': 'Đã nộp form thành công (chờ duyệt).'
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500

# --- API DỊCH TỰ ĐỘNG CHO GIAO DIỆN ---
@app.route('/api/translate', methods=['POST'])
@auth_required
def api_translate():
    try:
        data = request.get_json() or {}
        text = data.get('text', '')
        return jsonify({'success': True, 'translated': translate_free(text)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/download/<maso>', methods=['GET'])
@auth_required
def download_history(maso):
    try:
        # 1. Gọi Database lấy data theo maso
        record = FormHistory.query.filter_by(ma_so=maso).order_by(FormHistory.ngay_tao.desc()).first()
        if not record:
            return jsonify({"error": "Không tìm thấy hồ sơ với mã số này"}), 404
            
        data = json.loads(record.data_json)
        
        # 2. Render lại HTML từ data lấy được (Đảm bảo đã chạy qua bộ lọc dịch tiếng Trung)
        html_content = generate_html_resume(data, 'fct_template_v6.18.html')
        
        # 3. Tạo tên file chuẩn (Không dấu, viết hoa chữ cái đầu)
        clean_name = sanitize_filename_master(record.ho_ten)
        filename = f"{maso}_{clean_name}.html"
        
        # 4. Trả file về trình duyệt
        import io
        return send_file(
            io.BytesIO(html_content.encode('utf-8')),
            mimetype='text/html',
            as_attachment=True,
            download_name=filename
        )
    except Exception:
        return jsonify({"error": "Không thể tải file, vui lòng kiểm tra lại data"}), 400

@app.route('/resume-<int:record_id>.html')
@auth_required
def api_preview(record_id):
    try:
        record = FormHistory.query.get(record_id)
        if not record:
            return "Không tìm thấy hồ sơ", 404
        
        data = json.loads(record.data_json)
        html_content = generate_html_resume(data)
        return Response(html_content, mimetype="text/html", headers={"Content-Type": "text/html; charset=utf-8"})
    except Exception as e:
        import traceback
        return f"Lỗi render: {str(e)}<pre>{traceback.format_exc()}</pre>", 500

@app.route('/api/view-photo')
@auth_required
def api_view_photo():
    """Route để hiển thị ảnh từ path tuyệt đối trong thẻ <img>"""
    path = request.args.get('path')
    if path and os.path.exists(path):
        return send_file(path)
    return "Not Found", 404

# --- API LẤY DANH SÁCH LỊCH SỬ ---
@app.route('/api/history', methods=['GET'])
@auth_required
def api_history():
    try:
        records = FormHistory.query.order_by(FormHistory.ngay_tao.desc()).limit(100).all()
        data = [{
            'id': r.id,
            'ma_so': r.ma_so,
            'ho_ten': r.ho_ten,
            'ten_file': r.ten_file,
            'data_json': json.loads(r.data_json) if r.data_json else None,
            'ngay_tao': r.ngay_tao.strftime("%d/%m/%Y %H:%M:%S")
        } for r in records]
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# --- API XÓA LỊCH SỬ ---
@app.route('/api/history/<int:record_id>', methods=['DELETE'])
@auth_required
def api_delete_history(record_id):
    try:
        record = FormHistory.query.get(record_id)
        if not record:
            return jsonify({'success': False, 'error': 'Không tìm thấy bản ghi'}), 404
        
        # Xóa file Word vật lý trong thư mục output (nếu còn tồn tại)
        if record.ten_file:
            file_path = os.path.join(OUT_DIR, record.ten_file)
            if os.path.exists(file_path) and os.path.isfile(file_path):
                os.remove(file_path)
            
        db.session.delete(record)
        db.session.commit()
        return jsonify({'success': True, 'msg': 'Xóa thành công'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    import os
    # Lấy cổng do Render cấp, nếu không có thì mặc định 5000
    port = int(os.environ.get("PORT", 5000))
    # Chạy trên host 0.0.0.0 để có thể truy cập từ Internet
    app.run(host='0.0.0.0', port=port, debug=True, use_reloader=False)