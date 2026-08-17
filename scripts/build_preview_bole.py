import base64, os
from jinja2 import Template

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_b64(rel_path):
    full_path = os.path.join(BASE_DIR, rel_path)
    if not os.path.exists(full_path): return ''
    with open(full_path, 'rb') as f:
        ext = os.path.splitext(full_path)[1][1:].lower()
        if ext == 'jpg': ext = 'jpeg'
        return f"data:image/{ext};base64,{base64.b64encode(f.read()).decode('utf-8')}"

# Đọc mẫu fct_template_bole.html
with open(os.path.join(BASE_DIR, 'templates', 'fct_template_bole.html'), 'r', encoding='utf-8') as f:
    tpl_str = f.read()

tpl = Template(tpl_str)
bg_b64 = get_b64(os.path.join('static', 'banner_bole_qianlima.jpg'))
logo_b64 = get_b64(os.path.join('static', 'logo.png'))

sample_data = {
    'Maso': 'FD4128',
    'clean_name_pdf': 'TRAN_PHUONG_THUY',
    'Hoten': 'TRẦN PHƯƠNG THUỲ',
    'TentiengTrung': '陳 芳 垂',
    'Ngaysinh': '12/09/2007',
    'Tuoi': '18',
    'Chieucao': '162',
    'Cannang': '48',
    'Honnhan': '未婚',
    'Socon': '0',
    'Hocvan': '高中',
    'Noio': '太原',
    'ThiLuc': '正常',
    'TayThuan': '右手',
    'f12': '無',
    'HutRuou': '不抽菸 / 不喝酒',
    'f48': 'Admin - HR Team',
    'loi_binh_1': 'Chăm chỉ, nhanh nhẹn, có sức khỏe tốt, mong muốn làm việc lâu dài.',
    'KyNangList_HTML': '<span class="skill-tag">Công nhân điện tử / 電子工</span>',
    'photo_base64': 'https://images.unsplash.com/photo-1544005313-94ddf0286df2?auto=format&fit=crop&w=400&q=80',
    'qr_line_base64': '',
    'logo_base64': logo_b64,
    'bg_base64': bg_b64,
    'raw_data_json': '{}',
    'document_images': []
}

rendered = tpl.render(sample_data)

preview_file_path = os.path.join(BASE_DIR, 'preview_bole_ver2026.html')
with open(preview_file_path, 'w', encoding='utf-8') as f:
    f.write(rendered)

print(f"SUCCESS: Updated {preview_file_path}")
