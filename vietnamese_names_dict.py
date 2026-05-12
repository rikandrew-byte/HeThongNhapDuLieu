# -*- coding: utf-8 -*-
"""
Vietnamese Names to Chinese Dictionary
Từ điển dịch tên tiếng Việt sang tiếng Trung
Bao gồm: Họ từ 百家姓 + Tên phổ biến Việt Nam
"""

# ─── HỌ PHỔ BIẾN (từ 百家姓 - 100 họ Trung Quốc) ───
VIETNAMESE_SURNAMES = {
    'nguyễn': '阮',
    'trần': '陳',
    'hoàng': '黃',
    'huỳnh': '黃',
    'phạm': '范',
    'võ': '武',
    'vũ': '武',
    'đặng': '鄧',
    'bùi': '裴',
    'đỗ': '杜',
    'hồ': '胡',
    'ngô': '吳',
    'dương': '楊',
    'lý': '李',
    'vương': '王',
    'lê': '黎',
    'lương': '梁',
    'lưu': '劉',
    'trương': '張',
    'phan': '潘',
    'tạ': '謝',
    'trịnh': '鄭',
    'đoàn': '段',
    'mai': '梅',
    'tô': '蘇',
    'đinh': '丁',
    'đào': '陶',
    'lâm': '林',
    'phùng': '馮',
    'hà': '何',
}

# ─── TÊN PHỔ BIẾN VIỆT NAM (Dùng cho cả tên đệm và tên chính) ───
VIETNAMESE_COMMON_NAMES = {
    # Tên nam và tên đệm phổ biến
    'văn': '文',
    'hữu': ' hữu',
    'đình': '庭',
    'quốc': '國',
    'gia': '嘉',
    'minh': '明',
    'anh': '英',
    'tuấn': '俊',
    'hùng': '雄',
    'mạnh': '強',
    'kiên': '堅',
    'hải': '海',
    'sơn': '山',
    'thắng': '勝',
    'chiến': '戰',
    'quân': '軍',
    'dũng': '勇',
    'trung': '忠',
    'thành': '成',
    'vinh': '榮',
    'phúc': '福',
    'lộc': '祿',
    'thọ': '壽',
    'khang': '康',
    'an': '安',
    'bình': '平',
    'hoà': '和',
    'hòa': '和',
    'đức': '德',
    'tâm': '心',
    'nghĩa': '義',
    'hiếu': '孝',
    'nhân': '仁',
    'tài': '才',
    'quang': '光',
    'huy': '輝',
    'nam': '南',
    'việt': '越',
    'hoàng': '黃',
    'tôn': '尊',
    'trọng': '仲',
    'khoa': '科',
    'khải': '啟',
    'duy': '維',
    'long': '龍',
    'phong': '風',
    'thái': '泰',
    'thịnh': '盛',
    'tiến': '進',
    'bảo': '寶',
    'cường': '強',
    'loan': '鸞',
    'ly': '璃',
    'my': '美',
    'hằng': '姮',
    'linh': '玲',
    'chi': '芝',
    
    # Tên nữ phổ biến
    'thị': '氏',
    'diệu': '妙',
    'quỳnh': '瓊',
    'lan': '蘭',
    'hồng': '紅',
    'mai': '梅',
    'đào': '桃',
    'cúc': '菊',
    'trúc': '竹',
    'tùng': '松',
    'bích': '碧',
    'ngọc': '玉',
    'tuyết': '雪',
    'hương': '香',
    'hoa': '花',
    'liên': '蓮',
    'ánh': '光',
    'oanh': '鶯',
    'yến': '燕',
    'huyền': '玄',
    'trang': '庄',
    'phương': '芳',
    'thảo': '草',
    'thủy': '水',
    'thu': '秋',
    'xuân': '春',
    'hạ': '夏',
    'đông': '冬',
    'vân': '雲',
    'vy': '薇',
    'nhi': '兒',
    'nhung': '絨',
    'linh': '玲',
    'trinh': '貞',
    'tiên': '仙',
    'tú': '秀',
    'nguyệt': '月',
    'chi': '芝',
    'diệp': '葉',
    'hạnh': '幸',
    'hiền': '賢',
}

# ─── TÊN ĐẦY ĐỦ ĐẶC BIỆT ───
VIETNAMESE_FULL_NAMES = {
    'nguyễn văn a': '阮文甲',
    'trần văn b': '陳文乙',
}

def get_vietnamese_name_in_chinese(full_name: str) -> str:
    """
    Dịch tên tiếng Việt sang tiếng Trung bằng từ điển nội bộ
    """
    if not full_name:
        return full_name
    
    full_name_lower = full_name.strip().lower()
    
    # 1. Kiểm tra bảng tên đầy đủ trước
    if full_name_lower in VIETNAMESE_FULL_NAMES:
        return VIETNAMESE_FULL_NAMES[full_name_lower]
    
    # 2. Tách họ và tên
    parts = full_name_lower.split()
    if not parts:
        return full_name
        
    if len(parts) == 1:
        # Chỉ có 1 từ - ưu tiên họ trước, rồi đến tên
        if parts[0] in VIETNAMESE_SURNAMES:
            return VIETNAMESE_SURNAMES[parts[0]]
        elif parts[0] in VIETNAMESE_COMMON_NAMES:
            return VIETNAMESE_COMMON_NAMES[parts[0]]
        return full_name
    
    # 3. Dịch họ (phần đầu)
    surname = parts[0]
    surname_chinese = VIETNAMESE_SURNAMES.get(surname, surname)
    
    # 4. Dịch các phần còn lại (Tên đệm + Tên chính)
    name_parts = parts[1:]
    name_chinese_parts = []
    for part in name_parts:
        if part in VIETNAMESE_COMMON_NAMES:
            name_chinese_parts.append(VIETNAMESE_COMMON_NAMES[part])
        elif part in VIETNAMESE_SURNAMES: # Một số tên đệm là họ (ví dụ: Nguyễn Hoàng)
            name_chinese_parts.append(VIETNAMESE_SURNAMES[part])
        else:
            name_chinese_parts.append(part)
    
    # 5. Ghép lại
    result = surname_chinese + ''.join(name_chinese_parts)
    return result
