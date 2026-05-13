# -*- coding: utf-8 -*-
"""
Vietnamese Names to Chinese Dictionary
Từ điển dịch tên tiếng Việt sang tiếng Trung
Bao gồm: Họ từ 百家姓 + Tên phổ biến Việt Nam
Logic: Chỉ trả về kết quả khi TẤT CẢ các phần đều dịch được.
       Nếu có phần nào không dịch được → trả về None → fallback Google Translate.
"""

# ─── HỌ PHỔ BIẾN (từ 百家姓 + họ Việt phổ biến) ───
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
    'tống': '宋',
    'chu': '朱',
    'mã': '馬',
    'cao': '高',
    'khương': '姜',
    'tần': '秦',
    'hạ': '夏',
    'tấn': '晉',
    'thẩm': '沈',
    'hàn': '韓',
    'dư': '余',
    'từ': '徐',
    'tăng': '曾',
    'tiêu': '蕭',
    'lạc': '駱',
    'thang': '湯',
    'tiền': '錢',
    'tôn': '孫',
    'lý': '李',
    'chu': '周',
    'ngô': '吳',
    'trịnh': '鄭',
    'vương': '王',
    'phùng': '馮',
    'trần': '陳',
    'chử': '褚',
    'vệ': '衛',
    'tưởng': '蔣',
    'thẩm': '沈',
    'hàn': '韓',
    'dương': '楊',
    'chu': '朱',
    'tần': '秦',
    'vưu': '尤',
    'hứa': '許',
    'hà': '何',
    'lữ': '呂',
    'thi': '施',
    'trương': '張',
    'khổng': '孔',
    'tào': '曹',
    'nghiêm': '嚴',
    'hoa': '華',
    'kim': '金',
    'ngụy': '魏',
    'đào': '陶',
    'khương': '姜',
    'giới': '戚',
    'tạ': '謝',
    'trâu': '鄒',
    'du': '喻',
    'bách': '柏',
    'thủy': '水',
    'đậu': '竇',
    'chương': '章',
    'vân': '雲',
    'tô': '蘇',
    'phan': '潘',
    'cát': '葛',
    'phó': '傅',
    'lục': '陸',
    'phong': '豐',
    'ô': '烏',
    'tiêu': '焦',
    'ba': '巴',
    'cung': '弓',
    'mục': '牧',
    'ung': '雍',
    'tần': '秦',
    'lục': '陸',
    'bành': '彭',
    'lại': '賴',
}

# ─── TÊN PHỔ BIẾN VIỆT NAM (Tên đệm + Tên chính) ───
VIETNAMESE_COMMON_NAMES = {
    # Tên đệm phổ biến
    'văn': '文',
    'thị': '氏',
    'hữu': '有',
    'đình': '庭',
    'quốc': '國',
    'gia': '嘉',
    'trọng': '仲',
    'công': '公',
    'bá': '伯',
    'thái': '泰',
    'đức': '德',
    'xuân': '春',
    'thu': '秋',
    'đông': '冬',

    # Tên nam phổ biến
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
    'hòa': '和',
    'hoà': '和',
    'tâm': '心',
    'nghĩa': '義',
    'hiếu': '孝',
    'nhân': '仁',
    'tài': '才',
    'quang': '光',
    'huy': '輝',
    'nam': '南',
    'việt': '越',
    'tôn': '尊',
    'khoa': '科',
    'khải': '啟',
    'duy': '維',
    'long': '龍',
    'phong': '風',
    'thịnh': '盛',
    'tiến': '進',
    'bảo': '寶',
    'cường': '強',
    'hưng': '興',
    'lâm': '林',
    'khiêm': '謙',
    'khánh': '慶',
    'lực': '力',
    'nhật': '日',
    'nguyên': '元',
    'phát': '發',
    'phú': '富',
    'quý': '貴',
    'sang': '昇',
    'tân': '新',
    'thắng': '勝',
    'thiện': '善',
    'thuận': '順',
    'toàn': '全',
    'trí': '智',
    'trực': '直',
    'tú': '秀',
    'tuệ': '慧',
    'tường': '祥',
    'uy': '威',
    'vĩnh': '永',
    'vượng': '旺',
    'xuyên': '川',

    # Tên nữ phổ biến
    'diệu': '妙',
    'quỳnh': '瓊',
    'lan': '蘭',
    'hồng': '紅',
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
    'trang': '莊',
    'phương': '芳',
    'thảo': '草',
    'thủy': '水',
    'vân': '雲',
    'vy': '薇',
    'nhi': '兒',
    'nhung': '絨',
    'linh': '玲',
    'trinh': '貞',
    'tiên': '仙',
    'nguyệt': '月',
    'chi': '芝',
    'diệp': '葉',
    'hạnh': '幸',
    'hiền': '賢',
    'loan': '鸞',
    'ly': '璃',
    'my': '美',
    'hằng': '姮',
    'mai': '梅',
    'nga': '娥',
    'nhàn': '閒',
    'như': '如',
    'oanh': '鶯',
    'phụng': '鳳',
    'sen': '蓮',
    'thanh': '清',
    'thơm': '芬',
    'thu': '秋',
    'thúy': '翠',
    'thương': '霜',
    'trang': '莊',
    'trâm': '簪',
    'trầm': '沉',
    'uyên': '鴛',
    'vi': '薇',
    'xuân': '春',
}

# ─── TÊN ĐẦY ĐỦ ĐẶC BIỆT (override toàn bộ) ───
VIETNAMESE_FULL_NAMES = {
    'nguyễn văn a': '阮文甲',
    'trần văn b': '陳文乙',
    'trần kiên': '陳堅',
}


def get_vietnamese_name_in_chinese(full_name: str):
    """
    Dịch tên tiếng Việt sang tiếng Trung bằng từ điển nội bộ.

    Returns:
        str  - tên tiếng Trung nếu TẤT CẢ các phần đều dịch được
        None - nếu có bất kỳ phần nào không có trong từ điển
               (caller sẽ fallback sang Google Translate)
    """
    if not full_name or not full_name.strip():
        return None

    full_name_stripped = full_name.strip()
    full_name_lower = full_name_stripped.lower()

    # 1. Kiểm tra bảng tên đầy đủ trước (ưu tiên cao nhất)
    if full_name_lower in VIETNAMESE_FULL_NAMES:
        return VIETNAMESE_FULL_NAMES[full_name_lower]

    # 2. Tách từng phần
    parts = full_name_lower.split()
    if not parts:
        return None

    # 3. Dịch từng phần — nếu bất kỳ phần nào không có trong dict → trả về None
    result_parts = []
    for i, part in enumerate(parts):
        if i == 0:
            # Phần đầu tiên: ưu tiên tra họ
            if part in VIETNAMESE_SURNAMES:
                result_parts.append(VIETNAMESE_SURNAMES[part])
            elif part in VIETNAMESE_COMMON_NAMES:
                result_parts.append(VIETNAMESE_COMMON_NAMES[part])
            else:
                return None  # Họ không tìm thấy → fallback
        else:
            # Phần còn lại: tra tên đệm/tên chính
            if part in VIETNAMESE_COMMON_NAMES:
                result_parts.append(VIETNAMESE_COMMON_NAMES[part])
            elif part in VIETNAMESE_SURNAMES:
                # Một số tên đệm trùng với họ (ví dụ: Nguyễn Hoàng Tôn)
                result_parts.append(VIETNAMESE_SURNAMES[part])
            else:
                return None  # Tên không tìm thấy → fallback

    return ''.join(result_parts)
