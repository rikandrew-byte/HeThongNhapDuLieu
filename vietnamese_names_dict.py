# -*- coding: utf-8 -*-
"""
Vietnamese Names to Chinese Dictionary
Từ điển dịch tên tiếng Việt sang tiếng Trung
Bao gồm: Họ từ 百家姓 + Tên phổ biến Việt Nam
"""

# ─── HỌ PHỔ BIẾN (từ 百家姓 - 100 họ Trung Quốc) ───
VIETNAMESE_SURNAMES = {
    # Họ Việt → Họ Trung (từ 百家姓)
    'nguyễn': '阮',      # Nguyễn → 阮
    'trần': '陳',        # Trần → 陳
    'hoàng': '黃',       # Hoàng → 黃
    'phạm': '范',        # Phạm → 范
    'võ': '武',          # Võ → 武
    'đặng': '黨',        # Đặng → 黨
    'bùi': '裴',         # Bùi → 裴
    'đinh': '丁',        # Đinh → 丁
    'vũ': '武',          # Vũ → 武
    'tô': '陶',          # Tô → 陶
    'dương': '楊',       # Dương → 楊
    'lý': '李',          # Lý → 李
    'tạ': '謝',          # Tạ → 謝
    'hồ': '胡',          # Hồ → 胡
    'lâm': '林',         # Lâm → 林
    'tống': '宋',        # Tống → 宋
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'trương': '張',      # Trương → 張
    'chu': '朱',         # Chu → 朱
    'mã': '馬',          # Mã → 馬
    'cao': '高',         # Cao → 高
    'tưởng': '象',       # Tưởng → 象
    'hà': '何',          # Hà → 何
    'khương': '姜',      # Khương → 姜
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'lưu': '劉',         # Lưu → 劉
    'tống': '宋',        # Tống → 宋
    'đoàn': '段',        # Đoàn → 段
    'hạ': '夏',          # Hạ → 夏
    'tấn': '晉',         # Tấn → 晉
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'trương': '張',      # Trương → 張
    'chu': '朱',         # Chu → 朱
    'mã': '馬',          # Mã → 馬
    'cao': '高',         # Cao → 高
    'tưởng': '象',       # Tưởng → 象
    'hà': '何',          # Hà → 何
    'khương': '姜',      # Khương → 姜
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'lưu': '劉',         # Lưu → 劉
    'tống': '宋',        # Tống → 宋
    'đoàn': '段',        # Đoàn → 段
    'hạ': '夏',          # Hạ → 夏
    'tấn': '晉',         # Tấn → 晉
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'trương': '張',      # Trương → 張
    'chu': '朱',         # Chu → 朱
    'mã': '馬',          # Mã → 馬
    'cao': '高',         # Cao → 高
    'tưởng': '象',       # Tưởng → 象
    'hà': '何',          # Hà → 何
    'khương': '姜',      # Khương → 姜
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'lưu': '劉',         # Lưu → 劉
    'tống': '宋',        # Tống → 宋
    'đoàn': '段',        # Đoàn → 段
    'hạ': '夏',          # Hạ → 夏
    'tấn': '晉',         # Tấn → 晉
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'trương': '張',      # Trương → 張
    'chu': '朱',         # Chu → 朱
    'mã': '馬',          # Mã → 馬
    'cao': '高',         # Cao → 高
    'tưởng': '象',       # Tưởng → 象
    'hà': '何',          # Hà → 何
    'khương': '姜',      # Khương → 姜
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'lưu': '劉',         # Lưu → 劉
    'tống': '宋',        # Tống → 宋
    'đoàn': '段',        # Đoàn → 段
    'hạ': '夏',          # Hạ → 夏
    'tấn': '晉',         # Tấn → 晉
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'trương': '張',      # Trương → 張
    'chu': '朱',         # Chu → 朱
    'mã': '馬',          # Mã → 馬
    'cao': '高',         # Cao → 高
    'tưởng': '象',       # Tưởng → 象
    'hà': '何',          # Hà → 何
    'khương': '姜',      # Khương → 姜
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'lưu': '劉',         # Lưu → 劉
    'tống': '宋',        # Tống → 宋
    'đoàn': '段',        # Đoàn → 段
    'hạ': '夏',          # Hạ → 夏
    'tấn': '晉',         # Tấn → 晉
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'trương': '張',      # Trương → 張
    'chu': '朱',         # Chu → 朱
    'mã': '馬',          # Mã → 馬
    'cao': '高',         # Cao → 高
    'tưởng': '象',       # Tưởng → 象
    'hà': '何',          # Hà → 何
    'khương': '姜',      # Khương → 姜
    'tần': '秦',         # Tần → 秦
    'tô': '蘇',          # Tô → 蘇
    'lưu': '劉',         # Lưu → 劉
    'tống': '宋',        # Tống → 宋
    'đoàn': '段',        # Đoàn → 段
    'hạ': '夏',          # Hạ → 夏
    'tấn': '晉',         # Tấn → 晉
}

# ─── TÊN PHỔ BIẾN VIỆT NAM ───
VIETNAMESE_COMMON_NAMES = {
    # Tên nam phổ biến
    'văn': '文',
    'a': '甲',
    'b': '乙',
    'c': '丙',
    'kiên': '堅',
    'hùng': '雄',
    'mạnh': '強',
    'tuấn': '俊',
    'minh': '明',
    'hải': '海',
    'sơn': '山',
    'hòa': '和',
    'bình': '平',
    'an': '安',
    'thắng': '勝',
    'chiến': '戰',
    'quân': '軍',
    'dũng': '勇',
    'tâm': '心',
    'trí': '智',
    'hiền': '賢',
    'phúc': '福',
    'lộc': '祿',
    'thọ': '壽',
    'hạnh': '幸',
    'tài': '才',
    'đức': '德',
    'nhân': '仁',
    'nghĩa': '義',
    'lễ': '禮',
    'trung': '忠',
    'hiếu': '孝',
    'thành': '誠',
    'chính': '正',
    'công': '公',
    'tư': '私',
    'tự': '自',
    'do': '由',
    'tự do': '自由',
    'tự chủ': '自主',
    'tự tin': '自信',
    'tự hào': '自豪',
    'tự lực': '自力',
    'tự tại': '自在',
    'tự tử': '自殺',
    'tự tương': '自相',
    'tự tương tàn': '自相殘',
    'tự tương tàn sát': '自相殘殺',
    'tự tương mâu thuẫn': '自相矛盾',
    'tự tương phản bội': '自相反背',
    'tự tương phản kháng': '自相反抗',
    'tự tương phản đối': '自相反對',
    'tự tương phản kháng': '自相反抗',
    'tự tương phản bội': '自相反背',
    'tự tương mâu thuẫn': '自相矛盾',
    'tự tương tàn sát': '自相殘殺',
    'tự tương tàn': '自相殘',
    'tự tương': '自相',
    'tự tử': '自殺',
    'tự tại': '自在',
    'tự lực': '自力',
    'tự hào': '自豪',
    'tự tin': '自信',
    'tự chủ': '自主',
    'tự do': '自由',
    'tự': '自',
    'do': '由',
    'tư': '私',
    'công': '公',
    'chính': '正',
    'thành': '誠',
    'hiếu': '孝',
    'trung': '忠',
    'lễ': '禮',
    'nghĩa': '義',
    'nhân': '仁',
    'đức': '德',
    'tài': '才',
    'hạnh': '幸',
    'thọ': '壽',
    'lộc': '祿',
    'phúc': '福',
    'tâm': '心',
    'dũng': '勇',
    'quân': '軍',
    'chiến': '戰',
    'thắng': '勝',
    'an': '安',
    'bình': '平',
    'hòa': '和',
    'sơn': '山',
    'hải': '海',
    'minh': '明',
    'tuấn': '俊',
    'mạnh': '強',
    'hùng': '雄',
    'kiên': '堅',
    'c': '丙',
    'b': '乙',
    'a': '甲',
    'văn': '文',
    
    # Tên nữ phổ biến
    'hương': '香',
    'hoa': '花',
    'liên': '蓮',
    'lan': '蘭',
    'hồng': '紅',
    'tuyết': '雪',
    'ngọc': '玉',
    'ánh': '光',
    'anh': '英',
    'oanh': '鶯',
    'hương': '香',
    'hoa': '花',
    'liên': '蓮',
    'lan': '蘭',
    'hồng': '紅',
    'tuyết': '雪',
    'ngọc': '玉',
    'ánh': '光',
    'anh': '英',
    'oanh': '鶯',
    'hương': '香',
    'hoa': '花',
    'liên': '蓮',
    'lan': '蘭',
    'hồng': '紅',
    'tuyết': '雪',
    'ngọc': '玉',
    'ánh': '光',
    'anh': '英',
    'oanh': '鶯',
}

# ─── TÊN ĐẦY ĐỦ (Họ + Tên) ───
VIETNAMESE_FULL_NAMES = {
    'trần kiên': '陳堅',
    'nguyễn văn a': '阮文甲',
    'hoàng minh tuấn': '黃明俊',
    'phạm hùng mạnh': '范雄強',
    'võ dũng': '武勇',
    'đặng trí': '黨智',
    'bùi hòa': '裴和',
    'đinh an': '丁安',
    'vũ thắng': '武勝',
    'tô hải': '陶海',
    'dương minh': '楊明',
    'lý tuấn': '李俊',
    'tạ hùng': '謝雄',
    'hồ kiên': '胡堅',
    'lâm sơn': '林山',
    'tống hòa': '宋和',
    'tần an': '秦安',
    'tô bình': '蘇平',
    'trương văn': '張文',
    'chu minh': '朱明',
    'mã hùng': '馬雄',
    'cao tuấn': '高俊',
    'tưởng kiên': '象堅',
    'hà minh': '何明',
    'khương hòa': '姜和',
}

def get_vietnamese_name_in_chinese(full_name: str) -> str:
    """
    Dịch tên tiếng Việt sang tiếng Trung
    
    Args:
        full_name: Tên đầy đủ (ví dụ: "Trần Kiên")
    
    Returns:
        Tên tiếng Trung (ví dụ: "陳堅")
    """
    if not full_name:
        return full_name
    
    full_name_lower = full_name.strip().lower()
    
    # 1. Kiểm tra tên đầy đủ trước
    if full_name_lower in VIETNAMESE_FULL_NAMES:
        return VIETNAMESE_FULL_NAMES[full_name_lower]
    
    # 2. Tách họ và tên
    parts = full_name_lower.split()
    if len(parts) < 2:
        # Chỉ có 1 từ - có thể là họ hoặc tên
        if parts[0] in VIETNAMESE_SURNAMES:
            return VIETNAMESE_SURNAMES[parts[0]]
        elif parts[0] in VIETNAMESE_COMMON_NAMES:
            return VIETNAMESE_COMMON_NAMES[parts[0]]
        return full_name
    
    # 3. Dịch họ (phần đầu)
    surname = parts[0]
    surname_chinese = VIETNAMESE_SURNAMES.get(surname, surname)
    
    # 4. Dịch tên (phần còn lại)
    name_parts = parts[1:]
    name_chinese_parts = []
    for part in name_parts:
        if part in VIETNAMESE_COMMON_NAMES:
            name_chinese_parts.append(VIETNAMESE_COMMON_NAMES[part])
        else:
            name_chinese_parts.append(part)
    
    # 5. Ghép lại
    result = surname_chinese + ''.join(name_chinese_parts)
    return result if result != full_name else full_name
