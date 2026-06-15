import os
from deep_translator import GoogleTranslator
import google.generativeai as genai
from unicodedata import normalize
from dotenv import load_dotenv

load_dotenv()

def translate_free(text: str) -> str:
    gemini_api_key = os.environ.get('GEMINI_API_KEY')
    if gemini_api_key:
        genai.configure(api_key=gemini_api_key)
    
    text_normalized = normalize('NFC', text)
    processed_text = text_normalized

    try:
        if gemini_api_key:
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = f"Bạn là chuyên gia dịch thuật CV xuất khẩu lao động Đài Loan. Hãy dịch đoạn kinh nghiệm làm việc sau sang tiếng Trung Phồn Thể. Yêu cầu: dịch sát nghĩa, chuẩn thuật ngữ nghề nghiệp (cơ khí, điện, xây dựng, nhà máy, dệt may...), giữ nguyên cách dòng và định dạng nếu có. Tuyệt đối KHÔNG kèm theo lời giải thích hay bình luận, chỉ trả về đúng kết quả dịch. Đoạn văn bản cần dịch: '{processed_text.strip()}'"
            response = model.generate_content(prompt)
            result = response.text.strip()
            print('Gemini Result:', result)
        else:
            result = GoogleTranslator(source='vi', target='zh-TW').translate(processed_text.strip())
            print('Google Result:', result)
        return result
    except Exception as e:
        print('Error:', e)
        return text_normalized

print("Translation 1:")
translate_free('KIỂM TRA CHẤT LƯỢNG CẦU ĐƯỜNG: CHẤT LƯỢNG BÊ TÔNG, NHỰA ĐƯỜNG... LU NỀN NỀN ĐƯỜNG.... VẬT LIỆU XÂY DỰNG 越南 (2021- NAY)')
print("\nTranslation 2:")
translate_free('Nhanh nhẹn, chịu khó、伏地挺身 10~30下、搬重 50kg以上')
