import time, json
from app import app, db, FormHistory

with app.app_context():
    start = time.time()
    records = FormHistory.query.filter_by(is_deleted=False).all()
    db_load_time = time.time() - start
    print(f"Loaded {len(records)} records in {db_load_time:.4f}s")
    
    start_search = time.time()
    q = "may"
    matches = []
    
    def check_val(val):
        if isinstance(val, str):
            if val.startswith('data:image/') or len(val) > 2000:
                return False
            return q in val.lower()
        elif isinstance(val, (int, float)):
            return q in str(val)
        elif isinstance(val, list):
            return any(check_val(item) for item in val)
        elif isinstance(val, dict):
            return any(check_val(v) for k, v in val.items() if k not in ('photo', 'qr_line', 'document_images', 'signature'))
        return False

    for r in records:
        if r.data_json:
            try:
                data = json.loads(r.data_json)
                if check_val(data):
                    matches.append(r.ho_ten)
            except:
                pass
                
    search_time = time.time() - start_search
    print(f"Searched {len(records)} parsed JSONs in {search_time:.4f}s. Found {len(matches)} matches: {matches}")
