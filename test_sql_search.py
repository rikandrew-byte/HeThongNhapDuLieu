import time
from app import app, db, FormHistory
from sqlalchemy import text

with app.app_context():
    start = time.time()
    q = "may"
    # Cast to jsonb, remove heavy keys, cast to text and ILIKE
    sql_expr = text("(data_json::jsonb - 'photo' - 'qr_line' - 'document_images' - 'signature')::text ILIKE :q").bindparams(q=f'%{q}%')
    records = (
        FormHistory.query
        .filter_by(is_deleted=False)
        .filter(sql_expr)
        .all()
    )
    db_load_time = time.time() - start
    print(f"Loaded {len(records)} records in {db_load_time:.4f}s")
    for r in records:
        print(f"Match: {r.ho_ten} | {r.ma_so}")
