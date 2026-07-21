import os
import sys
import json
import firebase_admin
from firebase_admin import credentials, firestore

# Add root folder to python path so we can import app.py
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from app import app, db, FormHistory, Broker, Factory, OrderDoc

# Path to service account
SERVICE_ACCOUNT_PATH = r"D:\Software\FCT Human Resource\FCT Human Resource\FCT-HR-MANAGER\service-account.json"
FIREBASE_CONFIG_PATH = r"D:\Software\FCT Human Resource\FCT Human Resource\FCT-HR-MANAGER\firebase-applet-config.json"

if not os.path.exists(SERVICE_ACCOUNT_PATH):
    print(f"Error: service-account.json not found at {SERVICE_ACCOUNT_PATH}")
    sys.exit(1)

# Read config for databaseId
database_id = "(default)"
if os.path.exists(FIREBASE_CONFIG_PATH):
    with open(FIREBASE_CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
        database_id = cfg.get("firestoreDatabaseId", "(default)")

print(f"Initializing Firebase Admin with database: {database_id}...")
cred = credentials.Certificate(SERVICE_ACCOUNT_PATH)
firebase_admin.initialize_app(cred)
firestore_db = firestore.client(database_id=database_id)

with app.app_context():
    # 1. Migrate Brokers
    print("Migrating Brokers...")
    brokers_ref = firestore_db.collection('brokers').stream()
    for doc in brokers_ref:
        data = doc.to_dict()
        b_id = doc.id
        b_name = data.get('name', '')
        
        broker = Broker.query.get(b_id)
        if not broker:
            broker = Broker(id=b_id, name=b_name)
            db.session.add(broker)
        else:
            broker.name = b_name
    db.session.commit()
    print("Brokers migrated.")

    # 2. Migrate Factories
    print("Migrating Factories...")
    factories_ref = firestore_db.collection('factories').stream()
    for doc in factories_ref:
        data = doc.to_dict()
        f_id = doc.id
        f_name = data.get('name', '')
        b_id = data.get('broker_id', '')
        
        factory = Factory.query.get(f_id)
        if not factory:
            factory = Factory(id=f_id, name=f_name, broker_id=b_id)
            db.session.add(factory)
        else:
            factory.name = f_name
            factory.broker_id = b_id
    db.session.commit()
    print("Factories migrated.")

    # 3. Migrate Docs
    print("Migrating Order Docs...")
    docs_ref = firestore_db.collection('docs').stream()
    for doc in docs_ref:
        data = doc.to_dict()
        d_id = doc.id
        factory_id = data.get('factory_id', '')
        d_type = data.get('type', '')
        code = data.get('code', '')
        expiry_date = data.get('expiry_date', '')
        received_date = data.get('received_date', '')
        capacity = int(data.get('capacity', 0) or 0)
        parent_appraisal_id = data.get('parent_appraisal_id', '')
        note = data.get('note', '')
        
        order_doc = OrderDoc.query.get(d_id)
        if not order_doc:
            order_doc = OrderDoc(
                id=d_id, factory_id=factory_id, type=d_type, code=code,
                expiry_date=expiry_date, received_date=received_date,
                capacity=capacity, parent_appraisal_id=parent_appraisal_id,
                note=note
            )
            db.session.add(order_doc)
        else:
            order_doc.factory_id = factory_id
            order_doc.type = d_type
            order_doc.code = code
            order_doc.expiry_date = expiry_date
            order_doc.received_date = received_date
            order_doc.capacity = capacity
            order_doc.parent_appraisal_id = parent_appraisal_id
            order_doc.note = note
    db.session.commit()
    print("Order Docs migrated.")

    # 4. Migrate Placements and Workers
    print("Migrating Placements & Workers...")
    placements_ref = firestore_db.collection('placements').stream()
    
    # Pre-fetch all workers into a map for fast access
    workers_map = {}
    workers_ref = firestore_db.collection('workers').stream()
    for doc in workers_ref:
        workers_map[doc.id] = doc.to_dict()
        
    for doc in placements_ref:
        placement = doc.to_dict()
        worker_id = placement.get('worker_id')
        if not worker_id:
            continue
            
        worker = workers_map.get(worker_id, {})
        
        # Look for worker in SQL FormHistory by ma_so
        record = FormHistory.query.filter_by(ma_so=worker_id).order_by(FormHistory.ngay_tao.desc()).first()
        
        if not record:
            # Create a new record in FormHistory if not found, to avoid losing data
            full_name = worker.get('full_name') or placement.get('worker_name') or worker_id
            record = FormHistory(
                ma_so=worker_id,
                ho_ten=full_name,
                ten_file=f"{worker_id}_imported.html",
                data_json=json.dumps({
                    "Maso": worker_id,
                    "Hoten": full_name,
                    "Ngaysinh": worker.get('date_of_birth') or '',
                    "Lienhe": worker.get('phone_number') or '',
                    "Noio": worker.get('hometown') or '',
                    "photo": worker.get('avatar_url') or ''
                }, ensure_ascii=False)
            )
            db.session.add(record)
            
        # Update progress and document fields
        record.is_selected = True
        record.selected_job = placement.get('factory_name', '')
        
        record.passport_expiry = worker.get('passport_expiry', '')
        record.id_card_expiry = worker.get('id_card_expiry', '')
        record.health_check_expiry = worker.get('health_check_expiry', '')
        record.judicial_record_2_expiry = worker.get('judicial_record_2_expiry', '')
        
        record.placement_status = placement.get('status', 'GOM_HO_SO')
        record.factory_id = placement.get('factory_id', '')
        record.appraisal_id = placement.get('appraisal_id', '')
        record.visa_id = placement.get('visa_id', '')
        record.placement_note = placement.get('note', '')
        
        # Timeline dates
        record.date_trinh_cuc = placement.get('date_trinh_cuc', '')
        record.date_trinh_cuc_expected = placement.get('date_trinh_cuc_expected', '')
        record.date_lam_visa = placement.get('date_lam_visa', '')
        record.date_nhan_visa = placement.get('date_nhan_visa', '')
        record.date_xuat_canh = placement.get('date_xuat_canh', '')
        record.date_xuat_canh_actual = placement.get('date_xuat_canh_actual', '')
        
    db.session.commit()
    print("Placements & Workers migrated successfully!")
