# -*- coding: utf-8 -*-
"""
Script Di Dân Dữ Liệu: Xata → Neon
Chiến dịch "Đại Di Tản" — Chạy một lần duy nhất.

Cách dùng:
  1. Đặt 2 biến môi trường bên dưới.
  2. Chạy: python migrate_xata_to_neon.py
"""

import os
import sys
import psycopg2
from datetime import datetime

# ═══════════════════════════════════════════════════════════════
# ⚙️  CẤU HÌNH — Đặt 2 chuỗi kết nối vào đây trước khi chạy
# ═══════════════════════════════════════════════════════════════
XATA_URL  = os.environ.get('XATA_URL', '')   # Connection string Xata cũ
NEON_URL  = os.environ.get('NEON_URL', '')   # Connection string Neon mới
# ═══════════════════════════════════════════════════════════════

def check_config():
    if not XATA_URL or not NEON_URL:
        print("❌ Thiếu cấu hình!")
        print("   Chạy lại với:")
        print("   $env:XATA_URL='postgresql://xata:...'  (PowerShell)")
        print("   $env:NEON_URL='postgresql://neondb_owner:...'")
        sys.exit(1)

def migrate():
    check_config()
    print("=" * 60)
    print("🚀 Bắt đầu Chiến dịch Đại Di Tản: Xata → Neon")
    print("=" * 60)

    # Kết nối nguồn (Xata)
    print("\n📡 Đang kết nối Xata (nguồn)...")
    try:
        xata_conn = psycopg2.connect(XATA_URL, connect_timeout=10)
        xata_cur  = xata_conn.cursor()
        print("✅ Kết nối Xata thành công.")
    except Exception as e:
        print(f"❌ Không thể kết nối Xata: {e}")
        sys.exit(1)

    # Kết nối đích (Neon)
    print("📡 Đang kết nối Neon (đích)...")
    try:
        neon_conn = psycopg2.connect(NEON_URL, connect_timeout=10)
        neon_cur  = neon_conn.cursor()
        print("✅ Kết nối Neon thành công.")
    except Exception as e:
        print(f"❌ Không thể kết nối Neon: {e}")
        xata_conn.close()
        sys.exit(1)

    # Tạo bảng trên Neon nếu chưa có
    print("\n📋 Đảm bảo bảng form_history tồn tại trên Neon...")
    neon_cur.execute("""
        CREATE TABLE IF NOT EXISTS form_history (
            id         SERIAL PRIMARY KEY,
            ma_so      VARCHAR(50),
            ho_ten     VARCHAR(100),
            ten_file   VARCHAR(255),
            data_json  TEXT,
            ngay_tao   TIMESTAMP DEFAULT NOW(),
            is_selected BOOLEAN DEFAULT FALSE
        );
    """)
    neon_conn.commit()
    print("✅ Bảng đã sẵn sàng.")

    # Đọc dữ liệu từ Xata
    print("\n📥 Đang đọc dữ liệu từ Xata...")
    try:
        xata_cur.execute("""
            SELECT id, ma_so, ho_ten, ten_file, data_json, ngay_tao, is_selected
            FROM form_history
            ORDER BY id ASC
        """)
        rows = xata_cur.fetchall()
        print(f"✅ Đã đọc {len(rows)} bản ghi từ Xata.")
    except Exception as e:
        # Thử lại không có cột is_selected (nếu Xata không có)
        try:
            xata_cur.execute("""
                SELECT id, ma_so, ho_ten, ten_file, data_json, ngay_tao
                FROM form_history
                ORDER BY id ASC
            """)
            rows_raw = xata_cur.fetchall()
            rows = [(r[0], r[1], r[2], r[3], r[4], r[5], False) for r in rows_raw]
            print(f"✅ Đã đọc {len(rows)} bản ghi từ Xata (không có cột is_selected).")
        except Exception as e2:
            print(f"❌ Lỗi đọc Xata: {e2}")
            sys.exit(1)

    if not rows:
        print("⚠️  Không có dữ liệu nào trong Xata để di chuyển.")
        return

    # Ghi vào Neon
    print(f"\n📤 Đang di chuyển {len(rows)} bản ghi sang Neon...")
    success = 0
    errors  = 0
    for row in rows:
        id_, ma_so, ho_ten, ten_file, data_json, ngay_tao, is_selected = row
        try:
            neon_cur.execute("""
                INSERT INTO form_history (id, ma_so, ho_ten, ten_file, data_json, ngay_tao, is_selected)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    ma_so       = EXCLUDED.ma_so,
                    ho_ten      = EXCLUDED.ho_ten,
                    ten_file    = EXCLUDED.ten_file,
                    data_json   = EXCLUDED.data_json,
                    ngay_tao    = EXCLUDED.ngay_tao,
                    is_selected = EXCLUDED.is_selected;
            """, (id_, ma_so, ho_ten, ten_file, data_json, ngay_tao, is_selected or False))
            success += 1
            if success % 10 == 0:
                print(f"   ... đã di chuyển {success}/{len(rows)}")
        except Exception as e:
            print(f"   ⚠️  Lỗi bản ghi ID={id_}: {e}")
            errors += 1

    # Đồng bộ sequence ID
    neon_cur.execute("SELECT setval('form_history_id_seq', (SELECT MAX(id) FROM form_history));")
    neon_conn.commit()

    # Kết quả
    print("\n" + "=" * 60)
    print(f"✅ Di chuyển thành công: {success}/{len(rows)} bản ghi")
    if errors:
        print(f"⚠️  Lỗi: {errors} bản ghi")
    print("🎉 Chiến dịch Đại Di Tản hoàn thành!")
    print("=" * 60)

    xata_conn.close()
    neon_conn.close()

if __name__ == '__main__':
    migrate()
