import sqlite3
import hashlib
import os
from pathlib import Path
from datetime import datetime

DB_PATH = Path(__file__).resolve().parents[1] / "data" / "auth.db"

def get_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn

def hash_pw(password: str, salt: str = "PneumoScan_CAD_Salt_2026") -> str:
    return hashlib.sha256((password + salt).encode('utf-8')).hexdigest()

def init_auth_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS doctors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        doctor_id TEXT UNIQUE NOT NULL,
        email TEXT UNIQUE NOT NULL,
        full_name TEXT NOT NULL,
        cchn TEXT NOT NULL,
        department TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        created_at TEXT NOT NULL,
        last_login TEXT
    )
    """)
    conn.commit()
    
    # Tạo sẵn các tài khoản demo chuẩn y khoa nếu chưa có
    sample_doctors = [
        ("BV-RAD-1092", "bs.nguyenvana@hospital.vn", "BS. Nguyễn Văn A", "012345/BYT-CCHN", "Khoa Chẩn đoán Hình ảnh", "password123"),
        ("BV-RAD-8839", "bs.leminhtri@hospital.vn", "BS. Lê Minh Trí", "045892/BYT-CCHN", "Khoa Chẩn đoán Hình ảnh", "password123"),
        ("BV-ICU-2024", "bs.tranmai@hospital.vn", "BS. Trần Thị Mai", "098712/BYT-CCHN", "Khoa Hồi sức Cấp cứu (ICU)", "password123"),
    ]
    
    for doc_id, email, name, cchn, dept, pw in sample_doctors:
        cursor.execute("SELECT id FROM doctors WHERE doctor_id = ? OR email = ?", (doc_id, email))
        if cursor.fetchone() is None:
            cursor.execute("""
            INSERT INTO doctors (doctor_id, email, full_name, cchn, department, password_hash, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (doc_id, email, name, cchn, dept, hash_pw(pw), datetime.now().isoformat()))
    conn.commit()
    conn.close()

def authenticate_doctor(identifier: str, password: str):
    """Xác thực bằng Mã Bác sĩ hoặc Email bệnh viện"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT * FROM doctors 
    WHERE (doctor_id = ? OR LOWER(email) = LOWER(?))
    """, (identifier.strip(), identifier.strip()))
    row = cursor.fetchone()
    if row is None:
        conn.close()
        return None
    
    target_hash = hash_pw(password)
    if row["password_hash"] == target_hash:
        # Cập nhật thời điểm đăng nhập cuối
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("UPDATE doctors SET last_login = ? WHERE id = ?", (now_str, row["id"]))
        conn.commit()
        user_dict = dict(row)
        user_dict["last_login"] = now_str
        conn.close()
        return user_dict
    
    conn.close()
    return None

def register_doctor(doctor_id: str, email: str, full_name: str, cchn: str, department: str, password: str):
    """Đăng ký bác sĩ mới với đầy đủ chứng chỉ CCHN"""
    if not doctor_id or not email or not full_name or not cchn or not password:
        return False, "Vui lòng điền đầy đủ tất cả các trường bắt buộc."
    
    if len(password) < 6:
        return False, "Mật khẩu phải có độ dài tối thiểu 6 ký tự."
        
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        INSERT INTO doctors (doctor_id, email, full_name, cchn, department, password_hash, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            doctor_id.strip().upper(),
            email.strip().lower(),
            full_name.strip(),
            cchn.strip().upper(),
            department.strip(),
            hash_pw(password),
            datetime.now().isoformat()
        ))
        conn.commit()
        conn.close()
        return True, "Đăng ký tài khoản Bác sĩ thành công! Bạn có thể đăng nhập ngay."
    except sqlite3.IntegrityError as e:
        conn.close()
        if "doctor_id" in str(e):
            return False, f"Mã Bác sĩ '{doctor_id}' đã tồn tại trong hệ thống."
        elif "email" in str(e):
            return False, f"Email '{email}' đã được đăng ký trước đó."
        return False, "Dữ liệu đăng ký bị trùng lặp trong hệ thống."
    except Exception as e:
        conn.close()
        return False, f"Lỗi hệ thống: {str(e)}"

if __name__ == "__main__":
    init_auth_db()
    user = authenticate_doctor("BV-RAD-1092", "password123")
    print("Test auth success:", user is not None, user["full_name"] if user else None)

