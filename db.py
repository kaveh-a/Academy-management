import sqlite3
import os

# --- تنظیمات دیتابیس ---
DB_PATH = "nihan_danesh.db"


def get_connection():
    """اتصال به دیتابیس SQLite"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row  # برای دسترسی به ستون‌ها به صورت نام (مانند دیکشنری)
    return conn


def init_db():
    """ایجاد دیتابیس و جداول اگر وجود نداشته باشن"""
    if not os.path.exists(DB_PATH):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # جدول users
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_code INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name TEXT NOT NULL UNIQUE,
                user_password TEXT NOT NULL,
                role TEXT NOT NULL
            )
        """)

        # جدول students
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS students (
                student_code INTEGER PRIMARY KEY AUTOINCREMENT,
                student_name TEXT NOT NULL,
                student_family TEXT NOT NULL,
                student_mobile TEXT,
                student_grade TEXT,
                student_gender TEXT
            )
        """)

        # جدول special_classes
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS special_classes (
                class_id INTEGER PRIMARY KEY AUTOINCREMENT,
                class_name TEXT NOT NULL,
                teacher_name TEXT NOT NULL,
                class_time TEXT,
                class_start_date TEXT,
                class_end_date TEXT
            )
        """)

        # جدول presence
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS presence (
                student_code INTEGER,
                price REAL,
                date TEXT,
                hour TEXT,
                payment_type TEXT,
                FOREIGN KEY (student_code) REFERENCES students(student_code)
            )
        """)

        # جدول payment
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS payment (
                payment_id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_code INTEGER,
                payment_amount REAL,
                payment_date TEXT,
                payment_time TEXT,
                payment_type TEXT,
                installments INTEGER,
                description TEXT,
                FOREIGN KEY (student_code) REFERENCES students(student_code)
            )
        """)

        # جدول class_enrollments
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS class_enrollments (
                enrollment_id INTEGER PRIMARY KEY AUTOINCREMENT,
                class_id INTEGER,
                student_code INTEGER,
                FOREIGN KEY (class_id) REFERENCES special_classes(class_id),
                FOREIGN KEY (student_code) REFERENCES students(student_code),
                UNIQUE (class_id, student_code)  -- ⚠️ جلوگیری از تکراری بودن
            )
        """)

        # اضافه کردن کاربران نمونه
        cursor.execute("INSERT OR IGNORE INTO users (user_name, user_password, role) VALUES (?, ?, ?)",
                       ("admin", "admin", "admin"))
        cursor.execute("INSERT OR IGNORE INTO users (user_name, user_password, role) VALUES (?, ?, ?)",
                       ("1", "1", "manager"))

        conn.commit()
        conn.close()
        print("✅ دیتابیس SQLite با تمام روابط ایجاد شد.")
    else:
        print("ℹ️ دیتابیس SQLite قبلاً وجود داشت.")


def check_user_login(username, password):
    """بررسی ورود کاربر"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT role FROM users WHERE user_name = ? AND user_password = ?", (username, password))
        result = cursor.fetchone()
        return result["role"] if result else None
    except Exception as e:
        print(f"❌ خطا در ورود کاربر: {e}")
        return None
    finally:
        conn.close()


def get_all_students():
    """دریافت لیست تمام دانش‌آموزان"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT student_code, student_name, student_family, student_mobile, 
                   student_grade, student_gender
            FROM students
        """)
        students = [dict(row) for row in cursor.fetchall()]

        for s in students:
            s["student_name"] = f"{s['student_name']} {s['student_family']}".strip()
            s.pop("student_family", None)
            s["id"] = str(s["student_code"])
            s["payments"] = []
            s["class"] = ""
            s["attendance"] = 0

        return students
    except Exception as e:
        print(f"❌ خطا در دریافت دانش‌آموزان: {e}")
        return []
    finally:
        conn.close()


def add_student(name, family, mobile, grade, gender):
    """افزودن دانش‌آموز جدید"""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO students (student_name, student_family, student_mobile, student_grade, student_gender)
            VALUES (?, ?, ?, ?, ?)
        """, (name, family, mobile, grade, gender))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ خطا در افزودن دانش‌آموز: {e}")
        return False


def update_student(student_code, name, family, mobile, grade):
    """ویرایش اطلاعات دانش‌آموز"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE students 
            SET student_name = ?, student_family = ?, student_mobile = ?, student_grade = ? 
            WHERE student_code = ?
        """, (name, family, mobile, grade, student_code))
        conn.commit()
        return cursor.rowcount > 0
    except Exception as e:
        print(f"❌ خطا در ویرایش دانش‌آموز: {e}")
        return False
    finally:
        conn.close()


def find_student_by_code(student_code):
    """جستجوی دانش‌آموز بر اساس کد"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT student_code, student_name, student_family, student_mobile, student_grade, student_gender
            FROM students WHERE student_code = ?
        """, (student_code,))
        row = cursor.fetchone()
        if row:
            row_dict = dict(row)
            row_dict["student_name"] = f"{row_dict['student_name']} {row_dict['student_family']}".strip()
            row_dict.pop("student_family", None)
            row_dict["id"] = str(row_dict["student_code"])
            return row_dict
        return None
    except Exception as e:
        print(f"❌ خطا در جستجوی دانش‌آموز: {e}")
        return None
    finally:
        conn.close()


def create_special_class(name, teacher, time, start_date, end_date):
    """ایجاد کلاس تخصصی جدید"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO special_classes (class_name, teacher_name, class_time, class_start_date, class_end_date)
            VALUES (?, ?, ?, ?, ?)
        """, (name, teacher, time, start_date, end_date))
        conn.commit()
        return True
    except Exception as e:
        print(f"❌ خطا در ایجاد کلاس: {e}")
        return False
    finally:
        conn.close()


def get_all_special_classes():
    """دریافت لیست تمام کلاس‌های تخصصی"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT class_id, class_name, teacher_name, class_time, class_start_date, class_end_date
            FROM special_classes ORDER BY class_id DESC
        """)
        return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"❌ خطا در دریافت کلاس‌ها: {e}")
        return []
    finally:
        conn.close()


def enroll_student_in_class(class_id, student_code):
    """ثبت دانش‌آموز در کلاس (بدون تکراری بودن)"""
    conn = get_connection()
    try:
        cursor = conn.cursor()

        # چک کردن تکراری بودن
        cursor.execute("""
            SELECT COUNT(*) FROM class_enrollments 
            WHERE class_id = ? AND student_code = ?
        """, (class_id, student_code))
        count = cursor.fetchone()[0]
        if count > 0:
            return False

        # اضافه کردن
        cursor.execute("""
            INSERT INTO class_enrollments (class_id, student_code)
            VALUES (?, ?)
        """, (class_id, student_code))
        conn.commit()
        return True
    except Exception as e:
        print(f"❌ خطا در ثبت دانش‌آموز در کلاس: {e}")
        return False
    finally:
        conn.close()


def get_students_in_class(class_id):
    """دریافت لیست دانش‌آموزان یک کلاس"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT s.student_code, s.student_name, s.student_family, s.student_mobile, s.student_grade
            FROM students s
            INNER JOIN class_enrollments ce ON s.student_code = ce.student_code
            WHERE ce.class_id = ?
        """, (class_id,))
        students = [dict(row) for row in cursor.fetchall()]
        for s in students:
            s["student_name"] = f"{s['student_name']} {s['student_family']}".strip()
            s.pop("student_family", None)
        return students
    except Exception as e:
        print(f"❌ خطا در دریافت دانش‌آموزان کلاس: {e}")
        return []
    finally:
        conn.close()


def add_payment_to_presence(student_code, amount, date, time, payment_type, installments=None, description=None):
    """افزودن پرداخت به جدول payment"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO payment 
            (student_code, payment_amount, payment_date, payment_time, payment_type, installments, description)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (student_code, amount, date, time, payment_type, installments, description))
        conn.commit()
        return True
    except Exception as e:
        print(f"❌ خطا در ثبت پرداخت: {e}")
        return False
    finally:
        conn.close()


def get_payments_by_student_code(student_code):
    """دریافت تمام پرداخت‌های یک دانش‌آموز"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM payment 
            WHERE student_code = ? 
            ORDER BY payment_id DESC
        """, (student_code,))
        return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"❌ خطا در دریافت پرداخت‌ها: {e}")
        return []
    finally:
        conn.close()


def get_all_payments():
    """دریافت تمام پرداخت‌های کامل"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM payment WHERE payment_type = ?", ('کامل',))
        return [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        print(f"❌ خطا در خواندن پرداخت‌ها: {e}")
        return []
    finally:
        conn.close()