import sqlite3, os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nomina.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS form1_applicant (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        full_name TEXT, mother_name TEXT, gender TEXT, birth_date TEXT,
        birth_country TEXT, residence TEXT, religion TEXT, marital_status TEXT,
        children_count TEXT, health_status TEXT, injury_type TEXT,
        employment_status TEXT, ready_to_work TEXT, travel_allowed TEXT,
        address TEXT, phone TEXT, whatsapp TEXT, email TEXT,
        education TEXT, field TEXT, specialty TEXT, institution TEXT,
        graduation_rate TEXT, graduation_year TEXT,
        work_name TEXT, work_title TEXT, work_start TEXT, work_end TEXT, work_reason TEXT,
        course_type TEXT, course_level TEXT, course_duration TEXT, course_date TEXT,
        course_org TEXT, course_doc TEXT,
        signature TEXT, form_date TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS form2_technical (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        applicant_id INTEGER, full_name TEXT, form_number TEXT, form_date TEXT,
        motivation TEXT, teamwork TEXT, communication TEXT, problem_solving TEXT,
        professionalism TEXT, confidence TEXT, organization TEXT, values_alignment TEXT,
        tech_scores TEXT, strengths TEXT, weaknesses TEXT, recommendation TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    c.execute("""
    CREATE TABLE IF NOT EXISTS form3_administrative (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        applicant_id INTEGER, full_name TEXT, specialty TEXT, form_number TEXT, form_date TEXT,
        degree_match TEXT, experience_required TEXT, certs_exist TEXT, courses_exist TEXT,
        employment_status TEXT, flexible_start TEXT, flexible_hours TEXT, willing_to_travel TEXT,
        expected_salary TEXT, within_budget TEXT,
        motivation TEXT, teamwork TEXT, communication TEXT, appearance TEXT,
        confidence TEXT, organization TEXT, values_alignment TEXT,
        strengths TEXT, weaknesses TEXT, recommendation TEXT,
        committee_rec TEXT, security_status TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.commit()
    conn.close()
    print("Database initialized at:", DB_PATH)

if __name__ == "__main__":
    init_db()
