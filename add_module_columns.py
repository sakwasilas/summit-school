from connections import engine
from sqlalchemy import text

with engine.begin() as conn:
    try:
        conn.execute(text("ALTER TABLE subjects ADD COLUMN module_id INT NULL"))
        conn.execute(text(
            "ALTER TABLE subjects ADD CONSTRAINT fk_subjects_module "
            "FOREIGN KEY (module_id) REFERENCES modules(id)"
        ))
        print("[OK] subjects.module_id added")
    except Exception as e:
        print(f"[SKIP] subjects.module_id: {e}")

    try:
        conn.execute(text("ALTER TABLE student_profiles ADD COLUMN module_id INT NULL"))
        conn.execute(text(
            "ALTER TABLE student_profiles ADD CONSTRAINT fk_students_module "
            "FOREIGN KEY (module_id) REFERENCES modules(id)"
        ))
        print("[OK] student_profiles.module_id added")
    except Exception as e:
        print(f"[SKIP] student_profiles.module_id: {e}")