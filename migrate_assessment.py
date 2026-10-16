# ============================================================
# ONE-TIME MIGRATION SCRIPT
# Adds assessment_type/max_score/periodic_labels to courses,
# and creates short_course_marks + periodic_marks tables.
# Run once, then delete this file.
# ============================================================
from sqlalchemy import text
from connections import engine


STATEMENTS = [
    # ---- 1. Columns on courses ----
    ("Add courses.assessment_type",
     "ALTER TABLE courses ADD COLUMN assessment_type VARCHAR(20) NOT NULL DEFAULT 'knec'"),

    ("Add courses.max_score",
     "ALTER TABLE courses ADD COLUMN max_score INT NOT NULL DEFAULT 100"),

    ("Add courses.periodic_labels",
     "ALTER TABLE courses ADD COLUMN periodic_labels VARCHAR(255) NULL"),

    # ---- 2. Default existing courses to knec ----
    ("Backfill existing courses to 'knec'",
     "UPDATE courses SET assessment_type = 'knec' WHERE assessment_type IS NULL OR assessment_type = ''"),

    # ---- 3. short_course_marks ----
    ("Create short_course_marks",
     """
     CREATE TABLE short_course_marks (
         id INT AUTO_INCREMENT PRIMARY KEY,
         student_id INT NOT NULL,
         subject_id INT NOT NULL,
         semester_id INT NULL,
         score FLOAT NULL,
         remarks VARCHAR(255) NULL,
         entered_at DATETIME DEFAULT CURRENT_TIMESTAMP,
         FOREIGN KEY (student_id) REFERENCES student_profiles(id),
         FOREIGN KEY (subject_id) REFERENCES subjects(id),
         FOREIGN KEY (semester_id) REFERENCES semesters(id),
         INDEX idx_scm_student_subject (student_id, subject_id),
         INDEX idx_scm_semester (semester_id)
     )
     """),

    # ---- 4. periodic_marks ----
    ("Create periodic_marks",
     """
     CREATE TABLE periodic_marks (
         id INT AUTO_INCREMENT PRIMARY KEY,
         student_id INT NOT NULL,
         subject_id INT NOT NULL,
         semester_id INT NULL,
         label VARCHAR(50) NOT NULL,
         score FLOAT NULL,
         remarks VARCHAR(255) NULL,
         entered_at DATETIME DEFAULT CURRENT_TIMESTAMP,
         FOREIGN KEY (student_id) REFERENCES student_profiles(id),
         FOREIGN KEY (subject_id) REFERENCES subjects(id),
         FOREIGN KEY (semester_id) REFERENCES semesters(id),
         INDEX idx_pm_student_subject (student_id, subject_id),
         INDEX idx_pm_semester (semester_id),
         INDEX idx_pm_label (label)
     )
     """),
]


def already_exists_error(msg: str) -> bool:
    """Return True if the error is one of the 'already done' kinds we can safely skip."""
    m = msg.lower()
    return (
        "duplicate column name" in m          # column already added
        or "already exists" in m              # table already created
        or "duplicate key name" in m          # index already created
    )


def main():
    print("🚀 Starting migration on:", engine.url.render_as_string(hide_password=True))
    print("-" * 60)

    applied = 0
    skipped = 0
    failed = 0

    with engine.begin() as conn:
        for label, sql in STATEMENTS:
            try:
                conn.execute(text(sql))
                print(f"✅ {label}")
                applied += 1
            except Exception as e:
                msg = str(e)
                if already_exists_error(msg):
                    print(f"⏭  {label} — already done, skipping")
                    skipped += 1
                else:
                    print(f"❌ {label}")
                    print(f"   → {msg}")
                    failed += 1

    print("-" * 60)
    print(f"Applied: {applied}  |  Skipped: {skipped}  |  Failed: {failed}")

    if failed == 0:
        print("\n🎉 Migration complete. You can delete migrate_assessment.py now.")
    else:
        print("\n⚠️  Some statements failed. Read the errors above and fix, then re-run.")
        print("   Re-running is safe — 'already done' steps are skipped.")


if __name__ == "__main__":
    main()