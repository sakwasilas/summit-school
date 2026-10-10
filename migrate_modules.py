"""
One-time migration: create the `modules` table, add `module_id` FKs,
and backfill from the existing string `module` values.

Safe to run more than once — it checks before doing anything.
"""

import re
from connections import engine, SessionLocal, Base
from models import Course, Subject, StudentProfile, Module


# ------------------------------------------------------------
# 1. Create the new `modules` table and add `module_id` columns.
#    create_all only creates things that don't already exist, so this is safe.
# ------------------------------------------------------------
Base.metadata.create_all(engine)
print("[OK] Ensured `modules` table and `module_id` columns exist")


# ------------------------------------------------------------
# 2. Backfill.
# ------------------------------------------------------------
db = SessionLocal()


def parse_module_label(s):
    """
    Normalize 'MODULE 1', 'module 1', 'Mod 1', '1', 'Module 2', etc.
    to the canonical form 'Module N'. Returns None if nothing usable.
    """
    if not s:
        return None
    s = str(s).strip()
    if not s:
        return None
    m = re.search(r"(\d+)", s)
    if m:
        return f"Module {int(m.group(1))}"
    # no digit — keep the cleaned text, Title Case
    return s.title()


def get_or_create_module(course_id, name, order):
    """Return the existing Module for (course_id, name), or create one."""
    existing = db.query(Module).filter_by(course_id=course_id, name=name).first()
    if existing:
        return existing
    mod = Module(course_id=course_id, name=name, order=order)
    db.add(mod)
    db.flush()  # assign an id
    return mod


# ------------------------------------------------------------
# 2a. Backfill subjects
# ------------------------------------------------------------
subjects_total = 0
subjects_linked = 0

for s in db.query(Subject).all():
    subjects_total += 1

    # Try the string column first
    label = parse_module_label(s.module)

    # If empty, try to derive from the code prefix (e.g. '2920/101 ...' -> Module 1)
    if not label:
        m = re.match(r"^\s*(\d+)/(\d+)", (s.name or "").upper())
        if m:
            num = m.group(2)
            if num and num[0].isdigit():
                label = f"Module {int(num[0])}"

    if not label:
        continue  # leave unassigned — admin can fix later

    mnum = re.search(r"(\d+)", label)
    order = int(mnum.group(1)) if mnum else 0

    mod = get_or_create_module(s.course_id, label, order)
    s.module_id = mod.id
    subjects_linked += 1


# ------------------------------------------------------------
# 2b. Backfill students
# ------------------------------------------------------------
students_total = 0
students_linked = 0

for p in db.query(StudentProfile).all():
    students_total += 1

    label = parse_module_label(p.module)
    if not label:
        continue

    mnum = re.search(r"(\d+)", label)
    order = int(mnum.group(1)) if mnum else 0

    mod = get_or_create_module(p.course_id, label, order)
    p.module_id = mod.id
    students_linked += 1


db.commit()

print(f"[OK] Subjects:  {subjects_linked}/{subjects_total} linked to modules")
print(f"[OK] Students:  {students_linked}/{students_total} linked to modules")


# ------------------------------------------------------------
# 3. Show every module that now exists
# ------------------------------------------------------------
print("\nModules now defined in the database:")
modules = (
    db.query(Module)
    .order_by(Module.course_id, Module.order, Module.name)
    .all()
)
if not modules:
    print("  (none — nothing to backfill, or no subjects/students had module labels)")
else:
    for m in modules:
        # Get the course name for readability
        c = db.query(Course).filter_by(id=m.course_id).first()
        cname = c.name if c else "?"
        print(f"  id={m.id:3d}  order={m.order}  {m.name:<12s}  course='{cname}'")

db.close()
print("\nMigration complete.")