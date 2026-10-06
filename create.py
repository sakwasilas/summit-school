'''from connections import Base, engine
from models import User,Admin  


if __name__ == "__main__":
    Base.metadata.create_all(bind=engine)
    print("✅ Tables created successfully!")'''

'''
create.py
'''
# create.py
import os
from connections import Base, engine, SessionLocal
from models import User, Admin

# ✅ Create tables
Base.metadata.create_all(bind=engine)
print("✅ Tables created successfully!")

# ✅ Add default admin user
def add_admin_user():
    db = SessionLocal()
    try:
        existing_admin = db.query(Admin).filter_by(username="admin").first()
        if existing_admin:
            print("ℹ️ Admin user already exists.")
        else:
            admin = Admin(
                username="admin",
                password="admin123"   # ⚠️ plain text, later hash with werkzeug
            )
            db.add(admin)
            db.commit()
            print("✅ Admin user created! Username: admin | Password: admin123")
    finally:
        db.close()

if __name__ == "__main__":
    add_admin_user()
