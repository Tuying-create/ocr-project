# init_db.py
from app import app, db
# 如果你的 models.py 里定义了所有表，确保这里被导入
from models import Project, ProjectField, Patient, PatientValue, ProjectMember

with app.app_context():
    db.create_all()
    print("✅ 数据库表已检查/创建完毕。")
