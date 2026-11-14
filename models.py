from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Project(db.Model):
    __tablename__ = 'projects'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.String(500), default='')
    created_at = db.Column(db.String(50))
    fields = db.relationship('ProjectField', backref='project', cascade='all, delete-orphan')


class ProjectField(db.Model):
    __tablename__ = 'project_fields'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    field_name = db.Column(db.String(100), nullable=False)
    field_unit = db.Column(db.String(50), default='')
    sort_order = db.Column(db.Integer, default=0)
    field_type = db.Column(db.String(20), default='number')


class Patient(db.Model):
    __tablename__ = 'patients'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    patient_code = db.Column(db.String(50), default='')
    created_at = db.Column(db.String(50))
    hospital_id = db.Column(db.String(50), default='')
    admission_count = db.Column(db.String(20), default='')
    admission_date = db.Column(db.String(50), default='')
    admission_dept = db.Column(db.String(100), default='')
    discharge_date = db.Column(db.String(50), default='')
    name = db.Column(db.String(50), default='')
    gender = db.Column(db.String(10), default='')
    age = db.Column(db.String(20), default='')
    ethnicity = db.Column(db.String(30), default='')
    hometown = db.Column(db.String(100), default='')
    address = db.Column(db.String(200), default='')
    occupation = db.Column(db.String(50), default='')
    marital_status = db.Column(db.String(20), default='')
    height = db.Column(db.String(20), default='')
    weight = db.Column(db.String(20), default='')
    main_diagnosis = db.Column(db.String(200), default='')
    diagnosis = db.Column(db.String(200), default='')
    discharge_status = db.Column(db.String(50), default='')
    outcome = db.Column(db.String(50), default='')
    other_complications = db.Column(db.Text, default='')
    values = db.relationship('PatientValue', backref='patient', cascade='all, delete-orphan')
    created_by = db.Column(db.String(50), default='') # 录入者姓名
    updated_by = db.Column(db.String(50), default='') # 最后修改者
    center = db.Column(db.String(50), default='')     # 所属中心


class PatientValue(db.Model):
    __tablename__ = 'patient_values'
    id = db.Column(db.Integer, primary_key=True)
    patient_id = db.Column(db.Integer, db.ForeignKey('patients.id'), nullable=False)
    field_id = db.Column(db.Integer, db.ForeignKey('project_fields.id'), nullable=False)
    value = db.Column(db.String(500), default='')
class ProjectMember(db.Model):
    __tablename__ = 'project_members'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    name = db.Column(db.String(50), nullable=False)
    role = db.Column(db.String(20), default='成员')
    center = db.Column(db.String(50), default='未分类中心') # 👈 确认有这一行
    centers_count = db.Column(db.Integer, default=1) 
    created_at = db.Column(db.String(20))