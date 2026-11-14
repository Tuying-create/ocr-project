import os
from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
import re
import uuid
import datetime
import threading
import time
import cv2
import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///data.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

from models import db, Project, ProjectField, Patient, PatientValue, ProjectMember
db.init_app(app)

with app.app_context():
    db.create_all()

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


# ========== OCR 引擎 ==========
_ocr_engine = None
_ocr_loading = False
_ocr_lock = threading.Lock()

def _load_ocr():
    global _ocr_engine, _ocr_loading
    try:
        from rapidocr import RapidOCR
        _ocr_engine = RapidOCR()
        print('RapidOCR loaded successfully')
    except Exception as e:
        print('OCR load error:', e)
    finally:
        _ocr_loading = False

def get_ocr():
    global _ocr_loading
    with _ocr_lock:
        if _ocr_engine is not None:
            return _ocr_engine
        if not _ocr_loading:
            _ocr_loading = True
            t = threading.Thread(target=_load_ocr, daemon=True)
            t.start()
    return None

def ocr_image(image_path):
    try:
        ocr = get_ocr()
        if ocr is None:
            for _ in range(90):
                time.sleep(2)
                ocr = get_ocr()
                if ocr is not None:
                    break
            if ocr is None:
                return None

        img = cv2.imread(image_path)
        if img is not None:
            h, w = img.shape[:2]
            max_side = 1280
            if max(h, w) > max_side:
                scale = max_side / max(h, w)
                img = cv2.resize(img, None, fx=scale, fy=scale)
            result = ocr(img)
        else:
            result = ocr(image_path)

        if not result or not result.txts:
            return None
        return '\n'.join(result.txts)
    except Exception as e:
        print('OCR error:', e)
        return None


# ========== 项目 API ==========
@app.route('/api/projects', methods=['GET'])
def list_projects():
    projects = Project.query.order_by(Project.id.desc()).all()
    return jsonify([{
        'id': p.id,
        'name': p.name,
        'description': p.description,
        'created_at': p.created_at,
        'field_count': len(p.fields)
    } for p in projects])

@app.route('/api/project', methods=['POST'])
def create_project():
    data = request.json
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'error': '项目名称不能为空'}), 400
    project = Project(
        name=name,
        description=data.get('description', ''),
        created_at=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    )
    db.session.add(project)
    db.session.flush()
    fields = data.get('fields', [])
    for i, f in enumerate(fields):
        field_name = f.get('field_name', '').strip()
        f_type = f.get('field_type', 'number') 
        field = ProjectField(
            project_id=project.id,
            field_name=field_name,
            field_unit=f.get('field_unit', '').strip(),
            sort_order=i,
            field_type=f_type
        )
        db.session.add(field)
    db.session.commit()
    return jsonify({'success': True, 'project_id': project.id})

@app.route('/api/project/<int:project_id>', methods=['GET'])
def get_project(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404
    return jsonify({
        'id': project.id,
        'name': project.name,
        'description': project.description,
        'created_at': project.created_at,
        'fields': [{
            'id': f.id,
            'field_name': f.field_name,
            'field_unit': f.field_unit,
            'sort_order': f.sort_order,
            'field_type': f.field_type
        } for f in sorted(project.fields, key=lambda x: x.sort_order)]
    })

@app.route('/api/project/<int:project_id>', methods=['DELETE'])
def delete_project(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404
    db.session.delete(project)
    db.session.commit()
    return jsonify({'success': True})

@app.route('/api/project/<int:project_id>/stats', methods=['GET'])
def get_project_stats(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404
    patients = Patient.query.filter_by(project_id=project_id).all()
    total_patients = len(patients)
    recorded_patients = 0
    total_values = 0
    for p in patients:
        vals = [v.value for v in p.values if v.value and v.value.strip()]
        if vals:
            recorded_patients += 1
            total_values += len(vals)
    return jsonify({
        'total_patients': total_patients,
        'recorded_patients': recorded_patients,
        'total_values': total_values
    })


# ========== 解析 Excel 表头与导入名单 ==========
@app.route('/api/parse_excel_header', methods=['POST'])
def parse_excel_header():
    if 'excel' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    file = request.files['excel']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ['.xls', '.xlsx']:
        return jsonify({'error': '只支持 .xls 或 .xlsx 格式'}), 400

    filename = str(uuid.uuid4()) + ext
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        headers = []
        if ext == '.xls':
            book = xlrd.open_workbook(filepath)
            sh = book.sheet_by_index(0)
            for col_idx in range(sh.ncols):
                cell_value = sh.cell_value(0, col_idx)
                if cell_value:
                    headers.append(str(cell_value).strip())
        else:
            wb = load_workbook(filepath, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                for cell in row:
                    if cell is not None:
                        headers.append(str(cell).strip())
            wb.close()
            
        return jsonify({'success': True, 'headers': headers})
    except Exception as e:
        return jsonify({'error': f'解析失败: {str(e)}'}), 500

@app.route('/api/project/<int:project_id>/import_patients', methods=['POST'])
def import_patients(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404

    if 'excel' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    file = request.files['excel']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ['.xls', '.xlsx']:
        return jsonify({'error': '只支持 .xls 或 .xlsx 格式'}), 400

    filename = str(uuid.uuid4()) + ext
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        hospital_ids = []
        if ext == '.xls':
            book = xlrd.open_workbook(filepath)
            sh = book.sheet_by_index(0)
            for row_idx in range(1, sh.nrows):
                cell_value = sh.cell_value(row_idx, 0)
                if cell_value:
                    hospital_ids.append(str(cell_value).strip())
        else:
            wb = load_workbook(filepath, read_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[0] is not None:
                    hospital_ids.append(str(row[0]).strip())
            wb.close()

        count = 0
        for hospital_id in hospital_ids:
            if not hospital_id:
                continue
            existing = Patient.query.filter_by(project_id=project_id, hospital_id=hospital_id).first()
            if existing:
                continue
            patient = Patient(
                project_id=project_id,
                patient_code='',
                hospital_id=hospital_id,
                created_at=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            )
            db.session.add(patient)
            count += 1

        db.session.commit()
        return jsonify({'success': True, 'count': count})
    except Exception as e:
        return jsonify({'error': f'导入失败: {str(e)}'}), 500


# ========== 患者数据核心 API ==========
@app.route('/api/project/<int:project_id>/patient/upload', methods=['POST'])
def upload_patient(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404

    if 'images' not in request.files:
        return jsonify({'error': '没有上传图片'}), 400

    files = request.files.getlist('images')
    if not files:
        return jsonify({'error': '没有图片'}), 400

    all_text = ''
    image_urls = []

    for file in files:
        if file.filename == '':
            continue
        ext = os.path.splitext(file.filename)[1] or '.jpg'
        filename = str(uuid.uuid4()) + ext
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        image_urls.append(f'/uploads/{filename}')
        
        text = ocr_image(filepath)
        if text:
            all_text += text + '\n'

    if not all_text:
        return jsonify({'error': 'OCR识别失败'}), 500

    # 识别基本信息
    basic_info = {'hospital_id': '', 'name': '', 'gender': '', 'age': ''}
    match = re.search(r'住院号[：:\s]*([A-Za-z0-9]+)', all_text)
    if match: basic_info['hospital_id'] = match.group(1)
    match = re.search(r'姓名[：:\s]*([^\s\n]+)', all_text)
    if match: basic_info['name'] = match.group(1).strip()
    match = re.search(r'性别[：:\s]*([男女])', all_text)
    if match: basic_info['gender'] = match.group(1)
    match = re.search(r'年龄[：:\s]*(\d+)', all_text)
    if match: basic_info['age'] = match.group(1)

    # 识别指标
    fields = sorted(project.fields, key=lambda x: x.sort_order)
    matched = []

    # 识别住院/出院日期
    date_value = ''
    date_patterns = [
        r'住院日期[：:\s]*(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
        r'入院日期[：:\s]*(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
        r'入院时间[：:\s]*(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
    ]
    for p in date_patterns:
        match = re.search(p, all_text)
        if match:
            date_value = match.group(1).replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-').replace('.', '-')
            break
    matched.append({'field_id': -100, 'field_name': '住院日期', 'field_unit': '', 'field_type': 'text', 'value': date_value})

    discharge_value = ''
    discharge_patterns = [
        r'出院日期[：:\s]*(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
        r'出院时间[：:\s]*(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
    ]
    for p in discharge_patterns:
        match = re.search(p, all_text)
        if match:
            discharge_value = match.group(1).replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-').replace('.', '-')
            break
    matched.append({'field_id': -101, 'field_name': '出院日期', 'field_unit': '', 'field_type': 'text', 'value': discharge_value})

    for f in fields:
        value = ''
        field_type = f.field_type or 'number'

        if field_type == 'yesno':
            value = '是' if f.field_name in all_text else ''
        elif field_type == 'text':
            pattern = re.escape(f.field_name) + r'[：:\s]*([^\s\n]+)'
            match = re.search(pattern, all_text)
            if match: value = match.group(1).strip()
        else:
            safe_name = re.escape(f.field_name)
            patterns = [safe_name + r'[^\d]*([\d.]+)', safe_name + r'[^\d]{0,10}?([\d.]+)']
            for p in patterns:
                match = re.search(p, all_text)
                if match:
                    value = match.group(1)
                    break
            if not value:
                chars = [c for c in f.field_name if '\u4e00' <= c <= '\u9fff']
                if chars:
                    hit_count = sum(1 for c in chars if c in all_text)
                    if hit_count >= len(chars) * 0.6:
                        for c in chars:
                            idx = all_text.find(c)
                            if idx >= 0:
                                m = re.search(r'([\d.]+)', all_text[idx:idx+30])
                                if m: 
                                    value = m.group(1)
                                    break

        matched.append({'field_id': f.id, 'field_name': f.field_name, 'field_unit': f.field_unit, 'field_type': field_type, 'value': value})

    return jsonify({
        'success': True,
        'raw_text': all_text,
        'basic_info': basic_info,
        'data': matched,
        'image_urls': image_urls
    })


# ========== 患者数据保存与查询（整合协作功能） ==========
@app.route('/api/project/<int:project_id>/patient/save', methods=['POST'])
def save_patient(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404

    data = request.json
    basic = data.get('basic', {})
    
    # 获取前端传来的当前操作者信息
    operator_name = data.get('operator_name', '未知用户')
    operator_center = data.get('operator_center', '未分类中心')

    patient_id = data.get('patient_id')
    if patient_id:
        patient = Patient.query.get(patient_id)
        if not patient:
            return jsonify({'error': '患者不存在'}), 404
        for key, value in basic.items():
            if hasattr(patient, key) and value:
                setattr(patient, key, value)
        # 记录最后修改人
        patient.updated_by = operator_name 
    else:
        patient_code = data.get('patient_code', '').strip() or basic.get('hospital_id', '') or f'P{Patient.query.filter_by(project_id=project_id).count() + 1:03d}'
        patient = Patient(
            project_id=project_id,
            patient_code=patient_code,
            created_at=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            hospital_id=basic.get('hospital_id', ''),
            admission_count=basic.get('admission_count', ''),
            admission_date=basic.get('admission_date', ''),
            admission_dept=basic.get('admission_dept', ''),
            discharge_date=basic.get('discharge_date', ''),
            name=basic.get('name', ''),
            gender=basic.get('gender', ''),
            age=basic.get('age', ''),
            ethnicity=basic.get('ethnicity', ''),
            hometown=basic.get('hometown', ''),
            address=basic.get('address', ''),
            occupation=basic.get('occupation', ''),
            marital_status=basic.get('marital_status', ''),
            height=basic.get('height', ''),
            weight=basic.get('weight', ''),
            main_diagnosis=basic.get('main_diagnosis', ''),
            diagnosis=basic.get('diagnosis', ''),
            discharge_status=basic.get('discharge_status', ''),
            outcome=basic.get('outcome', ''),
            other_complications=basic.get('other_complications', ''),
            # 记录协作信息
            created_by=operator_name,  
            updated_by=operator_name,
            center=operator_center      
        )
        db.session.add(patient)
        db.session.flush()

    if 'other_complications' in basic:
        patient.other_complications = basic['other_complications']

    for item in data.get('values', []):
        field_id = item.get('field_id')
        value = item.get('value', '')
        if field_id in [-100, -101]:
            continue
        existing = PatientValue.query.filter_by(patient_id=patient.id, field_id=field_id).first()
        if existing:
            existing.value = value
        else:
            db.session.add(PatientValue(patient_id=patient.id, field_id=field_id, value=value))

    db.session.commit()
    return jsonify({'success': True, 'patient_id': patient.id})


@app.route('/api/patient/<int:patient_id>', methods=['GET'])
def get_patient(patient_id):
    patient = Patient.query.get(patient_id)
    if not patient:
        return jsonify({'error': '患者不存在'}), 404
    values = {v.field_id: v.value for v in patient.values}
    return jsonify({
        'id': patient.id,
        'project_id': patient.project_id,
        'patient_code': patient.patient_code,
        'hospital_id': patient.hospital_id,
        'admission_count': patient.admission_count,
        'admission_date': patient.admission_date,
        'admission_dept': patient.admission_dept,
        'discharge_date': patient.discharge_date,
        'name': patient.name,
        'gender': patient.gender,
        'age': patient.age,
        'ethnicity': patient.ethnicity,
        'hometown': patient.hometown,
        'address': patient.address,
        'occupation': patient.occupation,
        'marital_status': patient.marital_status,
        'height': patient.height,
        'weight': patient.weight,
        'main_diagnosis': patient.main_diagnosis,
        'diagnosis': patient.diagnosis,
        'discharge_status': patient.discharge_status,
        'outcome': patient.outcome,
        'other_complications': patient.other_complications or '',
        'values': values
    })

@app.route('/api/project/<int:project_id>/patients', methods=['GET'])
def list_patients(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404
    patients = Patient.query.filter_by(project_id=project_id).order_by(Patient.id.desc()).all()
    fields = sorted(project.fields, key=lambda x: x.sort_order)
    result = []
    for p in patients:
        values = {v.field_id: v.value for v in p.values}
        result.append({
            'id': p.id,
            'patient_code': p.patient_code,
            'hospital_id': p.hospital_id,
            'name': p.name,
            'created_at': p.created_at,
            'has_data': len([v for v in values.values() if v and v.strip()]) > 0,
            'values_count': len([v for v in values.values() if v and v.strip()]),
            # 返回协作信息 (使用 getattr 兼容未更新 models 或旧数据库的情况)
            'created_by': getattr(p, 'created_by', '未知') or '未知',
            'center': getattr(p, 'center', '未分类中心') or '未分类中心'
        })
    return jsonify({'patients': result, 'fields': [{'id': f.id, 'name': f.field_name} for f in fields]})


# ========== 成员协作 API ==========
@app.route('/api/project/<int:project_id>/members', methods=['GET'])
def get_members(project_id):
    members = ProjectMember.query.filter_by(project_id=project_id).all()
    return jsonify([{
        'id': m.id, 
        'name': m.name, 
        'role': m.role, 
        'center': getattr(m, 'center', '未分类中心') # 兼容老数据
    } for m in members])

@app.route('/api/project/<int:project_id>/member', methods=['POST'])
def add_member(project_id):
    data = request.json
    name = data.get('name', '').strip()
    if not name: return jsonify({'error': '名字不能为空'}), 400
    
    operator_role = data.get('operator_role', 'member') # 获取当前操作者的角色
    new_role = data.get('role', '成员')

    # 1. 权限校验
    if operator_role == 'member':
        return jsonify({'error': '权限不足：普通成员不能添加成员'}), 403
    
    if operator_role == 'admin' and new_role in ['管理员', '创建者', 'admin', 'creator']:
        return jsonify({'error': '权限不足：管理员不能添加其他管理员'}), 403

    # 2. 写入数据库
    member = ProjectMember(
        project_id=project_id,
        name=name,
        role=new_role,
        center=data.get('center', '未分类中心'),
        centers_count=data.get('centers_count', 1),
        created_at=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    )
    db.session.add(member)
    db.session.commit()
    return jsonify({'success': True, 'member_id': member.id})

@app.route('/api/project/<int:project_id>/member/<int:member_id>', methods=['DELETE'])
def delete_member(project_id, member_id):
    # 获取前端传过来的操作者角色
    operator_role = request.args.get('operator_role', 'member')
    
    if operator_role == 'member':
        return jsonify({'error': '权限不足：普通成员不能删除成员'}), 403
        
    member = ProjectMember.query.get(member_id)
    if not member: return jsonify({'error': '成员不存在'}), 404
    
    if operator_role == 'admin' and member.role in ['管理员', '创建者']:
        return jsonify({'error': '权限不足：管理员不能删除管理员或创建者'}), 403

    db.session.delete(member)
    db.session.commit()
    return jsonify({'success': True})


# ========== 导出 Excel ==========
@app.route('/api/project/<int:project_id>/export', methods=['GET'])
def export_project(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404

    fields = sorted(project.fields, key=lambda x: x.sort_order)
    patients = Patient.query.filter_by(project_id=project_id).all()

    wb = Workbook()
    ws = wb.active
    ws.title = project.name[:30]

    # 在导出表头加上 录入者 和 录入中心
    header = ['住院号', '住院日期', '出院日期', '其他合并症', '录入者', '所属中心'] + [f.field_name + (f' ({f.field_unit})' if f.field_unit else '') for f in fields]
    ws.append(header)

    for p in patients:
        row = [
            p.hospital_id or '', 
            p.admission_date or '', 
            p.discharge_date or '', 
            p.other_complications or '',
            getattr(p, 'created_by', '') or '',
            getattr(p, 'center', '') or ''
        ]
        for f in fields:
            val = PatientValue.query.filter_by(patient_id=p.id, field_id=f.id).first()
            row.append(val.value if val else '')
        ws.append(row)

    for i in range(len(header)):
        col_letter = get_column_letter(i + 1)
        ws.column_dimensions[col_letter].width = 15

    filename = f'project_{project_id}.xlsx'
    filepath = filename
    wb.save(filepath)

    return send_file(filepath, as_attachment=True, download_name=f'{project.name}.xlsx')


# ========== 图片访问路由 ==========
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename)


# ========== 页面路由 ==========
@app.route('/')
def index():
    return send_from_directory('templates', 'projects.html')

@app.route('/projects')
def projects_page():
    return send_from_directory('templates', 'projects.html')

@app.route('/dashboard')
def dashboard_page():
    return send_from_directory('templates', 'dashboard.html')

@app.route('/project/new')
def project_new_page():
    return send_from_directory('templates', 'project_new.html')

@app.route('/project/<int:project_id>')
def project_detail_page(project_id):
    return send_from_directory('templates', 'project_detail.html')

@app.route('/project/<int:project_id>/patient/upload')
def patient_upload_page(project_id):
    return send_from_directory('templates', 'patient_upload.html')

@app.route('/patient/<int:patient_id>/edit')
def patient_edit_page(patient_id):
    return send_from_directory('templates', 'patient_upload.html')

@app.route('/project/<int:project_id>/members')
def members_page(project_id):
    return send_from_directory('templates', 'members.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)