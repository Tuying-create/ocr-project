import os
from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
import re
import uuid
import datetime
import threading
import time
import cv2
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///data.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

from models import db, Project, ProjectField, Patient, PatientValue
db.init_app(app)

with app.app_context():
    db.create_all()

UPLOAD_FOLDER = 'uploads'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


# ========== 字段类型判断 ==========

TEXT_INPUT_FIELDS = ['民族', '职业', '入院时间', '治疗方案']
YESNO_FIELDS = [
    '头痛', '视朦', '发热', '肢体肌力异常', '意识障碍', '脑膜刺激征',
    '高血压', '心肌损害', '低蛋白血症', '脑膜炎',
    '脑炎，累及脑实质'
]


def get_field_type(field_name):
    name = field_name.strip().lstrip('*').strip()
    if name in TEXT_INPUT_FIELDS:
        return 'text'
    if name in YESNO_FIELDS:
        return 'yesno'
    return 'number'


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


# ========== OCR API ==========

@app.route('/ocr', methods=['POST'])
def ocr_api():
    if 'image' not in request.files:
        return jsonify({'error': 'No image'}), 400
    file = request.files['image']
    if file.filename == '':
        return jsonify({'error': 'Empty filename'}), 400
    ext = os.path.splitext(file.filename)[1] or '.jpg'
    filename = str(uuid.uuid4()) + ext
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)
    text = ocr_image(filepath)
    if text is None:
        return jsonify({'error': 'OCR failed'}), 500
    return jsonify({'success': True, 'raw_text': text})


@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'message': 'OCR service running'})


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
        field = ProjectField(
            project_id=project.id,
            field_name=field_name,
            field_unit=f.get('field_unit', '').strip(),
            sort_order=i,
            field_type=get_field_type(field_name)
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


# ========== 解析 Excel 表头 ==========

@app.route('/api/parse_excel_header', methods=['POST'])
def parse_excel_header():
    if 'excel' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    file = request.files['excel']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    ext = os.path.splitext(file.filename)[1]
    filename = str(uuid.uuid4()) + ext
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        wb = load_workbook(filepath, read_only=True)
        ws = wb.active
        headers = []
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            for cell in row:
                if cell is not None:
                    headers.append(str(cell).strip())
        wb.close()
        return jsonify({'success': True, 'headers': headers})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ========== 导入患者名单 ==========

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

    ext = os.path.splitext(file.filename)[1]
    filename = str(uuid.uuid4()) + ext
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        wb = load_workbook(filepath, read_only=True)
        ws = wb.active

        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            hospital_id = str(row[0]).strip()
            if not hospital_id:
                continue

            existing = Patient.query.filter_by(
                project_id=project_id,
                hospital_id=hospital_id
            ).first()
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
        wb.close()
        return jsonify({'success': True, 'count': count})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ========== 患者数据 API ==========

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
    for file in files:
        if file.filename == '':
            continue
        ext = os.path.splitext(file.filename)[1] or '.jpg'
        filename = str(uuid.uuid4()) + ext
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        text = ocr_image(filepath)
        if text:
            all_text += text + '\n'

    if not all_text:
        return jsonify({'error': 'OCR识别失败'}), 500

    fields = sorted(project.fields, key=lambda x: x.sort_order)
    matched = []

    # 识别住院日期
    date_value = ''
    date_patterns = [
        r'住院日期[：:\s]*(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
        r'入院日期[：:\s]*(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
        r'入院时间[：:\s]*(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
        r'(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)',
    ]
    for p in date_patterns:
        match = re.search(p, all_text)
        if match:
            date_value = match.group(1).replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-').replace('.', '-')
            break

    matched.append({
        'field_id': -100,
        'field_name': '住院日期',
        'field_unit': '',
        'field_type': 'text',
        'value': date_value
    })

    # 识别出院日期
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

    matched.append({
        'field_id': -101,
        'field_name': '出院日期',
        'field_unit': '',
        'field_type': 'text',
        'value': discharge_value
    })

    for f in fields:
        value = ''
        field_type = f.field_type or get_field_type(f.field_name)

        if field_type == 'yesno':
            if f.field_name in all_text:
                value = '是'
            else:
                value = ''
        elif field_type == 'text':
            pattern = re.escape(f.field_name) + r'[：:\s]*([^\s\n]+)'
            match = re.search(pattern, all_text)
            if match:
                value = match.group(1).strip()
        else:
            safe_name = re.escape(f.field_name)
            patterns = [
                safe_name + r'[^\d]*([\d.]+)',
                safe_name + r'[^\d]{0,10}?([\d.]+)',
            ]
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
                                snippet = all_text[idx:idx+30]
                                m = re.search(r'([\d.]+)', snippet)
                                if m:
                                    value = m.group(1)
                                    break

        matched.append({
            'field_id': f.id,
            'field_name': f.field_name,
            'field_unit': f.field_unit,
            'field_type': field_type,
            'value': value
        })

    return jsonify({
        'success': True,
        'raw_text': all_text,
        'data': matched,
        'count': len([m for m in matched if m['value']])
    })


@app.route('/api/project/<int:project_id>/patient/save', methods=['POST'])
def save_patient(project_id):
    project = Project.query.get(project_id)
    if not project:
        return jsonify({'error': '项目不存在'}), 404

    data = request.json
    basic = data.get('basic', {})

    patient_id = data.get('patient_id')
    if patient_id:
        patient = Patient.query.get(patient_id)
        if not patient:
            return jsonify({'error': '患者不存在'}), 404
        for key, value in basic.items():
            if hasattr(patient, key) and value:
                setattr(patient, key, value)
    else:
        patient_code = data.get('patient_code', '').strip()
        if not patient_code:
            count = Patient.query.filter_by(project_id=project_id).count()
            patient_code = f'P{count + 1:03d}'

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
            other_complications=basic.get('other_complications', '')
        )
        db.session.add(patient)
        db.session.flush()

    if 'other_complications' in basic:
        patient.other_complications = basic['other_complications']

    for item in data.get('values', []):
        field_id = item.get('field_id')
        value = item.get('value', '')
        if field_id == -100 or field_id == -101:
            continue
        existing = PatientValue.query.filter_by(
            patient_id=patient.id,
            field_id=field_id
        ).first()
        if existing:
            existing.value = value
        else:
            pv = PatientValue(
                patient_id=patient.id,
                field_id=field_id,
                value=value
            )
            db.session.add(pv)

    db.session.commit()
    return jsonify({'success': True, 'patient_id': patient.id})


@app.route('/api/patient/<int:patient_id>', methods=['GET'])
def get_patient(patient_id):
    patient = Patient.query.get(patient_id)
    if not patient:
        return jsonify({'error': '患者不存在'}), 404

    values = {}
    for v in patient.values:
        values[v.field_id] = v.value

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

    patients = Patient.query.filter_by(project_id=project_id).order_by(Patient.id.asc()).all()
    fields = sorted(project.fields, key=lambda x: x.sort_order)

    result = []
    for p in patients:
        values = {}
        for v in p.values:
            values[v.field_id] = v.value
        result.append({
            'id': p.id,
            'patient_code': p.patient_code,
            'hospital_id': p.hospital_id,
            'name': p.name,
            'gender': p.gender,
            'age': p.age,
            'has_data': len([v for v in values.values() if v and v.strip()]) > 0,
            'values': values
        })
    return jsonify({
        'patients': result,
        'fields': [{'id': f.id, 'name': f.field_name} for f in fields]
    })


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

    header = ['住院号', '住院日期', '出院日期', '其他合并症'] + [f.field_name + (f' ({f.field_unit})' if f.field_unit else '') for f in fields]
    ws.append(header)

    for p in patients:
        row = [p.hospital_id or '', p.admission_date or '', p.discharge_date or '', p.other_complications or '']
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


# ========== 页面路由 ==========

@app.route('/')
def index():
    return send_from_directory('templates', 'projects.html')


@app.route('/projects')
def projects_page():
    return send_from_directory('templates', 'projects.html')


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


@app.route('/ocr')
def ocr_page():
    return send_from_directory('templates', 'index.html')


if __name__ == '__main__':
    print("OCR API starting...")
    app.run(host='0.0.0.0', port=5000, debug=False)