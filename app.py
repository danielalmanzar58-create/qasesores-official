import os
import io
from datetime import date, datetime
from flask import Flask, render_template, redirect, url_for, request, flash, send_file, abort
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'qasesores_global_enterprise_key_2026')
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('SQLALCHEMY_DATABASE_URI', 'sqlite:///qasesores.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Por favor autentíquese para acceder al sistema oficial Qasesores.'
login_manager.login_message_category = 'error'

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(30), nullable=False, default='supervisor') # 'supervisor' o 'jefe'
    is_admin = db.Column(db.Boolean, default=False)

class WorkRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    work_date = db.Column(db.String(20), nullable=False)
    location_type = db.Column(db.String(50), nullable=False) # 'Nuevo Local' / 'Viejo Local'
    description = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(20), nullable=False) # 'Pendiente' / 'Finalizado'
    observations = db.Column(db.Text, nullable=True)
    worker_name = db.Column(db.String(100), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)

class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)
    username = db.Column(db.String(80), nullable=False)
    action = db.Column(db.String(50), nullable=False) # 'CREACIÓN', 'EDICIÓN', 'ELIMINACIÓN'
    details = db.Column(db.Text, nullable=False)

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))

@app.context_processor
def inject_global_vars():
    logo_path = os.path.join(app.static_folder, 'logo.png')
    return {
        'logo_exists': os.path.exists(logo_path)
    }

@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        
        if user and check_password_hash(user.password, password):
            login_user(user)
            flash(f'¡Acceso autorizado! Bienvenido/a al sistema oficial Qasesores, {user.username}.', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('Credenciales incorrectas o acceso no autorizado.', 'error')
            
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Sesión finalizada correctamente.', 'success')
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    works = WorkRecord.query.order_by(WorkRecord.created_at.desc()).all()
    return render_template('dashboard.html', works=works)

@app.route('/archivos')
@login_required
def archivos():
    works = WorkRecord.query.order_by(WorkRecord.created_at.desc()).all()
    return render_template('archivos.html', works=works)

@app.route('/movimientos')
@login_required
def movimientos():
    if current_user.role != 'jefe':
        abort(403)
    logs = AuditLog.query.order_by(AuditLog.timestamp.desc()).all()
    total_works = WorkRecord.query.count()
    return render_template('movimientos.html', logs=logs, total_works=total_works)

@app.route('/exportar-excel')
@login_required
def export_excel():
    works = WorkRecord.query.order_by(WorkRecord.created_at.desc()).all()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trabajos Oficiales"
    
    header_fill = PatternFill(start_color="0A2540", end_color="0A2540", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    border_side = Side(border_style="thin", color="CBD5E1")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    headers = ["ID", "Fecha", "Tipo de Local", "Descripción del Trabajo", "Estado", "Observación", "Realizado Por"]
    ws.append(headers)
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for w in works:
        ws.append([
            w.id,
            w.work_date,
            w.location_type,
            w.description,
            w.status,
            w.observations or "",
            w.worker_name
        ])
        
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column in [1, 2, 3, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"QASESORES_Official_Database_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

@app.route('/work/new', methods=['GET', 'POST'])
@login_required
def work_form():
    if request.method == 'POST':
        work_date = request.form.get('work_date')
        location_type = request.form.get('location_type')
        description = request.form.get('description')
        status = request.form.get('status')
        observations = request.form.get('observations')
        worker_name = request.form.get('worker_name')
        
        new_work = WorkRecord(
            work_date=work_date,
            location_type=location_type,
            description=description,
            status=status,
            observations=observations,
            worker_name=worker_name,
            user_id=current_user.id
        )
        db.session.add(new_work)
        
        # Registrar auditoría
        audit = AuditLog(
            username=current_user.username,
            action="CREACIÓN",
            details=f"Creó trabajo en {location_type} - Estado: {status} por {worker_name}"
        )
        db.session.add(audit)
        
        db.session.commit()
        flash('Trabajo registrado y movimiento auditado exitosamente.', 'success')
        return redirect(url_for('dashboard'))
        
    today_date = date.today().isoformat()
    return render_template('work_form.html', title='Registro Oficial de Trabajo', work=None, today_date=today_date)

@app.route('/work/<int:work_id>/edit', methods=['GET', 'POST'])
@login_required
def edit_work(work_id):
    work = WorkRecord.query.get_or_404(work_id)
    if request.method == 'POST':
        old_status = work.status
        work.work_date = request.form.get('work_date')
        work.location_type = request.form.get('location_type')
        work.description = request.form.get('description')
        work.status = request.form.get('status')
        work.observations = request.form.get('observations')
        work.worker_name = request.form.get('worker_name')
        
        # Registrar auditoría de edición
        audit = AuditLog(
            username=current_user.username,
            action="EDICIÓN",
            details=f"Modificó el trabajo #{work.id} ({work.location_type}). Estado anterior: {old_status} -> Nuevo estado: {work.status}"
        )
        db.session.add(audit)
        db.session.commit()
        
        flash('Trabajo actualizado y registrado en movimientos.', 'success')
        return redirect(url_for('dashboard'))
        
    return render_template('work_form.html', title='Editar Registro de Trabajo', work=work)

@app.route('/users')
@login_required
def users():
    if not current_user.is_admin and current_user.role != 'jefe':
        abort(403)
    all_users = User.query.all()
    return render_template('users.html', users=all_users)

@app.route('/users/delete/<int:user_id>', methods=['POST'])
@login_required
def delete_user(user_id):
    if not current_user.is_admin and current_user.role != 'jefe':
        abort(403)
    
    user_to_delete = User.query.get_or_404(user_id)
    if user_to_delete.username == 'JefaturaGlobal':
        flash('No se puede revocar al usuario principal de la Jefatura.', 'error')
        return redirect(url_for('users'))
        
    db.session.delete(user_to_delete)
    db.session.commit()
    flash('Acceso de usuario revocado correctamente.', 'success')
    return redirect(url_for('users'))

@app.errorhandler(403)
def forbidden(e):
    return render_template('error.html', error_message='Acceso denegado. Área exclusiva de la Jefatura Departamental.'), 403

@app.errorhandler(404)
def page_not_found(e):
    return render_template('error.html', error_message='Página no encontrada (404)'), 404

@app.errorhandler(500)
def internal_server_error(e):
    return render_template('error.html', error_message='Error interno del servidor (500)'), 500

def init_db():
    with app.app_context():
        db.create_all()
        # 1. Crear a Diosaly como Supervisor
        if not User.query.filter_by(username='Diosaly').first():
            hashed_pw = generate_password_hash('Aguacate 001', method='pbkdf2:sha256')
            supervisor_user = User(username='Diosaly', password=hashed_pw, role='supervisor', is_admin=False)
            db.session.add(supervisor_user)
            db.session.commit()
        
        # 2. Crear a la Jefatura / Dirección con credenciales exclusivas
        if not User.query.filter_by(username='JefaturaGlobal').first():
            hashed_jefe = generate_password_hash('QasesoresMaster2026', method='pbkdf2:sha256')
            jefe_user = User(username='JefaturaGlobal', password=hashed_jefe, role='jefe', is_admin=True)
            db.session.add(jefe_user)
            db.session.commit()

if __name__ == '__main__':
    init_db()
    app.run(debug=True, host='0.0.0.0', port=5000)
