import os
import re
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from flask_sqlalchemy import SQLAlchemy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'spet-clave-secret-2026')

# Configuración para Neon (PostgreSQL)
db_uri = os.environ.get('DATABASE_URL')
if db_uri and db_uri.startswith("postgres://"):
    db_uri = db_uri.replace("postgres://", "postgresql://", 1)
app.config['SQLALCHEMY_DATABASE_URI'] = db_uri or 'sqlite:///bitacora.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

def hora_cr():
    return datetime.now(ZoneInfo("America/Costa_Rica")).replace(tzinfo=None)

class Prestamo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    inventario = db.Column(db.String(50), nullable=False)
    marca_modelo = db.Column(db.String(100), nullable=False)
    cedula = db.Column(db.String(30), nullable=False)
    estudiante = db.Column(db.String(100), nullable=False)
    grupo = db.Column(db.String(50), nullable=False)
    profesor = db.Column(db.String(100), default="Docente")
    asignatura = db.Column(db.String(100), default="General")
    cargador_entregado = db.Column(db.String(10), default="Sí")
    caja_entregada = db.Column(db.String(10), default="Sí")
    estado_inicial = db.Column(db.String(200), default="Excelente estado")
    fecha_entrega = db.Column(db.DateTime, default=hora_cr)
    
    confirmado_estudiante = db.Column(db.Boolean, default=False)
    cargador_recibido = db.Column(db.String(10), nullable=True)
    reporte_estudiante = db.Column(db.Text, nullable=True)
    novedad_detectada = db.Column(db.Boolean, default=False)
    fecha_reporte = db.Column(db.DateTime, nullable=True)

with app.app_context():
    db.create_all()

@app.template_filter('strftime')
def _jinja2_filter_datetime(date, fmt=None):
    if date is None:
        return "-"
    return date.strftime(fmt or '%d/%m/%Y %H:%M')

def limpiar_cedula(val):
    return re.sub(r'\D', '', val or '')

def formatear_inventario(val):
    digitos = re.sub(r'\D', '', val or '')
    if len(digitos) == 4:
        return f"3982-{digitos}"
    elif len(digitos) == 8 and digitos.startswith('3982'):
        return f"3982-{digitos[4:]}"
    return val.strip().upper()

@app.route('/')
@app.route('/direccion')
def vista_direccion():
    query = Prestamo.query
    asignatura_filtro = request.args.get('asignatura', '')
    grupo_filtro = request.args.get('grupo', '')
    fecha_inicio = request.args.get('fecha_inicio', '')
    fecha_fin = request.args.get('fecha_fin', '')
    
    if asignatura_filtro: query = query.filter_by(asignatura=asignatura_filtro)
    if grupo_filtro: query = query.filter_by(grupo=grupo_filtro)
    if fecha_inicio:
        try: query = query.filter(Prestamo.fecha_entrega >= datetime.strptime(fecha_inicio, '%Y-%m-%d'))
        except: pass
    if fecha_fin:
        try: query = query.filter(Prestamo.fecha_entrega <= datetime.strptime(fecha_fin + ' 23:59:59', '%Y-%m-%d %H:%M:%S'))
        except: pass
            
    prestamos = query.order_by(Prestamo.id.desc()).all()
    todas_asignaturas = [a[0] for a in db.session.query(Prestamo.asignatura).distinct().all() if a[0]]
    todos_grupos = [g[0] for g in db.session.query(Prestamo.grupo).distinct().all() if g[0]]
    
    return render_template('direccion.html', prestamos=prestamos, total=len(prestamos), 
                           novedades=sum(1 for p in prestamos if p.novedad_detectada), 
                           confirmados=sum(1 for p in prestamos if p.confirmado_estudiante), 
                           todas_asignaturas=todas_asignaturas, todos_grupos=todos_grupos,
                           asignatura_filtro=asignatura_filtro, grupo_filtro=grupo_filtro,
                           fecha_inicio=fecha_inicio, fecha_fin=fecha_fin, ahora=hora_cr())

@app.route('/descargar_reporte')
def descargar_reporte():
    prestamos = Prestamo.query.order_by(Prestamo.id.desc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['ID', 'Fecha Entrega', 'Docente', 'Asignatura', 'Inventario', 'Equipo', 'Cédula', 'Estudiante', 'Grupo'])
    for p in prestamos:
        ws.append([p.id, p.fecha_entrega.strftime('%d/%m/%Y %H:%M'), p.profesor, p.asignatura, p.inventario, p.marca_modelo, p.cedula, p.estudiante, p.grupo])
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name="Reporte.xlsx")

@app.route('/profesor', methods=['GET', 'POST'])
def vista_profesor():
    if request.method == 'POST':
        nuevo_prestamo = Prestamo(
            inventario=formatear_inventario(request.form.get('inventario', '')),
            marca_modelo=request.form['marca_modelo'][:100],
            cedula=limpiar_cedula(request.form.get('cedula', '')),
            estudiante=request.form['estudiante'][:100],
            grupo=request.form['grupo'],
            profesor=request.form['profesor'][:100],
            asignatura=request.form['asignatura'],
            cargador_entregado=request.form.get('cargador', 'No'),
            caja_entregada=request.form.get('caja', 'No'),
            estado_inicial=request.form['estado_inicial'][:200]
        )
        db.session.add(nuevo_prestamo)
        db.session.commit()
        flash("¡Registrado!", "success")
        return redirect(url_for('vista_profesor'))
    return render_template('profesor.html')

@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar_prestamo(id):
    prestamo = Prestamo.query.get_or_404(id)
    if request.method == 'POST':
        prestamo.inventario = formatear_inventario(request.form.get('inventario', ''))
        prestamo.grupo = request.form['grupo']
        db.session.commit()
        return redirect(url_for('vista_direccion'))
    return render_template('editar.html', prestamo=prestamo)

@app.route('/eliminar/<int:id>', methods=['POST'])
def eliminar_prestamo(id):
    db.session.delete(Prestamo.query.get_or_404(id))
    db.session.commit()
    return redirect(request.referrer or url_for('vista_direccion'))

@app.route('/estudiante', methods=['GET', 'POST'])
def vista_estudiante():
    if request.method == 'POST':
        inv_final = formatear_inventario(request.form.get('inventario', ''))
        cedula_ingresada = limpiar_cedula(request.form.get('cedula', ''))
        
        # --- SEGURIDAD PASO 1: Validación obligatoria de inventario + cédula ---
        prestamo = Prestamo.query.filter_by(inventario=inv_final, cedula=cedula_ingresada).order_by(Prestamo.id.desc()).first()
        
        if not prestamo:
            flash("Error: Inventario o Cédula incorrectos.", "danger")
            return redirect(url_for('vista_estudiante'))
            
        prestamo.confirmado_estudiante = True
        prestamo.cargador_recibido = request.form.get('cargador_recibido', 'No')
        prestamo.reporte_estudiante = request.form.get('reporte', '')[:300]
        prestamo.fecha_reporte = hora_cr()
        prestamo.novedad_detectada = True if (len(prestamo.reporte_estudiante) > 0 or prestamo.cargador_recibido == 'No') else False
        db.session.commit()
        return render_template('confirmacion.html', prestamo=prestamo)
        
    return render_template('estudiante.html')

if __name__ == '__main__':
    app.run(debug=True)
