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
    asignatura_filtro = request.args.get('asignatura', '')
    grupo_filtro = request.args.get('grupo', '')
    fecha_inicio = request.args.get('fecha_inicio', '')
    fecha_fin = request.args.get('fecha_fin', '')
    
    query = Prestamo.query
    if asignatura_filtro: query = query.filter_by(asignatura=asignatura_filtro)
    if grupo_filtro: query = query.filter_by(grupo=grupo_filtro)
    if fecha_inicio:
        try: query = query.filter(Prestamo.fecha_entrega >= datetime.strptime(fecha_inicio, '%Y-%m-%d'))
        except: pass
    if fecha_fin:
        try: query = query.filter(Prestamo.fecha_entrega <= datetime.strptime(fecha_fin + ' 23:59:59', '%Y-%m-%d %H:%M:%S'))
        except: pass

    prestamos = query.order_by(Prestamo.id.desc()).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte Préstamos"
    ws.append(['ID', 'Fecha', 'Docente', 'Asignatura', 'Inventario', 'Equipo', 'Cédula', 'Estudiante', 'Sección', 'Cargador', 'Caja', 'Estado', 'Confirmado', 'Fecha Conf.', 'Obs'])
    
    for p in prestamos:
        ws.append([p.id, p.fecha_entrega.strftime('%d/%m/%Y %H:%M') if p.fecha_entrega else '', p.profesor, p.asignatura, p.inventario, p.marca_modelo, p.cedula, p.estudiante, p.grupo, p.cargador_entregado, p.caja_entregada, p.estado_inicial, 'Sí' if p.confirmado_estudiante else 'No', p.fecha_reporte.strftime('%d/%m/%Y %H:%M') if p.fecha_reporte else '', p.reporte_estudiante])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f"Reporte_{datetime.now().strftime('%Y%m%d')}.xlsx")

@app.route('/profesor', methods=['GET', 'POST'])
def vista_profesor():
    if request.method == 'POST':
        inventario_raw = request.form.get('inventario', '').strip()
        marca_modelo = request.form.get('marca_modelo', '').strip()
        cedula = limpiar_cedula(request.form.get('cedula', ''))
        estudiante = request.form.get('estudiante', '').strip()
        grupo = request.form.get('grupo', '').strip()
        
        if not inventario_raw or not marca_modelo or not cedula or not estudiante or not grupo:
            flash("Error: Todos los campos obligatorios deben estar completos.", "danger")
            return redirect(url_for('vista_profesor'))

        nuevo_prestamo = Prestamo(
            inventario=formatear_inventario(inventario_raw),
            marca_modelo=marca_modelo[:100],
            cedula=cedula,
            estudiante=estudiante[:100],
            grupo=grupo,
            profesor=request.form.get('profesor', 'Docente')[:100],
            asignatura=request.form.get('asignatura', 'General'),
            cargador_entregado=request.form.get('cargador', 'No'),
            caja_entregada=request.form.get('caja', 'No'),
            estado_inicial=request.form.get('estado_inicial', 'Excelente estado')[:200],
            fecha_entrega=hora_cr()
        )
        db.session.add(nuevo_prestamo)
        db.session.commit()
        flash("¡Préstamo registrado exitosamente!", "success")
        return redirect(url_for('vista_profesor'))
    return render_template('profesor.html')

@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar_prestamo(id):
    prestamo = Prestamo.query.get_or_404(id)
    if request.method == 'POST':
        inv_edit = request.form.get('inventario', '').strip()
        marca_edit = request.form.get('marca_modelo', '').strip()
        cedula_edit = limpiar_cedula(request.form.get('cedula', ''))
        estudiante_edit = request.form.get('estudiante', '').strip()
        grupo_edit = request.form.get('grupo', '').strip()
        asignatura_edit = request.form.get('asignatura', '').strip()
        
        if not inv_edit or not marca_edit or not cedula_edit or not estudiante_edit or not grupo_edit:
            flash("Error: Todos los campos obligatorios deben estar completos.", "danger")
            return redirect(url_for('editar_prestamo', id=id))
            
        prestamo.inventario = formatear_inventario(inv_edit)
        prestamo.marca_modelo = marca_edit[:100]
        prestamo.cedula = cedula_edit
        prestamo.estudiante = estudiante_edit[:100]
        prestamo.grupo = grupo_edit
        if asignatura_edit:
            prestamo.asignatura = asignatura_edit
            
        db.session.commit()
        flash("¡Registro actualizado correctamente!", "success")
        return redirect(url_for('vista_direccion'))
        
    return render_template('editar.html', prestamo=prestamo)

@app.route('/eliminar/<int:id>', methods=['POST'])
def eliminar_prestamo(id):
    prestamo = Prestamo.query.get_or_404(id)
    db.session.delete(prestamo)
    db.session.commit()
    flash("Registro eliminado correctamente.", "success")
    return redirect(request.referrer or url_for('vista_direccion'))

@app.route('/estudiante', methods=['GET', 'POST'])
@app.route('/reportar', methods=['GET', 'POST'])
def vista_estudiante():
    inv_query_raw = request.args.get('inv', '').strip()
    inv_query = formatear_inventario(inv_query_raw) if inv_query_raw else ''
    
    if request.method == 'POST':
        inv_post = request.form.get('inventario', '').strip()
        inv_final = formatear_inventario(inv_post)
        cedula_ingresada = limpiar_cedula(request.form.get('cedula', ''))
        
        if not inv_post or not cedula_ingresada:
            flash("Por favor ingrese el número de inventario y su cédula.", "danger")
            return redirect(url_for('vista_estudiante'))

        prestamo = Prestamo.query.filter_by(
            inventario=inv_final, 
            cedula=cedula_ingresada, 
            confirmado_estudiante=False
        ).first()
        
        if not prestamo:
            flash("Error: El inventario no tiene préstamos pendientes para esta cédula o los datos son incorrectos.", "danger")
            return redirect(url_for('vista_estudiante', inv=inv_post))
            
        prestamo.confirmado_estudiante = True
        prestamo.cargador_recibido = request.form.get('cargador_recibido', 'No')
        reporte = request.form.get('reporte', '').strip()[:300]
        prestamo.reporte_estudiante = reporte
        prestamo.fecha_reporte = hora_cr()
        prestamo.novedad_detectada = True if (len(reporte) > 0 or prestamo.cargador_recibido == 'No') else False
            
        db.session.commit()
        return render_template('confirmacion.html', prestamo=prestamo)
        
    prestamo_actual = None
    if inv_query:
        prestamo_actual = Prestamo.query.filter_by(
            inventario=inv_query, 
            confirmado_estudiante=False
        ).order_by(Prestamo.id.desc()).first()
        
    return render_template('estudiante.html', inv_query=inv_query_raw, prestamo=prestamo_actual)

if __name__ == '__main__':
    app.run(debug=True)

if __name__ == '__main__':
    app.run(debug=True)
