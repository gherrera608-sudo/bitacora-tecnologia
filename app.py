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
    
    if asignatura_filtro:
        query = query.filter_by(asignatura=asignatura_filtro)
        
    if grupo_filtro:
        query = query.filter_by(grupo=grupo_filtro)
        
    if fecha_inicio:
        try:
            dt_inicio = datetime.strptime(fecha_inicio, '%Y-%m-%d')
            query = query.filter(Prestamo.fecha_entrega >= dt_inicio)
        except ValueError:
            pass
            
    if fecha_fin:
        try:
            dt_fin = datetime.strptime(fecha_fin + ' 23:59:59', '%Y-%m-%d %H:%M:%S')
            query = query.filter(Prestamo.fecha_entrega <= dt_fin)
        except ValueError:
            pass
            
    prestamos = query.order_by(Prestamo.id.desc()).all()
    todas_asignaturas = [a[0] for a in db.session.query(Prestamo.asignatura).distinct().all() if a[0]]
    todos_grupos = [g[0] for g in db.session.query(Prestamo.grupo).distinct().all() if g[0]]
    
    total = len(prestamos)
    novedades = sum(1 for p in prestamos if p.novedad_detectada)
    confirmados = sum(1 for p in prestamos if p.confirmado_estudiante)
    
    return render_template('direccion.html', prestamos=prestamos, total=total, 
                           novedades=novedades, confirmados=confirmados, 
                           todas_asignaturas=todas_asignaturas,
                           todos_grupos=todos_grupos,
                           asignatura_filtro=asignatura_filtro,
                           grupo_filtro=grupo_filtro,
                           fecha_inicio=fecha_inicio, fecha_fin=fecha_fin,
                           ahora=hora_cr())

@app.route('/descargar_reporte')
def descargar_reporte():
    asignatura_filtro = request.args.get('asignatura', '')
    grupo_filtro = request.args.get('grupo', '')
    fecha_inicio = request.args.get('fecha_inicio', '')
    fecha_fin = request.args.get('fecha_fin', '')
    
    query = Prestamo.query
    if asignatura_filtro:
        query = query.filter_by(asignatura=asignatura_filtro)
    if grupo_filtro:
        query = query.filter_by(grupo=grupo_filtro)
    if fecha_inicio:
        try:
            dt_inicio = datetime.strptime(fecha_inicio, '%Y-%m-%d')
            query = query.filter(Prestamo.fecha_entrega >= dt_inicio)
        except ValueError:
            pass
    if fecha_fin:
        try:
            dt_fin = datetime.strptime(fecha_fin + ' 23:59:59', '%Y-%m-%d %H:%M:%S')
            query = query.filter(Prestamo.fecha_entrega <= dt_fin)
        except ValueError:
            pass

    prestamos = query.order_by(Prestamo.id.desc()).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte Préstamos"

    headers = [
        'ID', 'Fecha Entrega', 'Docente', 'Asignatura/Especialidad', 
        'N° Inventario', 'Equipo', 'Cédula Estudiante', 'Nombre Estudiante', 
        'Sección', 'Incluye Cargador', 'En Caja', 'Estado Inicial', 
        'Confirmado Estudiante', 'Fecha Confirmación', 'Reporte / Observaciones'
    ]
    ws.append(headers)

    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="990000", end_color="990000", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for p in prestamos:
        row = [
            p.id,
            p.fecha_entrega.strftime('%d/%m/%Y %H:%M') if p.fecha_entrega else '',
            p.profesor,
            p.asignatura,
            p.inventario,
            p.marca_modelo,
            p.cedula,
            p.estudiante,
            p.grupo,
            p.cargador_entregado,
            p.caja_entregada,
            p.estado_inicial,
            'Sí' if p.confirmado_estudiante else 'No',
            p.fecha_reporte.strftime('%d/%m/%Y %H:%M') if p.fecha_reporte else '',
            p.reporte_estudiante if p.reporte_estudiante else 'Sin novedades'
        ]
        ws.append(row)

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    nombre_archivo = f"Reporte_Prestamos_{hora_cr().strftime('%Y_%m_%d')}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=nombre_archivo
    )

@app.route('/profesor', methods=['GET', 'POST'])
def vista_profesor():
    if request.method == 'POST':
        inv_input = request.form.get('inventario', '').strip()
        inventario_final = formatear_inventario(inv_input)
        
        cedula_raw = request.form.get('cedula', '')
        cedula_clean = limpiar_cedula(cedula_raw)

        if not (8 <= len(cedula_clean) <= 12):
            flash("La cédula debe contener solo números (entre 8 y 12 dígitos sin guiones).", "error")
            return redirect(url_for('vista_profesor'))

        nuevo_prestamo = Prestamo(
            inventario=inventario_final,
            marca_modelo=request.form['marca_modelo'].strip()[:100],
            cedula=cedula_clean,
            estudiante=request.form['estudiante'].strip()[:100],
            grupo=request.form['grupo'].strip(),
            profesor=request.form['profesor'].strip()[:100],
            asignatura=request.form['asignatura'].strip(),
            cargador_entregado=request.form.get('cargador', 'No'),
            caja_entregada=request.form.get('caja', 'No'),
            estado_inicial=request.form['estado_inicial'].strip()[:200],
            fecha_entrega=hora_cr()
        )
        db.session.add(nuevo_prestamo)
        db.session.commit()
        
        flash("¡Préstamo registrado con éxito! Puede registrar otro equipo.", "success")
        return redirect(url_for('vista_profesor'))
        
    return render_template('profesor.html')

@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar_prestamo(id):
    prestamo = Prestamo.query.get_or_404(id)
    if request.method == 'POST':
        inv_input = request.form.get('inventario', '').strip()
        inventario_final = formatear_inventario(inv_input)
        
        cedula_raw = request.form.get('cedula', '')
        cedula_clean = limpiar_cedula(cedula_raw)

        if not (8 <= len(cedula_clean) <= 12):
            flash("La cédula debe contener solo números (entre 8 y 12 dígitos).", "error")
            return redirect(url_for('editar_prestamo', id=id))

        prestamo.inventario = inventario_final
        prestamo.marca_modelo = request.form['marca_modelo'].strip()[:100]
        prestamo.cedula = cedula_clean
        prestamo.estudiante = request.form['estudiante'].strip()[:100]
        prestamo.grupo = request.form['grupo'].strip()
        prestamo.profesor = request.form['profesor'].strip()[:100]
        prestamo.asignatura = request.form['asignatura'].strip()
        prestamo.cargador_entregado = request.form.get('cargador', 'No')
        prestamo.caja_entregada = request.form.get('caja', 'No')
        prestamo.estado_inicial = request.form['estado_inicial'].strip()[:200]
        db.session.commit()
        return redirect(url_for('vista_direccion'))
    return render_template('editar.html', prestamo=prestamo)

@app.route('/eliminar/<int:id>', methods=['POST'])
def eliminar_prestamo(id):
    prestamo = Prestamo.query.get_or_404(id)
    db.session.delete(prestamo)
    db.session.commit()
    return redirect(url_for('vista_direccion'))

@app.route('/estudiante', methods=['GET', 'POST'])
@app.route('/reportar', methods=['GET', 'POST'])
def vista_estudiante():
    inv_query_raw = request.args.get('inv', '').strip()
    inv_query = formatear_inventario(inv_query_raw) if inv_query_raw else ''
    
    if request.method == 'POST':
        inv_post = request.form.get('inventario', '').strip()
        inv_final = formatear_inventario(inv_post)
        
        prestamo = Prestamo.query.filter_by(inventario=inv_final).order_by(Prestamo.id.desc()).first()
        if not prestamo:
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
        prestamo_actual = Prestamo.query.filter_by(inventario=inv_query).order_by(Prestamo.id.desc()).first()
    return render_template('estudiante.html', inv_query=inv_query_raw, prestamo=prestamo_actual)

@app.route('/importar_datos')
def importar_datos():
    archivo = 'respaldo_datos.xlsx'
    if not os.path.exists(archivo):
        return "Archivo no encontrado. Asegúrate de subir 'respaldo_datos.xlsx' a la carpeta principal del proyecto."

    wb = openpyxl.load_workbook(archivo)
    ws = wb.active
    
    contador = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        fecha = row[1] if isinstance(row[1], datetime) else datetime.strptime(str(row[1]), '%d/%m/%Y %H:%M')
        
        nuevo = Prestamo(
            fecha_entrega=fecha,
            profesor=str(row[2]),
            asignatura=str(row[3]),
            inventario=str(row[4]),
            marca_modelo=str(row[5]),
            cedula=str(row[6]),
            estudiante=str(row[7]),
            grupo=str(row[8]),
            cargador_entregado=str(row[9]),
            caja_entregada=str(row[10]),
            estado_inicial=str(row[11]),
            confirmado_estudiante=True if row[12] == 'Sí' else False,
            reporte_estudiante=str(row[14]) if row[14] else ""
        )
        db.session.add(nuevo)
        contador += 1
    
    db.session.commit()
    return f"¡Éxito! Se han importado {contador} registros a la base de datos de Neon."

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
