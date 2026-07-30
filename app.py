import os
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'bitacora-clave-secret-2026')

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///bitacora.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

class Prestamo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    inventario = db.Column(db.String(50), nullable=False)
    marca_modelo = db.Column(db.String(100), nullable=False)
    cedula = db.Column(db.String(30), nullable=False)
    estudiante = db.Column(db.String(100), nullable=False)
    grupo = db.Column(db.String(50), nullable=False)
    cargador_entregado = db.Column(db.String(10), default="Sí")
    caja_entregada = db.Column(db.String(10), default="Sí")
    estado_inicial = db.Column(db.String(200), default="Excelente estado")
    fecha_entrega = db.Column(db.DateTime, default=datetime.now)
    
    confirmado_estudiante = db.Column(db.Boolean, default=False)
    cargador_recibido = db.Column(db.String(10), nullable=True)
    reporte_estudiante = db.Column(db.Text, nullable=True)
    novedad_detectada = db.Column(db.Boolean, default=False)
    fecha_reporte = db.Column(db.DateTime, nullable=True)

with app.app_context():
    # Elimina las tablas antiguas si tienen un esquema viejo para evitar el error 500
    db.drop_all()
    db.create_all()

@app.route('/')
@app.route('/direccion')
def vista_direccion():
    prestamos = Prestamo.query.order_by(Prestamo.id.desc()).all()
    total = len(prestamos)
    novedades = sum(1 for p in prestamos if p.novedad_detectada)
    confirmados = sum(1 for p in prestamos if p.confirmado_estudiante)
    return render_template('direccion.html', prestamos=prestamos, total=total, novedades=novedades, confirmados=confirmados)

@app.route('/profesor', methods=['GET', 'POST'])
def vista_profesor():
    if request.method == 'POST':
        nuevo_prestamo = Prestamo(
            inventario=request.form['inventario'].strip().upper(),
            marca_modelo=request.form['marca_modelo'].strip(),
            cedula=request.form['cedula'].strip(),
            estudiante=request.form['estudiante'].strip(),
            grupo=request.form['grupo'].strip(),
            cargador_entregado=request.form.get('cargador', 'No'),
            caja_entregada=request.form.get('caja', 'No'),
            estado_inicial=request.form['estado_inicial'].strip(),
            fecha_entrega=datetime.now()
        )
        db.session.add(nuevo_prestamo)
        db.session.commit()
        return redirect(url_for('vista_direccion'))
    return render_template('profesor.html')

@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar_prestamo(id):
    prestamo = Prestamo.query.get_or_404(id)
    if request.method == 'POST':
        prestamo.inventario = request.form['inventario'].strip().upper()
        prestamo.marca_modelo = request.form['marca_modelo'].strip()
        prestamo.cedula = request.form['cedula'].strip()
        prestamo.estudiante = request.form['estudiante'].strip()
        prestamo.grupo = request.form['grupo'].strip()
        prestamo.cargador_entregado = request.form.get('cargador', 'No')
        prestamo.caja_entregada = request.form.get('caja', 'No')
        prestamo.estado_inicial = request.form['estado_inicial'].strip()
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
    inv_query = request.args.get('inv', '').upper()
    if request.method == 'POST':
        inv = request.form['inventario'].strip().upper()
        prestamo = Prestamo.query.filter_by(inventario=inv).order_by(Prestamo.id.desc()).first()
        if not prestamo:
            return redirect(url_for('vista_estudiante', inv=inv))
            
        prestamo.confirmado_estudiante = True
        prestamo.cargador_recibido = request.form.get('cargador_recibido', 'No')
        reporte = request.form.get('reporte', '').strip()
        prestamo.reporte_estudiante = reporte
        prestamo.fecha_reporte = datetime.now()
        prestamo.novedad_detectada = True if (len(reporte) > 0 or prestamo.cargador_recibido == 'No') else False
            
        db.session.commit()
        return render_template('confirmacion.html', prestamo=prestamo)
        
    prestamo_actual = None
    if inv_query:
        prestamo_actual = Prestamo.query.filter_by(inventario=inv_query).order_by(Prestamo.id.desc()).first()
    return render_template('estudiante.html', inv_query=inv_query, prestamo=prestamo_actual)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
