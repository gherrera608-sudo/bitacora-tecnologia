import csv
from io import StringIO
from flask import Response

@app.route('/descargar_reporte')
def descargar_reporte():
    # Obtener filtros si se aplican
    asignatura_filtro = request.args.get('asignatura', '')
    fecha_inicio = request.args.get('fecha_inicio', '')
    fecha_fin = request.args.get('fecha_fin', '')
    
    query = Prestamo.query
    if asignatura_filtro:
        query = query.filter_by(asignatura=asignatura_filtro)
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

    # Crear archivo CSV en memoria
    si = StringIO()
    cw = csv.writer(si)
    
    # Encabezados de la tabla Excel
    cw.writerow([
        'ID', 'Fecha Entrega', 'Docente', 'Asignatura/Especialidad', 
        'N° Inventario', 'Equipo', 'Cédula Estudiante', 'Nombre Estudiante', 
        'Sección', 'Incluye Cargador', 'En Caja', 'Estado Inicial', 
        'Confirmado por Estudiante', 'Fecha Confirmación', 'Reporte / Observaciones'
    ])

    for p in prestamos:
        cw.writerow([
            p.id,
            p.fecha_entrega.strftime('%Y-%m-%d %H:%M') if p.fecha_entrega else '',
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
            p.fecha_reporte.strftime('%Y-%m-%d %H:%M') if p.fecha_reporte else '',
            p.reporte_estudiante if p.reporte_estudiante else 'Sin novedades'
        ])

    output = Response(si.getvalue(), mimetype='text/csv')
    output.headers["Content-Disposition"] = f"attachment; filename=Reporte_Prestamos_{datetime.now().strftime('%Y_%m_%d')}.csv"
    return output
