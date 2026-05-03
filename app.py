"""
SIAB - Sistema Informático Automatizado de Bomberos
====================================================
App Flask principal - Etapa 2 v2
"""

from flask import Flask, render_template, request, redirect, url_for, send_file, session, flash, jsonify
import mysql.connector
from mysql.connector import Error
from functools import wraps
from datetime import datetime
import os
import pandas as pd  # Importante para el reporte Excel
from io import BytesIO
from database import get_db

# ReportLab para el PDF institucional
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import landscape, A4

from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, HRFlowable
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
from datetime import datetime
import os

# Definición de ruta base
base_dir = os.path.abspath(os.path.dirname(__file__))
from controladores.moviles import obtener_estado_unidades

ahora = datetime.now()
# Esto genera "/04/2026", que es como termina la fecha en tu DB
mes_busqueda = ahora.strftime('/%m/%Y')

app = Flask(__name__)
app.secret_key = "siab_bomberos_2026_secretkey"

DB_CONFIG = {
    'user': 'root',      # Generalmente 'root'
    'password': 'siab1234', # Tu contraseña de MySQL
    'database': 'SIAB'         # El nombre de tu base de datos
}

# Aqui iba la conexión a la base de datos

# ============================================================
# HELPERS
# ============================================================

def login_requerido(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if "usuario_id" not in session:
            flash("Debés iniciar sesión.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def rol_requerido(*roles):
    from functools import wraps
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if session.get("rol") not in roles:
                flash("No tenés permisos para acceder a esa sección.", "danger")
                return redirect(url_for("inicio"))
            return f(*args, **kwargs)
        return decorated
    return decorator

# Función ÚNICA para verificar las claves del escritorio
def verificar_password_siab(password_ingresada, hash_almacenado):
    import hashlib
    import hmac
    try:
        if ':' not in hash_almacenado:
            return False
        salt_hex, dk_hex = hash_almacenado.split(':')
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(dk_hex)
        derived = hashlib.pbkdf2_hmac('sha256', password_ingresada.encode('utf-8'), salt, 100000)
        return hmac.compare_digest(derived, expected)
    except Exception:
        return False

# ============================================================
# LOGIN / LOGOUT
# ============================================================

@app.route("/", methods=["GET", "POST"])
def login():
    if "usuario_id" in session:
        return redirect(url_for("inicio"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()

        conn = get_db()
        if not conn:
            flash("Error de conexión a la base de datos.", "danger")
            return render_template("login.html")

        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT u.*, l.nombre, l.apellido, l.grado, l.cargo
            FROM usuarios u
            LEFT JOIN legajos l ON u.legajo = l.legajo
            WHERE u.username = %s AND u.activo = 1
        """, (username,))
        usuario = cur.fetchone()
        conn.close()

        # --- LÓGICA DE VALIDACIÓN CORREGIDA ---
        es_valido = False
        if usuario:
            hash_db = usuario["password_hash"]
            
            # Intento 1: Formato de Werkzeug/Flask (el que generamos recién: scrypt, pbkdf2, etc.)
            from werkzeug.security import check_password_hash
            if hash_db.startswith(('scrypt:', 'pbkdf2:')):
                es_valido = check_password_hash(hash_db, password)
            
            # Intento 2: ¿Es el formato del escritorio (con : pero sin prefijo)?
            elif ":" in hash_db:
                es_valido = verificar_password_siab(password, hash_db)
            
            # Intento 3: ¿Es el formato del ADMIN (SHA256 simple)?
            else:
                import hashlib
                hash_intento = hashlib.sha256(password.encode()).hexdigest()
                es_valido = (hash_db == hash_intento)

        if es_valido:
            apellido = usuario.get("apellido") or ""
            nombre   = usuario.get("nombre") or ""
            nombre_completo = f"{apellido} {nombre}".strip() or username

            session["usuario_id"] = usuario["id"]
            session["username"]   = usuario["username"]
            session["rol"]        = usuario["rol"]
            session["legajo"]     = usuario["legajo"]
            session["nombre"]     = nombre_completo
            session["grado"]      = usuario.get("grado") or "BOMBERO" 
            session["cargo"]      = usuario.get("cargo") or ""

            flash(f"Bienvenido, {nombre_completo}!", "success")
            return redirect(url_for("inicio"))
        else:
            flash("Usuario o contraseña incorrectos.", "danger")

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    flash("Sesión cerrada.", "info")
    return redirect(url_for("login"))


# ============================================================
# INICIO Y REDIRECCIÓN
# ============================================================

@app.route("/")
def index():
    # Esta función SOLO decide a dónde ir apenas abrís el navegador
    if "usuario_id" in session:
        return redirect(url_for("inicio"))
    return redirect(url_for("login"))

@app.route("/inicio")
@login_requerido
def inicio():
    # Esta función SOLO carga las estadísticas de Bomberos Almafuerte
    conn = get_db()
    stats = {}
    borradores = []
    
    if conn:
        cur = conn.cursor(dictionary=True)

        # Bomberos Activos
        cur.execute("SELECT COUNT(*) as total FROM legajos WHERE situacion = 'ACTIVO'")
        stats["bomberos_activos"] = cur.fetchone()["total"]

        # Eventos del mes
        cur.execute("""SELECT COUNT(*) as total FROM eventos 
                       WHERE MONTH(fecha) = MONTH(CURDATE()) 
                       AND YEAR(fecha) = YEAR(CURDATE()) 
                       AND estado = 'FINALIZADO'""")
        stats["eventos_mes"] = cur.fetchone()["total"]

        # Asistencias del mes
        cur.execute("""SELECT COUNT(*) as total FROM asistencia a 
                       JOIN eventos e ON a.evento_id = e.id 
                       WHERE a.estado = 'PRESENTE' 
                       AND MONTH(e.fecha) = MONTH(CURDATE()) 
                       AND e.estado = 'FINALIZADO'""")
        stats["asistencias_mes"] = cur.fetchone()["total"]

        # Cursos del año
        cur.execute("""SELECT COUNT(*) as total FROM cursos 
                       WHERE YEAR(fecha_inicio) = YEAR(CURDATE())""")
        stats["cursos_anio"] = cur.fetchone()["total"]

        # Borradores abiertos
        cur.execute("""
            SELECT e.id, e.tipo, e.descripcion, e.fecha, e.hora_inicio, 
                   COUNT(a.id) as total, 
                   SUM(a.estado = 'PRESENTE') as presentes, 
                   e.fecha_creacion 
            FROM eventos e 
            LEFT JOIN asistencia a ON e.id = a.evento_id 
            WHERE e.estado = 'BORRADOR' 
            GROUP BY e.id 
            ORDER BY e.fecha_creacion DESC
        """)
        borradores = cur.fetchall()
        conn.close()

    return render_template("inicio.html", stats=stats, borradores=borradores)


# ============================================================
# ASISTENCIA
# ============================================================
@app.route("/asistencia/bomberos")
@login_requerido
def get_bomberos():
    depto_id = request.args.get("departamento_id")
    conn = get_db()
    if not conn:
        return jsonify([])
    
    cur = conn.cursor(dictionary=True)

    # Si hay un depto_id específico y no es "todos"
    if depto_id and depto_id != "" and depto_id != "todos":
        cur.execute("""
            SELECT l.legajo, l.apellido, l.nombre, l.grado, l.cargo
            FROM legajos l
            JOIN bombero_departamento bd ON l.legajo = bd.legajo
            WHERE l.situacion = 'ACTIVO'
              AND bd.departamento_id = %s
              AND bd.activo = 1
            ORDER BY l.apellido, l.nombre
   
        """, (depto_id,))
    else:
        # Si no se eligió departamento, trae a TODOS los activos
        cur.execute("""
            SELECT legajo, apellido, nombre, grado, cargo
            FROM legajos
            WHERE situacion = 'ACTIVO'
            ORDER BY apellido, nombre
    
        """)
    
    bomberos = cur.fetchall()
    conn.close()
    return jsonify(bomberos)

@app.route("/asistencia", methods=["GET", "POST"])
@login_requerido
@rol_requerido('ADMIN', 'JEFATURA')
def asistencia():
    conn = get_db()
    conceptos     = []
    departamentos = []
    bomberos      = []
    config_puntos = None
    
    if conn:
        cur = conn.cursor(dictionary=True)
        
        # 1. Traer conceptos y departamentos
        cur.execute("SELECT id, concepto FROM conceptos WHERE activo = 1 ORDER BY concepto")
        conceptos = cur.fetchall()
        cur.execute("SELECT id, nombre FROM departamentos WHERE activo = 1 ORDER BY nombre")
        departamentos = cur.fetchall()
        
        # 2. TRAER LOS BOMBEROS (Limitado a 5 para prueba o todos según necesites)
        cur.execute("""
            SELECT legajo, apellido, nombre, grado 
            FROM legajos 
            WHERE situacion = 'ACTIVO' 
            ORDER BY apellido, nombre 

        """)
        bomberos = cur.fetchall()
        
        # 3. Traer config de puntos
        cur.execute("SELECT puntos_por_asistencia FROM config_puntos WHERE anio = YEAR(CURDATE()) LIMIT 1")
        config_puntos = cur.fetchone()
        
        conn.close()

    # Agregamos 'evento=None' para que la plantilla no de error de "Undefined"
    # También agregamos 'asistencias_previas' y 'postas_previas' como vacíos
    # para que el JavaScript de la plantilla funcione correctamente.
    return render_template("asistencia.html",
                           evento=None,
                           conceptos=conceptos,
                           departamentos=departamentos,
                           bomberos=bomberos,
                           config_puntos=config_puntos,
                           hoy=datetime.now().strftime('%Y-%m-%d'),
                           asistencias_previas={},
                           postas_previas=[])

@app.route("/asistencia/guardar", methods=["POST"])
@login_requerido
def guardar_asistencia():
    data = request.get_json()
    
    # --- 1. CAPTURA DE DATOS ---
    evento_id = data.get("evento_id")
    depto_id = data.get("departamento_id")
    if depto_id == "todos" or depto_id == "":
        depto_id = None

    # Capturamos los switches para saber quiénes califican
    # Asegúrate que en tu JS los envíes con estos nombres
    califica_oficiales = 1 if data.get("califica_oficiales") else 0
    califica_suboficiales = 1 if data.get("califica_suboficiales") else 0
    califica_encargados = 1 if data.get("califica_encargados") else 0

    tipo        = data.get("tipo")
    descripcion = data.get("descripcion", "")
    fecha       = data.get("fecha")
    hora_inicio = data.get("hora_inicio") or None
    hora_fin    = data.get("hora_fin") or None
    concepto_id = data.get("concepto_id") or None
    asistencias = data.get("asistencias", [])
    confirmar   = data.get("confirmar", False)
    temas       = data.get("temas", [])

    if not tipo or not fecha or not asistencias:
        return jsonify({"ok": False, "error": "Faltan datos obligatorios."})

    estado = "FINALIZADO" if confirmar else "BORRADOR"

    conn = get_db()
    if not conn:
        return jsonify({"ok": False, "error": "Error de conexión."})

    try:
        cur = conn.cursor()
        
        # --- 2. GUARDAR/ACTUALIZAR EVENTO ---
        if evento_id:
            cur.execute("""
                UPDATE eventos SET tipo=%s, descripcion=%s, fecha=%s,
                hora_inicio=%s, hora_fin=%s, concepto_id=%s, estado=%s, departamento_id=%s,
                califica_oficiales=%s, califica_suboficiales=%s, califica_encargados=%s
                WHERE id=%s
            """, (tipo, descripcion, fecha, hora_inicio, hora_fin,
                  concepto_id, estado, depto_id, 
                  califica_oficiales, califica_suboficiales, califica_encargados,
                  evento_id))
            
            # --- LIMPIEZA EN CASCADA MANUAL ---
            # 1. Borramos las notas de las postas de este evento específico
            cur.execute("""
                DELETE FROM asistencia_notas_temas 
                WHERE tema_id IN (SELECT id FROM evento_temas WHERE evento_id = %s)
            """, (evento_id,))

            # --- LIMPIEZA DE TEMAS PREVIOS (Para evitar duplicados al editar) ---
            cur.execute("DELETE FROM evento_temas WHERE evento_id = %s", (evento_id,))
        else:
            cur.execute("""
                INSERT INTO eventos (tipo, descripcion, fecha, hora_inicio, hora_fin,
                                     concepto_id, estado, creado_por, departamento_id,
                                     califica_oficiales, califica_suboficiales, califica_encargados)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (tipo, descripcion, fecha, hora_inicio, hora_fin,
                  concepto_id, estado, session["usuario_id"], depto_id,
                  califica_oficiales, califica_suboficiales, califica_encargados))
            evento_id = cur.lastrowid

        # --- 3. GUARDAR TEMAS/POSTAS ---
        if tipo == "CAPACITACION":
            for i, t in enumerate(temas):
                nombre_tema = t.get("nombre", "").strip()
                if nombre_tema:
                    cur.execute("""
                        INSERT INTO evento_temas (evento_id, nombre, calificador_legajo, orden)
                        VALUES (%s, %s, %s, %s)
                    """, (evento_id, nombre_tema, t.get("calificador_legajo") or None, i + 1))

        # --- 4. GUARDAR ASISTENCIAS ---
        for a in asistencias:
            nota_cruda = a.get("calificacion")
            nota_validada = None
            if nota_cruda is not None and nota_cruda != "":
                try:
                    val = float(nota_cruda)
                    nota_validada = max(0.0, min(5.0, val))
                except (ValueError, TypeError):
                    nota_validada = None

            cur.execute("""
                INSERT INTO asistencia (evento_id, legajo, estado, observacion, 
                                        calificacion, registrado_por)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE 
                    estado = VALUES(estado),
                    observacion = VALUES(observacion),
                    calificacion = %s,
                    registrado_por = VALUES(registrado_por)
            """, (evento_id, a["legajo"], a["estado"], 
                  a.get("observacion", ""), 
                  nota_validada,
                  session["usuario_id"],
                  nota_validada))
                  
        conn.commit()
        return jsonify({"ok": True, "evento_id": evento_id, "estado": estado})

    except Exception as e:
        conn.rollback()
        return jsonify({"ok": False, "error": str(e)})


@app.route("/asistencia/anular/<int:evento_id>", methods=["POST"])
@login_requerido
def anular_asistencia(evento_id):
    conn = get_db()
    if not conn:
        return jsonify({"ok": False, "error": "Error de conexión."})
    
    try:
        cur = conn.cursor()
        # Cambiamos el estado a ANULADO
        cur.execute("UPDATE eventos SET estado = 'ANULADO' WHERE id = %s", (evento_id,))
        
        # Opcional: Podrías querer borrar los registros de la tabla 'asistencia' 
        # o dejarlos pero que el sistema ignore los de eventos anulados.
        # cur.execute("DELETE FROM asistencia WHERE evento_id = %s", (evento_id,))
        
        conn.commit()
        cur.close()
        conn.close()
        return jsonify({"ok": True})
    except Exception as e:
        if conn: conn.rollback()
        return jsonify({"ok": False, "error": str(e)})

@app.route("/asistencia/borrador/<int:evento_id>")
@login_requerido
def editar_borrador(evento_id):
    conn = get_db()
    if not conn:
        flash("Error de conexión.", "danger")
        return redirect(url_for("inicio"))

    cur = conn.cursor(dictionary=True)
    
    # 1. Traer datos del evento
    cur.execute("SELECT * FROM eventos WHERE id = %s AND estado = 'BORRADOR'", (evento_id,))
    evento = cur.fetchone()

    if not evento:
        flash("El borrador no existe o ya fue confirmado.", "warning")
        return redirect(url_for("inicio"))

    # 2. TRAER LAS POSTAS/TEMAS GUARDADOS (Esto es lo que faltaba)
    cur.execute("""
        SELECT nombre, calificador_legajo 
        FROM evento_temas 
        WHERE evento_id = %s 
        ORDER BY orden ASC
    """, (evento_id,))
    postas_previas = cur.fetchall()

    # 3. Traer asistencias previas
    cur.execute("""
        SELECT a.legajo, a.estado, a.observacion, a.calificacion,
            l.apellido, l.nombre, l.grado, l.es_encargado  -- <-- CAMBIÁ 'autoriza' POR 'es_encargado'
        FROM asistencia a
        JOIN legajos l ON a.legajo = l.legajo
        WHERE a.evento_id = %s
    """, (evento_id,))
    asistencias = cur.fetchall()

    # Y asegurate de que el diccionario use el nuevo valor numérico (0 o 1)
    dict_asistencias = {str(a['legajo']): {
        'estado': a['estado'], 
        'observacion': a['observacion'],
        'calificacion': a['calificacion'],
        'apellido': a['apellido'],
        'nombre': a['nombre'],
        'grado': a['grado'],
        # Como ahora es TINYINT(1), esto lo convierte a True o False para JS
        'es_encargado': bool(a['es_encargado']) 
    } for a in asistencias}

    # Datos complementarios para los selectores
    cur.execute("SELECT id, concepto FROM conceptos WHERE activo = 1 ORDER BY concepto")
    conceptos = cur.fetchall()
    cur.execute("SELECT id, nombre FROM departamentos WHERE activo = 1 ORDER BY nombre")
    departamentos = cur.fetchall()
    cur.execute("SELECT puntos_por_asistencia FROM config_puntos WHERE anio = YEAR(CURDATE()) LIMIT 1")
    config_puntos = cur.fetchone()

    conn.close()

    return render_template("asistencia.html",
                           evento=evento,
                           postas_previas=postas_previas, # <--- Pasamos las postas al HTML
                           asistencias_previas=dict_asistencias,
                           conceptos=conceptos,
                           departamentos=departamentos,
                           config_puntos=config_puntos)

@app.route("/asistencia/historial")
@login_requerido
def historial_asistencia():
    ver_anulados = request.args.get("ver_anulados") == "1"
    legajo_usuario = session.get('legajo')
    
    conn = get_db()
    eventos = []
    
    # 1. DEFINIR LA QUERY BASE FUERA DE LOS IF DE FILTRADO
    query = """
        SELECT 
            e.*, 
            d.nombre as nombre_departamento,
            (SELECT COUNT(*) FROM asistencia WHERE evento_id = e.id AND estado = 'PRESENTE') as presentes,
            (SELECT COUNT(*) FROM asistencia WHERE evento_id = e.id AND estado = 'AUSENTE') as ausentes,
            (SELECT COUNT(*) FROM asistencia WHERE evento_id = e.id AND estado = 'JUSTIFICADO') as justificados,
            a_personal.estado as mi_estado
        FROM eventos e
        LEFT JOIN departamentos d ON e.departamento_id = d.id
        LEFT JOIN asistencia a_personal ON e.id = a_personal.evento_id AND a_personal.legajo = %s
        WHERE 1=1
    """

    if conn:
        cur = conn.cursor(dictionary=True)
        
        # 2. AHORA SÍ PODÉS CONCATENAR SIN RIESGO
        if not ver_anulados:
            query += " AND e.estado != 'ANULADO'"
            
        query += " ORDER BY e.fecha DESC, e.id DESC LIMIT 50"
        
        cur.execute(query, (legajo_usuario,))
        eventos = cur.fetchall()
        conn.close()

    return render_template("historial_asistencia.html", eventos=eventos, ver_anulados=ver_anulados)

@app.route("/asistencia/detalle/<int:evento_id>")
@login_requerido
@rol_requerido('ADMIN', 'JEFATURA')
def detalle_asistencia(evento_id):
    conn = get_db()
    evento    = None
    registros = []
    if conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT e.*, c.concepto, IFNULL(d.nombre, 'General / Grupal') as nombre_departamento
            FROM eventos e
            LEFT JOIN conceptos c ON e.concepto_id = c.id
            LEFT JOIN departamentos d ON e.departamento_id = d.id
            WHERE e.id = %s
        """, (evento_id,))
        evento = cur.fetchone()

        cur.execute("""
            SELECT a.estado, a.observacion, a.calificacion,
                   l.legajo, l.apellido, l.nombre, l.grado
            FROM asistencia a
            JOIN legajos l ON a.legajo = l.legajo
            WHERE a.evento_id = %s
            ORDER BY l.apellido, l.nombre
        """, (evento_id,))
        registros = cur.fetchall()
        conn.close()

    return render_template("detalle_asistencia.html", evento=evento, registros=registros)

# RUTA PARA GESTIONAR (Pantalla HTML normal)
@app.route('/registro-salidas/detalle/<int:id>')
@login_requerido
def detalle_siniestro(id):
    db = get_db()
    cur = db.cursor(dictionary=True)
    cur.execute("SELECT * FROM nexo_siniestros WHERE id = %s", (id,))
    siniestro = cur.fetchone()
    
    cur.execute("""
        SELECT p.*, l.nombre, l.apellido, l.grado as jerarquia 
        FROM nexo_personal p
        JOIN legajos l ON p.legajo = l.legajo
        WHERE p.siniestro_id = %s
    """, (id,))
    personal = cur.fetchall()
    cur.close()
    
    # IMPORTANTE: Esta renderiza la pantalla de GESTIÓN
    return render_template('nexo_detalle.html', siniestro=siniestro, personal=personal)

# RUTA PARA IMPRIMIR (El PDF Virtual con renglones)
@app.route('/registro-salidas/imprimir/<int:id>')
@login_requerido
def imprimir_siniestro(id):
    db = get_db()
    cur = db.cursor(dictionary=True)
    cur.execute("SELECT * FROM nexo_siniestros WHERE id = %s", (id,))
    siniestro = cur.fetchone()
    
    cur.execute("""
        SELECT p.*, l.nombre, l.apellido, l.jerarquia 
        FROM nexo_personal p
        JOIN legajos l ON p.legajo = l.legajo
        WHERE p.siniestro_id = %s
    """, (id,))
    personal = cur.fetchall()
    cur.close()
    
    # IMPORTANTE: Esta renderiza tu planilla de firmas
    return render_template('nexo_reporte.html', siniestro=siniestro, personal=personal)

# ============================================================
# CAPACITACIONES - POSTAS Y CALIFICACIONES
# ============================================================

@app.route("/evento/<int:evento_id>/temas/guardar", methods=["POST"])
@login_requerido
def guardar_temas_evento(evento_id):
    """Guarda los temas/postas que tendrá una capacitación específica"""
    nombres_temas = request.form.getlist("nombre_tema")
    calificadores  = request.form.getlist("calificador_legajo")
    
    conn = get_db()
    if conn:
        try:
            cur = conn.cursor()
            # 1. Limpiar temas existentes para este evento por si es una edición
            cur.execute("DELETE FROM evento_temas WHERE evento_id = %s", (evento_id,))
            
            # 2. Insertar los nuevos temas
            for i, nombre in enumerate(nombres_temas):
                if nombre.strip():  # Solo si escribieron algo en el nombre
                    # Si no eligieron calificador, ponemos None (NULL en la DB)
                    calificador = calificadores[i] if calificadores[i] else None
                    
                    cur.execute("""
                        INSERT INTO evento_temas (evento_id, nombre, calificador_legajo, orden)
                        VALUES (%s, %s, %s, %s)
                    """, (evento_id, nombre, calificador, i + 1))
            
            conn.commit()
            flash("Estructura de la capacitación configurada.", "success")
        except Error as e:
            flash(f"Error al guardar postas: {e}", "danger")
        finally:
            conn.close()
    
    # Redirige de vuelta al detalle para empezar a calificar o ver el resumen
    return redirect(url_for('detalle_asistencia', evento_id=evento_id))

@app.route("/asistencia/notas/guardar/<int:evento_id>", methods=["POST"])
@login_requerido
def guardar_calificaciones_postas(evento_id):
    conn = get_db()
    if not conn:
        flash("Error de conexión.", "danger")
        return redirect(url_for('detalle_asistencia', evento_id=evento_id))

    # --- NUEVO: Capturamos la acción del botón ---
    accion = request.form.get('accion') 

    try:
        cur = conn.cursor()
        notas_vacias = 0
        notas_guardadas = 0
        
        for key, value in request.form.items():
            if key.startswith("nota_"):
                if value.strip() == "":
                    notas_vacias += 1
                    continue
                
                parts = key.split("_")
                legajo = parts[1]
                tema_id = parts[2]
                nota = float(value)

                cur.execute("""
                    INSERT INTO asistencia_notas_temas (evento_id, tema_id, legajo, nota)
                    VALUES (%s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE nota = VALUES(nota)
                """, (evento_id, tema_id, legajo, nota))
                notas_guardadas += 1

        # --- NUEVO: Si la acción es finalizar, cambiamos el estado del evento ---
        if accion == 'finalizar':
            # Asumiendo que tu tabla eventos tiene una columna 'estado'
            cur.execute("UPDATE eventos SET estado = 'FINALIZADO' WHERE id = %s", (evento_id,))
            conn.commit()
            flash("Planilla finalizada y cerrada. Ya no se puede editar.", "success")
            return redirect(url_for('historial_asistencia'))

        # Si es solo guardar borrador
        conn.commit()
        
        if notas_guardadas > 0 and notas_vacias > 0:
            flash(f"Borrador guardado: {notas_guardadas} notas cargadas, faltan {notas_vacias}.", "warning")
        else:
            flash("Borrador actualizado correctamente.", "info")
            
    except Exception as e:
        if conn: conn.rollback()
        flash(f"Error al guardar: {e}", "danger")
    finally:
        if conn: conn.close()

    return redirect(url_for('cargar_notas', evento_id=evento_id))

# ============================================================
# DEPARTAMENTOS
# ============================================================

@app.route("/departamentos")
@login_requerido
def departamentos():
    conn = get_db()
    lista = []
    if conn:
        try:
            cur = conn.cursor(dictionary=True)
            # Tu consulta está perfecta: filtra deptos activos y cuenta miembros activos
            cur.execute("""
                SELECT d.*, COUNT(bd.id) as miembros
                FROM departamentos d
                LEFT JOIN bombero_departamento bd ON d.id = bd.departamento_id AND bd.activo = 1
                WHERE d.activo = 1
                GROUP BY d.id
                ORDER BY d.nombre
            """)
            lista = cur.fetchall()
        except Exception as e:
            print(f"Error al obtener departamentos: {e}")
        finally:
            conn.close()
    return render_template("departamentos.html", departamentos=lista)

@app.route("/departamentos/guardar", methods=["POST"])
@login_requerido  # Asegúrate de que use el nombre correcto
def guardar_departamento():
    if session.get('rol') != 'ADMIN':
        return redirect(url_for('inicio'))
        
    depto_id = request.form.get('id')
    nombre = request.form.get('nombre')
    descripcion = request.form.get('descripcion')
    
    # --- CORRECCIÓN AQUÍ ---
    conn = get_db()
    if conn:
        try:
            cur = conn.cursor()
            if depto_id: # Editar existente
                cur.execute("UPDATE departamentos SET nombre = %s, descripcion = %s WHERE id = %s", 
                           (nombre, descripcion, depto_id))
            else: # Nuevo
                cur.execute("INSERT INTO departamentos (nombre, descripcion, activo) VALUES (%s, %s, 1)", 
                           (nombre, descripcion))
            
            conn.commit()
        except Exception as e:
            print(f"Error al guardar: {e}")
        finally:
            conn.close()
    # -----------------------
    
    return redirect(url_for('departamentos'))

@app.route("/departamentos/eliminar/<int:id>", methods=["POST"])
@login_requerido
def eliminar_departamento(id):
    # Solo permitimos que el ADMIN realice esta acción
    if session.get('rol') != 'ADMIN':
        return redirect(url_for('inicio'))
        
    conn = get_db()
    if conn:
        try:
            cur = conn.cursor()
            # Cambiamos el estado a 0 (Baja Lógica)
            cur.execute("UPDATE departamentos SET activo = 0 WHERE id = %s", (id,))
            conn.commit()
        finally:
            conn.close()
            
    return redirect(url_for('departamentos'))

@app.route("/departamentos/gestionar/<int:depto_id>")
@login_requerido
def miembros_departamento(depto_id):
    # RESTRICCIÓN CRÍTICA:
    if session.get('rol') != 'ADMIN':
        flash("Acceso denegado: Solo el Administrador puede asignar personal.", "danger")
        return redirect(url_for('departamentos'))
    conn = get_db()
    depto = None
    miembros = []
    todos = []
    if conn:
        cur = conn.cursor(dictionary=True)
        # 1. Datos del depto
        cur.execute("SELECT * FROM departamentos WHERE id = %s", (depto_id,))
        depto = cur.fetchone()

        # 2. Miembros actuales del depto
        cur.execute("""
            SELECT l.legajo, l.apellido, l.nombre, l.grado,
                   bd.fecha_ingreso, bd.id as bd_id
            FROM bombero_departamento bd
            JOIN legajos l ON bd.legajo = l.legajo
            WHERE bd.departamento_id = %s AND bd.activo = 1
            ORDER BY l.apellido
        """, (depto_id,))
        miembros = cur.fetchall()

        # 3. LISTA PARA EL SELECTOR: Traemos a todos y sus deptos actuales
        cur.execute("""
            SELECT l.legajo, l.apellido, l.nombre, l.grado,
                GROUP_CONCAT(d.nombre SEPARATOR ', ') as deptos_nombres
            FROM legajos l
            LEFT JOIN bombero_departamento bd ON l.legajo = bd.legajo AND bd.activo = 1
            LEFT JOIN departamentos d ON bd.departamento_id = d.id
            WHERE l.situacion = 'ACTIVO'
            GROUP BY l.legajo, l.apellido, l.nombre, l.grado
            ORDER BY l.apellido, l.nombre
        """)
        todos = cur.fetchall()
        conn.close()

    return render_template("miembros_departamento.html", 
                           depto=depto, miembros=miembros, todos=todos)

@app.route("/departamentos/<int:depto_id>/agregar", methods=["POST"])
@login_requerido
@rol_requerido('ADMIN', 'JEFATURA') # <--- Bloqueo para bomberos
def agregar_miembro(depto_id):
    # 'getlist' permite capturar todos los bomberos seleccionados en el select múltiple
    legajos = request.form.getlist("legajo")
    fecha_ingreso = request.form.get("fecha_ingreso") or datetime.now().strftime("%Y-%m-%d")
    
    conn = get_db()
    if conn and legajos:
        try:
            cur = conn.cursor()
            for legajo in legajos:
                # Esta consulta es inteligente: solo inserta si el bombero NO está ya activo en ESTE depto
                cur.execute("""
                    INSERT INTO bombero_departamento (legajo, departamento_id, fecha_ingreso, activo)
                    SELECT %s, %s, %s, 1
                    WHERE NOT EXISTS (
                        SELECT 1 FROM bombero_departamento 
                        WHERE legajo = %s AND departamento_id = %s AND activo = 1
                    )
                """, (legajo, depto_id, fecha_ingreso, legajo, depto_id))
            
            conn.commit()
            flash(f"Proceso finalizado. Se intentaron agregar {len(legajos)} bomberos.", "success")
        except Error as e:
            flash(f"Error en la base de datos: {e}", "danger")
        finally:
            conn.close()
    return redirect(url_for("miembros_departamento", depto_id=depto_id))


@app.route("/departamentos/miembro/<int:bd_id>/quitar", methods=["POST"])
@login_requerido
def quitar_miembro(bd_id):
    depto_id = request.form.get("depto_id")
    fecha_egreso = datetime.now().strftime("%Y-%m-%d")
    
    conn = get_db()
    if conn:
        try:
            cur = conn.cursor()
            # Baja lógica: desactivamos el registro y marcamos fecha de egreso
            cur.execute("""
                UPDATE bombero_departamento
                SET activo = 0, fecha_egreso = %s
                WHERE id = %s
            """, (fecha_egreso, bd_id))
            conn.commit()
            flash("Bombero removido del departamento.", "success")
        except Error as e:
            flash(f"Error al quitar miembro: {e}", "danger")
        finally:
            conn.close()
            
    return redirect(url_for("miembros_departamento", depto_id=depto_id))

# ============================================================
# CURSOS
# ============================================================

@app.route("/cursos")
@login_requerido
def cursos():
    conn = get_db()
    lista = []
    legajo_usuario = session.get('legajo') # Obtenemos el legajo del usuario actual
    
    if conn:
        cur = conn.cursor(dictionary=True)
        # Agregamos la subconsulta 'soy_participante' para marcar si el usuario estuvo ahí
        cur.execute("""
            SELECT 
                c.*, 
                COUNT(cp.id) as participantes,
                (SELECT COUNT(*) FROM curso_participantes 
                 WHERE curso_id = c.id AND legajo = %s) as soy_participante
            FROM cursos c
            LEFT JOIN curso_participantes cp ON c.id = cp.curso_id
            GROUP BY c.id
            ORDER BY c.fecha_inicio DESC
        """, (legajo_usuario,))
        lista = cur.fetchall()
        conn.close()
    return render_template("cursos.html", cursos=lista)


@app.route("/cursos/nuevo", methods=["GET", "POST"])
@login_requerido
@rol_requerido('ADMIN', 'JEFATURA')
def nuevo_curso():
    if request.method == "POST":
        nombre      = request.form.get("nombre")
        institucion = request.form.get("institucion", "")
        fecha_ini   = request.form.get("fecha_inicio") or None
        fecha_fin   = request.form.get("fecha_fin") or None
        horas       = request.form.get("horas") or 0
        descripcion = request.form.get("descripcion", "")
        legajos     = request.form.getlist("participantes")

        if not nombre or not legajos:
            flash("Faltan datos obligatorios (Nombre del curso o participantes).", "warning")
            return redirect(url_for("nuevo_curso"))

        conn = get_db()
        if conn:
            try:
                cur = conn.cursor()
                # 1. Insertar el curso
                cur.execute("""
                    INSERT INTO cursos (nombre, institucion, fecha_inicio, fecha_fin,
                                        horas, descripcion, creado_por)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (nombre, institucion, fecha_ini, fecha_fin,
                      horas, descripcion, session.get("usuario_id")))
                
                curso_id = cur.lastrowid

                # 2. Insertar participantes (usando executemany para ser más rápido)
                valores_participantes = [(curso_id, legajo) for legajo in legajos]
                cur.executemany("""
                    INSERT INTO curso_participantes (curso_id, legajo)
                    VALUES (%s, %s)
                """, valores_participantes)

                conn.commit()
                flash(f"Curso '{nombre}' registrado con {len(legajos)} participantes.", "success")
                return redirect(url_for("cursos"))
            except Exception as e:
                conn.rollback()
                flash(f"Error crítico al guardar: {e}", "danger")
            finally:
                conn.close()

    # GET: Carga de datos para el formulario
    conn = get_db()
    bomberos = []
    departamentos = []
    if conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT legajo, apellido, nombre, grado
            FROM legajos WHERE situacion = 'ACTIVO'
            ORDER BY apellido, nombre
        """)
        bomberos = cur.fetchall()
        cur.execute("SELECT id, nombre FROM departamentos WHERE activo = 1 ORDER BY nombre")
        departamentos = cur.fetchall()
        conn.close()

    return render_template("nuevo_curso.html", bomberos=bomberos, departamentos=departamentos)

# ============================================================
# BOMBEROS
# ============================================================

@app.route("/bomberos")
@login_requerido
def bomberos():
    conn = get_db()
    lista = []
    if conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT legajo, apellido, nombre, grado, cargo, situacion, nro_cel, email
            FROM legajos
            WHERE situacion != 'BAJA'
            ORDER BY situacion, apellido, nombre
    
        """)
        lista = cur.fetchall()
        conn.close()
    return render_template("bomberos.html", bomberos=lista)

@app.route("/asistencia/bomberos")
@login_requerido
def asistencia_bomberos_json():
    # Obtenemos el departamento si viene en la URL, sino 'todos'
    dep_id = request.args.get('departamento_id', 'todos')
    
    conn = get_db()
    lista = []
    if conn:
        cur = conn.cursor(dictionary=True)
        
        # Consulta base
        query = "SELECT legajo, apellido, nombre, grado, rango_categoria, es_encargado FROM legajos WHERE situacion != 'BAJA'"
        params = []

        # Si se eligió un departamento específico, filtramos
        if dep_id != 'todos':
            query += " AND departamento_id = %s"
            params.append(dep_id)

        query += " ORDER BY apellido, nombre"
        
        cur.execute(query, params)
        lista = cur.fetchall()
        conn.close()
    
    # IMPORTANTE: Devolvemos JSON, no render_template
    return jsonify(lista)

# ============================================================
# CONFIGURACIÓN DE PUNTOS
# ============================================================

from datetime import datetime

@app.route("/config/puntos")
@login_requerido
@rol_requerido("ADMIN")
def config_puntos():
    conn = get_db()
    registros = []
    # Obtenemos el año actual para el formulario
    now_year = datetime.now().year 
    
    if conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM config_puntos ORDER BY anio DESC")
        registros = cur.fetchall()
        conn.close()
    
    # Pasamos 'now_year' al template
    return render_template("config_puntos.html", registros=registros, now_year=now_year)


@app.route("/config/puntos/guardar", methods=["POST"])
@login_requerido
@rol_requerido("ADMIN")
def guardar_config_puntos():
    anio        = request.form.get("anio")
    puntos      = request.form.get("puntos_por_asistencia")
    descripcion = request.form.get("descripcion", "")
    conn = get_db()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO config_puntos (anio, puntos_por_asistencia, descripcion, creado_por)
                VALUES (%s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                puntos_por_asistencia = VALUES(puntos_por_asistencia),
                descripcion = VALUES(descripcion)
            """, (anio, puntos, descripcion, session["usuario_id"]))
            conn.commit()
            flash("Configuración de puntos guardada.", "success")
        except Error as e:
            flash(f"Error: {e}", "danger")
        finally:
            conn.close()
    return redirect(url_for("config_puntos"))

# ============================================================
# EXPORTACIÓN Y CALIFICACIONES (NUEVO)
# ============================================================

@app.route("/asistencia/notas/<int:evento_id>")
@login_requerido
def cargar_notas(evento_id):
    conn = get_db()
    if not conn: return redirect(url_for("inicio"))
    
    cur = conn.cursor(dictionary=True)
    
    # 1. Datos del evento
    cur.execute("SELECT * FROM eventos WHERE id = %s", (evento_id,))
    evento = cur.fetchone()

    # 2. Traer solo bomberos que figuran como 'PRESENTE'
    cur.execute("""
        SELECT a.legajo, l.apellido, l.nombre, l.grado
        FROM asistencia a
        JOIN legajos l ON a.legajo = l.legajo
        WHERE a.evento_id = %s AND a.estado = 'PRESENTE'
        ORDER BY l.apellido
    """, (evento_id,))
    presentes = cur.fetchall()

    # 3. Traer los temas/postas de este evento
    cur.execute("SELECT * FROM evento_temas WHERE evento_id = %s ORDER BY orden", (evento_id,))
    temas = cur.fetchall()

    # 4. Traer notas ya existentes (para modo edición/borrador)
    cur.execute("SELECT * FROM asistencia_notas_temas WHERE evento_id = %s", (evento_id,))
    notas_db = cur.fetchall()
    
    # Mapeamos las notas en un dict {(legajo, tema_id): nota} para fácil acceso en el template
    notas_map = {(n['legajo'], n['tema_id']): n['nota'] for n in notas_db}

    conn.close()
    return render_template("cargar_notas.html", 
                            evento=evento, 
                            presentes=presentes, 
                            temas=temas, 
                            mapa_notas=notas_map) # <-- Cambié el nombre a mapa_notas

@app.route("/asistencia/exportar/<int:evento_id>/<formato>")
@login_requerido
def exportar_asistencia(evento_id, formato):
    import pandas as pd
    from io import BytesIO
    from flask import send_file

    conn = get_db()
    cur = conn.cursor(dictionary=True)
    
    # Buscamos los presentes
    cur.execute("""
        SELECT l.legajo, l.apellido, l.nombre, l.grado, a.estado, a.observacion
        FROM asistencia a
        JOIN legajos l ON a.legajo = l.legajo
        WHERE a.evento_id = %s
        ORDER BY l.apellido, l.nombre
    """, (evento_id,))
    asistencias = cur.fetchall()
    conn.close()

    if formato == 'excel':
        df = pd.DataFrame(asistencias)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Asistencia')
        output.seek(0)
        
        return send_file(output, 
                         download_name=f"asistencia_{evento_id}.xlsx", 
                         as_attachment=True)
    
    return f"Formato {formato} en desarrollo", 404

# ============================================================
# REPORTES Y ESTADÍSTICAS (AÑADIR AL FINAL DE APP.PY)
# ============================================================

@app.route("/reportes/actividad")
@login_requerido
@rol_requerido("ADMIN", "JEFATURA") # Solo Jefatura puede ver el rendimiento general
def reporte_actividad():
    # 1. Filtros de fecha (por defecto el mes actual)
    fecha_desde = request.args.get("desde", datetime.now().strftime("%Y-%m-01"))
    fecha_hasta = request.args.get("hasta", datetime.now().strftime("%Y-%m-%d"))
    
    conn = get_db()
    data_reporte = []
    
    if conn:
        cur = conn.cursor(dictionary=True)
        
        # 2. La consulta SQL: Enfocada en Actividad Operativa y Capacitación
        query = """
            SELECT 
                l.legajo, l.apellido, l.nombre, l.grado,
                COUNT(DISTINCT CASE WHEN a.estado = 'PRESENTE' THEN e.id END) as total_asistencias,
                AVG(ant.nota) as promedio_capacitacion
            FROM legajos l
            LEFT JOIN asistencia a ON l.legajo = a.legajo
            LEFT JOIN eventos e ON a.evento_id = e.id
            LEFT JOIN asistencia_notas_temas ant ON l.legajo = ant.legajo AND e.id = ant.evento_id
            WHERE (e.estado IN ('CONFIRMADO', 'FINALIZADO') OR e.id IS NULL)
              AND l.situacion = 'ACTIVO'
              AND (e.fecha BETWEEN %s AND %s OR e.fecha IS NULL)
            GROUP BY l.legajo, l.apellido, l.nombre, l.grado
            ORDER BY total_asistencias DESC, l.apellido ASC
        """
        cur.execute(query, (fecha_desde, fecha_hasta))
        resultados = cur.fetchall()
        
        # 3. Procesamos los datos
        for res in resultados:
            asistencias = res['total_asistencias'] or 0
            promedio = res['promedio_capacitacion'] or 0.0
            
            # Aquí podrías aplicar una fórmula de "Puntaje de Mérito" si la tienen
            # Por ahora, simplemente listamos la actividad real.
            data_reporte.append({
                'legajo': res['legajo'],
                'nombre': f"{res['apellido']}, {res['nombre']}",
                'grado': res['grado'],
                'asistencias': asistencias,
                'promedio': round(promedio, 2)
            })
        conn.close()

    return render_template("reporte_actividad.html", 
                           reporte=data_reporte, 
                           desde=fecha_desde, 
                           hasta=fecha_hasta)

def obtener_datos_completos_perfil(legajo):
    conn = get_db()
    datos = {
        'legajo': legajo, 'apellido': '---', 'nombre': '---',
        'pilar_vocacion': 0.0, 'pilar_tecnica': 0.0,
        'pilar_cualidades': 0.0, 'pilar_asistencia': 0.0,
        'puntaje_final': 0.0, 'pendientes_firma_bombero': 0,
        'calif_letra': '---', 'calif_desc': 'SIN CALIFICAR',
        'grado': '---', 'cargo': 'Sin asignar', 'situacion': '---', 'email': 'N/A',
        'clases_conteo': 0, 'total_salidas': 0,
        'horas_actividad_reales': 0.0, 'puntos_actividad': 0.0,
        'promedio_general': 0.0
    }

    if not conn: 
        print("❌ ERROR: No hay conexión a la base de datos")
        return datos, []

    try:
        cur = conn.cursor(dictionary=True)
        from datetime import datetime
        mes_actual = datetime.now().strftime('%Y-%m')
        
        print(f"\n--- DEBUG PERFIL LEGAJO: {legajo} (Mes: {mes_actual}) ---")

        # --- 1. DATOS PERSONALES ---
        cur.execute("SELECT apellido, nombre, grado, cargo, situacion FROM legajos WHERE legajo = %s", (legajo,))
        perfil = cur.fetchone()
        if perfil:
            datos.update(perfil)
            print(f"✅ Perfil encontrado: {perfil['apellido']}, {perfil['nombre']}")

        # --- 2. REGISTROS DE ASISTENCIA (EL CORAZÓN DEL PROBLEMA) ---
        query_asistencia = """
            SELECT a.estado, a.calificacion 
            FROM asistencia a 
            JOIN eventos e ON a.evento_id = e.id 
            WHERE a.legajo = %s AND e.fecha LIKE %s AND e.estado = 'FINALIZADO'
        """
        cur.execute(query_asistencia, (legajo, f"{mes_actual}%"))
        registros = cur.fetchall()
        
        print(f"📊 Registros encontrados: {len(registros)}")
        for i, r in enumerate(registros):
            print(f"   -> Evento {i+1}: Estado={r['estado']}, Nota={r['calificacion']}")

        total_clases = len(registros)
        asistidas = len([r for r in registros if r['estado'] == 'PRESENTE'])
        
        # --- 3. CÁLCULO DE PUNTOS CLASES ---
        pts_clases = (asistidas / total_clases * 5) if total_clases > 0 else 0.0
        print(f"⭐ Puntos Clases: {pts_clases} (Asistió {asistidas} de {total_clases})")

        # --- 4. PROMEDIO DE NOTAS ---
        notas = [float(r['calificacion']) for r in registros 
                 if r.get('calificacion') is not None and str(r['calificacion']).strip() != '']
        
        datos['promedio_general'] = round(sum(notas) / len(notas), 2) if notas else 0.0
        print(f"📝 Notas extraídas: {notas} -> Promedio: {datos['promedio_general']}")

        # --- 5. SINIESTROS ---
        cur.execute("""
            SELECT COUNT(p.id) as cant FROM nexo_personal p 
            JOIN nexo_siniestros s ON p.siniestro_id = s.id
            WHERE p.legajo = %s AND s.fecha LIKE %s
        """, (legajo, f"{mes_actual}%"))
        res_mias = cur.fetchone()
        datos['total_salidas'] = res_mias['cant'] if res_mias else 0
        print(f"🚒 Salidas del mes: {datos['total_salidas']}")

        # Forzamos pilar asistencia para el diamante
        pts_siniestros = 0.0 # Simplificado para el debug
        datos['pilar_asistencia'] = round((pts_clases + pts_siniestros) / 2, 2)
        print(f"💎 Pilar Asistencia (Diamante): {datos['pilar_asistencia']}")

        # --- 6. CONSOLIDACIÓN ---
        datos['puntaje_final'] = round(
            datos['pilar_vocacion'] + datos['pilar_tecnica'] + 
            datos['pilar_cualidades'] + datos['pilar_asistencia'], 2
        )
        print(f"🏆 Puntaje Final: {datos['puntaje_final']}")
        print("-------------------------------------------\n")

        return datos, registros

    except Exception as e:
        import traceback
        print(f"❌ ERROR CRÍTICO: {str(e)}")
        traceback.print_exc() # Esto te dirá la línea exacta del error
        return datos, []
    
@app.route("/ver-perfil/<int:legajo_id>")
@rol_requerido('ADMIN', 'JEFATURA')
def ver_perfil_ajeno(legajo_id):
    datos, historial = obtener_datos_completos_perfil(legajo_id)
    
    if not datos:
        flash("No se encontró el legajo.", "warning")
        return redirect(url_for('inicio'))
    
    # FILTRO: Limpiamos el historial de cualquier registro anulado
    # Esto evita que los errores de carga o legajos 50000 sumen aquí
    historial_filtrado = [h for h in historial if h.get('estado') != 'ANULADO']
    
    return render_template("mi_perfil.html", datos=datos, historial=historial_filtrado)

@app.route("/mi_perfil")
@login_requerido
def mi_perfil():
    legajo = session.get('legajo')
    conn = get_db()
    cur = conn.cursor(dictionary=True)

    # --- 1. DATOS BÁSICOS DEL BOMBERO ---
    cur.execute("SELECT * FROM legajos WHERE legajo = %s", (legajo,))
    usuario = cur.fetchone()

    if not usuario:
        cur.close()
        return f"Error: El legajo '{legajo}' no existe.", 404

    # --- 0. CONFIGURACIÓN DE FECHAS ---
    from datetime import datetime
    ahora = datetime.now()
    filtro_mes = ahora.strftime('%Y-%m') + "%"
    filtro_anio_actual = ahora.strftime('%Y') + "%"
    filtro_anio_anterior = str(int(ahora.strftime('%Y')) - 1) + "%"

    servicios_especiales = (
        'Servicios Especiales', 'Capacitación', 'Prevención', 'Falsa Alarma', 
        'Representación', 'Falso Aviso', 'Suministro de agua', 'Otros',
        'Extracción de panales', 'Retirado de ovito', 'Colaboración con fuerzas de seguridad',
        'Colocación de driza', 'Servicio: Suministro de Agua', 'Servicio: Otros',
        'Prevención: Eventos'
    )
    placeholders = ', '.join(['%s'] * len(servicios_especiales))

    # --- 1. RECOLECCIÓN DE DATOS (MES ACTUAL) ---

    # BOTÓN 1: Horas SIAB (Sin filtro de estado para asegurar visualización de cargas locales)
    cur.execute("""
        SELECT SUM(horas) as total FROM actividades 
        WHERE legajo = %s AND fecha_inicio LIKE %s 
        AND actividad NOT LIKE '%%CLASE OBLIGATORIA%%' 
        AND (anulada IS NULL OR anulada = 0)
    """, (legajo, filtro_mes))
    res_b1 = cur.fetchone()
    b1_horas_siab = float(res_b1['total'] or 0.0)

    # BOTÓN 2: Servicios RUBA
    cur.execute(f"""
        SELECT COUNT(p.id) as cant FROM nexo_personal p 
        JOIN nexo_siniestros s ON p.siniestro_id = s.id 
        WHERE p.legajo = %s AND s.fecha LIKE %s AND s.tipo_siniestro IN ({placeholders})
    """, (legajo, filtro_mes, *servicios_especiales))
    res_b2 = cur.fetchone()
    b2_servicios_ruba = int(res_b2['cant'] or 0)

    # BOTÓN 3: Notas Emergencia (Promedio últimas 5)
    cur.execute(f"""
        SELECT p.puntos_operativos FROM nexo_personal p
        JOIN nexo_siniestros s ON p.siniestro_id = s.id
        WHERE p.legajo = %s AND s.tipo_siniestro NOT IN ({placeholders})
        ORDER BY s.fecha DESC 
    """, (legajo, *servicios_especiales))
    ultimas_notas = cur.fetchall()
    b3_promedio_emergencia = round(sum(float(n['puntos_operativos'] or 0) for n in ultimas_notas) / len(ultimas_notas), 2) if ultimas_notas else 0.0

    # BOTÓN 6: Clases Obligatorias
    cur.execute("""
        SELECT COUNT(*) as cant FROM actividades 
        WHERE legajo = %s AND fecha_inicio LIKE %s 
        AND actividad LIKE '%%CLASE OBLIGATORIA%%' 
        AND (anulada IS NULL OR anulada = 0)
    """, (legajo, filtro_mes))
    res_b6 = cur.fetchone()
    b6_clases_oblig = int(res_b6['cant'] or 0)

    # BOTÓN 7: Siniestros Reales
    cur.execute(f"""
        SELECT COUNT(p.id) as cant FROM nexo_personal p 
        JOIN nexo_siniestros s ON p.siniestro_id = s.id 
        WHERE p.legajo = %s AND s.fecha LIKE %s 
        AND s.tipo_siniestro NOT IN ({placeholders})
        AND s.tipo_siniestro NOT LIKE 'Servicio%%'
    """, (legajo, filtro_mes, *servicios_especiales))
    res_b7 = cur.fetchone()
    b7_siniestros_reales = int(res_b7['cant'] or 0)

    # --- 2. TOTALES ANUALES (RECONOCIMIENTO) ---

    # Horas Totales Año Actual
    cur.execute("SELECT SUM(horas) as total FROM actividades WHERE legajo = %s AND fecha_inicio LIKE %s AND (anulada IS NULL OR anulada = 0)", (legajo, filtro_anio_actual))
    total_anio_actual = float(cur.fetchone()['total'] or 0.0)

    # Horas Totales Año Anterior
    cur.execute("SELECT SUM(horas) as total FROM actividades WHERE legajo = %s AND fecha_inicio LIKE %s AND (anulada IS NULL OR anulada = 0)", (legajo, filtro_anio_anterior))
    total_anio_anterior = float(cur.fetchone()['total'] or 0.0)

    # Máximo del Cuartel (Para proporcionalidad de asistencia)
    cur.execute(f"""
        SELECT COUNT(p.id) as cant FROM nexo_personal p 
        JOIN nexo_siniestros s ON p.siniestro_id = s.id 
        WHERE s.fecha LIKE %s AND s.tipo_siniestro NOT IN ({placeholders})
        GROUP BY p.legajo ORDER BY cant DESC LIMIT 1
    """, (filtro_mes, *servicios_especiales))
    res_max = cur.fetchone()
    max_cuartel = res_max['cant'] if res_max and res_max['cant'] > 0 else 1

    # --- NUEVO: MÁXIMO DE SERVICIOS ESPECIALES DEL CUARTEL ---
    cur.execute(f"""
        SELECT COUNT(p.id) as cant FROM nexo_personal p 
        JOIN nexo_siniestros s ON p.siniestro_id = s.id 
        WHERE s.fecha LIKE %s AND s.tipo_siniestro IN ({placeholders})
        GROUP BY p.legajo ORDER BY cant DESC LIMIT 1
    """, (filtro_mes, *servicios_especiales))
    res_max_servicios = cur.fetchone()
    
    # Si nadie hizo servicios en el mes, el divisor es 1 para evitar errores
    max_servicios_cuartel = res_max_servicios['cant'] if res_max_servicios and res_max_servicios['cant'] > 0 else 1

    # --- 3. CÁLCULO DE PILARES (ESCALA 0-5) ---

    # PILAR 1: VOCACIÓN
    # Definimos horas_puntuables primero para que el diccionario de abajo no falle
    horas_puntuables = min(10.0, b1_horas_siab) 
    
    # Parte A: Horas (2.5 pts máximo)
    puntos_horas = (horas_puntuables / 10.0) * 2.5 

    # Parte B: Servicios Especiales (2.5 pts máximo comparado con el mejor del mes)
    puntos_servicios = (b2_servicios_ruba / max_servicios_cuartel) * 2.5

    pilar_vocacion = min(5.0, round(puntos_horas + puntos_servicios, 2))

    # PILAR 2: TÉCNICA
    pilar_tecnica = min(5.0, b3_promedio_emergencia / 2)

    # PILAR 3: CUALIDADES (Fijo)
    cur.execute("""
        SELECT nota_cualidades 
        FROM calificaciones_cualidades 
        WHERE legajo = %s
    """, (legajo,))

    resultado = cur.fetchone()

    # Si existe nota, la usamos; si no, queda en 0.0
    pilar_cualidades = float(resultado['nota_cualidades']) if resultado else 0.0

    # PILAR 4: ASISTENCIA (50% Clases Oblig. / 50% Salidas Reales)
    puntos_clases = (b6_clases_oblig / 2) * 2.5
    puntos_salidas = (b7_siniestros_reales / max_cuartel) * 2.5
    pilar_asistencia = min(5.0, round(puntos_clases + puntos_salidas, 2))

    # PROMEDIO GENERAL (0-5)
    promedio_final = round((pilar_vocacion + pilar_tecnica + pilar_cualidades + pilar_asistencia) / 4, 2)

    # --- 4. DICCIONARIO PARA TEMPLATE ---
    datos_perfil = {
        **usuario,
        'b1_horas': b1_horas_siab,
        'b1_computables': horas_puntuables,
        'b2_servicios': b2_servicios_ruba,
        'b3_nota_emerg': b3_promedio_emergencia,
        'b4_practicas': 0.0,
        'b5_cualidades': pilar_cualidades,
        'b6_clases': b6_clases_oblig,
        'b7_siniestros': b7_siniestros_reales,
        'total_anio_actual': total_anio_actual,
        'total_anio_anterior': total_anio_anterior,
        'pilar_vocacion': round(pilar_vocacion, 2),
        'pilar_tecnica': round(pilar_tecnica, 2),
        'pilar_asistencia': pilar_asistencia,
        'pilar_cualidades': pilar_cualidades,
        'promedio_general': promedio_final,
        'max_cuartel': max_cuartel
    }

    cur.close()
    conn.close()
    return render_template("mi_perfil.html", datos=datos_perfil)

@login_requerido
@rol_requerido('ADMIN', 'JEFATURA', 'OFICIAL', 'SUB-OFICIAL')
@app.route("/mesa-calificadora", methods=["GET", "POST"])
def mesa_calificadora():
    # 1. Identificación del usuario
    mi_grado = session.get("grado", "").upper()
    mi_rol = session.get("rol", "").upper()
    mi_legajo = session.get("legajo")
    
    # 2. Control de Acceso (Solo rangos autorizados o Admin)
    autorizados = ['JEFATURA', 'OFICIAL', 'SUB-OFICIAL', 'ADMIN']
    if mi_grado not in autorizados and mi_rol != 'ADMIN':
        return "No tienes autorización para acceder a la Mesa Calificadora.", 403

    conn = get_db()
    cur = conn.cursor(dictionary=True)

    # 3. Procesar el Guardado (POST)
    if request.method == "POST":
        legajo_dest = request.form.get("legajo")
        nota = request.form.get("nota")
        obs = request.form.get("observacion")
        
        cur.execute("""
            INSERT INTO calificaciones_cualidades (legajo, nota_cualidades, observacion) 
            VALUES (%s, %s, %s)
            ON DUPLICATE KEY UPDATE nota_cualidades = %s, observacion = %s
        """, (legajo_dest, nota, obs, nota, obs))
        conn.commit()

    # 4. LÓGICA DE JERARQUÍA (El Filtro de visibilidad)
    if mi_rol == 'ADMIN' or mi_grado == 'JEFATURA':
        # El Jefe/Admin ve a todos (menos a sí mismo)
        query_filtro = "WHERE l.legajo != %s"
    elif mi_grado == 'OFICIAL':
        # El Oficial ve a Suboficiales y Bomberos (no ve Jefatura)
        query_filtro = "WHERE l.legajo != %s AND l.grado NOT IN ('JEFATURA')"
    elif mi_grado == 'SUB-OFICIAL':
        # El Suboficial solo ve a Bomberos y Aspirantes
        query_filtro = "WHERE l.legajo != %s AND l.grado NOT IN ('JEFATURA', 'OFICIAL', 'SUB-OFICIAL')"
    else:
        query_filtro = "WHERE 1=0" # Por seguridad, nadie más ve nada

    # 5. Consulta Final con comparación de performance
    # Consulta mejorada con datos históricos
    cur.execute(f"""
        SELECT 
            l.legajo, l.apellido, l.nombre, l.grado, 
            COALESCE(cc.nota_cualidades, 0) as nota_actual,
            COALESCE(cc.observacion, '') as observacion,
            COALESCE(cc.anio_anterior_puntos, 0) as nota_anterior
        FROM legajos l
        LEFT JOIN calificaciones_cualidades cc ON l.legajo = cc.legajo
        {query_filtro}
        ORDER BY l.apellido ASC
    """, (mi_legajo,))
    
    bomberos = cur.fetchall()
    conn.close()
    
    return render_template("mesa_calificadora.html", bomberos=bomberos)

@login_requerido
@rol_requerido('ADMIN', 'JEFATURA')
@app.route("/mesa-calificadora/cerrar-ciclo", methods=["POST"])
def cerrar_ciclo_anual():
    conn = get_db()
    if not conn:
        flash("Error de conexión a la base de datos.", "danger")
        return redirect(url_for('mesa_calificadora'))
    
    try:
        cur = conn.cursor()
        # 1. Pasamos la nota_actual a nota_anterior_puntos
        # 2. Seteamos nota_actual en 0 para el nuevo ciclo
        # 3. Limpiamos las observaciones para el nuevo año
        cur.execute("""
            UPDATE calificaciones_cualidades 
            SET anio_anterior_puntos = nota_cualidades,
                nota_cualidades = 0,
                observacion = ''
        """)
        
        conn.commit()
        flash("Ciclo anual cerrado con éxito. Las notas han sido archivadas en el historial.", "success")
    
    except Exception as e:
        conn.rollback()
        flash(f"Error al cerrar el ciclo: {str(e)}", "danger")
    
    finally:
        conn.close()
        
    return redirect(url_for('mesa_calificadora'))

@app.route("/dashboard")
@login_requerido
def dashboard():
    conn = get_db()
    cur = conn.cursor(dictionary=True)
    
    # Buscamos quiénes están para la baja (2 años < 2 pts)
    cur.execute("""
        SELECT l.apellido, l.nombre, l.legajo
        FROM calificaciones_cualidades cc
        JOIN legajos l ON cc.legajo = l.legajo
        WHERE cc.nota_cualidades < 2 
          AND cc.anio_anterior_puntos < 2 
          AND cc.nota_cualidades > 0
    """)
    personal_en_riesgo = cur.fetchall()
    conn.close()
    
    return render_template("dashboard.html", en_riesgo=personal_en_riesgo)

@app.route("/departamentos/ver/<int:id>")
@login_requerido
def ver_miembros(id):
    conn = get_db()
    departamento = {}
    miembros = []
    
    if conn:
        cur = conn.cursor(dictionary=True)
        # 1. Buscamos el nombre del departamento
        cur.execute("SELECT * FROM departamentos WHERE id = %s", (id,))
        departamento = cur.fetchone()
        
        # 2. Buscamos los miembros en la tabla real: bombero_departamento
        cur.execute("""
            SELECT l.apellido, l.nombre, l.grado, l.legajo
            FROM bombero_departamento bd
            JOIN legajos l ON bd.legajo = l.legajo
            WHERE bd.departamento_id = %s AND bd.activo = 1
            ORDER BY l.apellido, l.nombre
        """, (id,))
        miembros = cur.fetchall()
        conn.close()
        
    return render_template("ver_miembros.html", depto=departamento, miembros=miembros)

import os
import subprocess
from datetime import datetime
from flask import flash, redirect, url_for

@app.route('/actividades')
@login_requerido
def listado_actividades():
    try:
        conn = get_db()
        actividades = []
        if conn:
            cursor = conn.cursor(dictionary=True)
            # Traemos el nombre del concepto para mostrarlo como 'tipo' de actividad
            query = """
                SELECT 
                    e.id, 
                    e.fecha as fecha_inicio, 
                    c.concepto as tipo, 
                    e.descripcion, 
                    c.concepto as concepto_nombre,
                    5 as asignado 
                FROM eventos e 
                LEFT JOIN conceptos c ON e.concepto_id = c.id 
                WHERE e.estado = 'FINALIZADO'
                /* Aquí podrías filtrar para excluir capacitaciones si fuera necesario */
                AND c.concepto NOT IN ('CAPACITACIÓN', 'CURSO')
                ORDER BY e.fecha DESC
            """
            cursor.execute(query)
            actividades = cursor.fetchall()
            
            print(f"DEBUG SIAB: Actividades de cuartel encontradas: {len(actividades)}")
            
            cursor.close()
            conn.close()
        
        return render_template('actividades.html', actividades=actividades)

    except Exception as e:
        print(f"Error crítico en actividades: {e}")
        return f"Error detectado en el listado de actividades: {e}"
    
@app.route('/registro-salidas/listado')
@login_requerido
def listado_siniestros():
    db = get_db()
    cur = db.cursor(dictionary=True)
    
    # Filtramos para que NO muestre Servicios Especiales ni Capacitaciones
    # Suponiendo que las capacitaciones tienen un grupo específico o están vacías
    cur.execute("""
        SELECT id, nro_part_ruba, fecha, hora_salida, tipo_siniestro, lugar, estado 
        FROM nexo_siniestros 
        WHERE grupo_ruba NOT IN ('Servicios Especiales', 'Capacitaciones', '') 
        ORDER BY id DESC
    """)
    
    mis_datos = cur.fetchall()
    cur.close()
    db.close()
    
    return render_template('siniestros_listado.html', siniestros=mis_datos)

@app.route('/actividades/nueva', methods=['GET', 'POST'])
def nueva_actividad():
    conn = get_db()
    cursor = conn.cursor(dictionary=True)

    if request.method == 'POST':
        concepto_id = request.form['concepto_id']
        fecha = request.form['fecha']
        descripcion = request.form['descripcion']
        puntos = request.form['puntos']
        
        cursor.execute("""
            INSERT INTO actividades (concepto_id, fecha, descripcion, puntos) 
            VALUES (%s, %s, %s, %s)
        """, (concepto_id, fecha, descripcion, puntos))
        
        conn.commit()
        flash("Actividad registrada con éxito", "success")
        return redirect(url_for('listado_actividades'))

    # Para el GET: cargamos los conceptos para el desplegable
    cursor.execute("SELECT id, nombre FROM conceptos ORDER BY nombre")
    conceptos = cursor.fetchall()
    cursor.close()
    conn.close()
    return render_template('nueva_actividad.html', conceptos=conceptos)

@app.route("/mis-actividades")
@login_requerido
def mis_actividades():
    conn = get_db()
    lista_actividades = []
    
    if conn:
        cur = conn.cursor(dictionary=True)
        # Traemos las filas de la tabla asistencia que coinciden con el legajo del usuario logueado
        cur.execute("""
            SELECT a.id, e.fecha as fecha_inicio, c.concepto as concepto_nombre, a.calificacion as asignado
            FROM asistencia a
            JOIN eventos e ON a.evento_id = e.id
            LEFT JOIN conceptos c ON e.concepto_id = c.id
            WHERE a.legajo = %s
            ORDER BY e.fecha DESC
        """, (session.get('legajo'),))
        lista_actividades = cur.fetchall()
        conn.close()
    
    # Aquí es donde le "pasamos" la lista a la plantilla
    return render_template("actividades.html", actividades=lista_actividades)

@app.route("/actividades")
@login_requerido
def ver_actividades():
    conn = get_db()
    actividades = []
    stats = {"total": 0, "horas": 0, "ranking": []}
    
    if conn:
        try:
            cur = conn.cursor(dictionary=True)
            # 1. Listado de Actividades Finalizadas (Excluyendo capacitaciones según tu lógica)
            cur.execute("""
                SELECT e.id, e.fecha as fecha_inicio, c.concepto as tipo, 
                       e.descripcion, c.concepto as concepto_nombre
                FROM eventos e 
                LEFT JOIN conceptos c ON e.concepto_id = c.id 
                WHERE e.estado = 'FINALIZADO' 
                AND c.concepto NOT IN ('CAPACITACIÓN', 'CURSO')
                ORDER BY e.fecha DESC LIMIT 10
            """)
            actividades = cur.fetchall()

            # 2. Estadísticas para los Gráficos (Dashboard)
            # Contamos cuántas actividades hubo por concepto
            cur.execute("""
                SELECT c.concepto, COUNT(e.id) as cantidad
                FROM eventos e
                JOIN conceptos c ON e.concepto_id = c.id
                WHERE e.estado = 'FINALIZADO'
                GROUP BY c.concepto
            """)
            stats['ranking'] = cur.fetchall()
            stats['total'] = len(actividades)
            
        finally:
            conn.close()
    
    return render_template("actividades_dashboard.html", actividades=actividades, stats=stats)

from datetime import datetime, date
import calendar

from datetime import datetime, date
import calendar

@app.route("/mis-capacitaciones")
@login_requerido
def mis_capacitaciones():
    legajo = session.get('legajo')
    rol_usuario = session.get('rol')
    
    hoy = date.today()
    anio_actual = hoy.year
    anio_anterior = anio_actual - 1

    db = get_db()
    cur = db.cursor(dictionary=True)

    # A. TOTAL QUE DICTÓ EL CUARTEL (El denominador)
    cur.execute("""
        SELECT COUNT(*) as total 
        FROM eventos 
        WHERE tipo = 'CAPACITACION' 
        AND estado = 'FINALIZADO' 
        AND YEAR(fecha) = %s
    """, (anio_actual,))
    total_dictadas = cur.fetchone()['total'] or 0

    # B. LISTADO DE ASISTENCIAS DEL BOMBERO (Para la tabla)
    cur.execute("""
        SELECT e.fecha, e.descripcion as tema, a.calificacion, a.estado, a.observacion
        FROM asistencia a
        JOIN eventos e ON a.evento_id = e.id
        WHERE a.legajo = %s 
        AND e.tipo = 'CAPACITACION' 
        AND e.estado != 'ANULADO'
        AND YEAR(e.fecha) = %s
        ORDER BY e.fecha DESC
    """, (legajo, anio_actual))
    registros = cur.fetchall()

    # C. CÁLCULO DE PORCENTAJE
    asistidas = len([r for r in registros if r['estado'] == 'PRESENTE'])
    porcentaje = round((asistidas / total_dictadas * 100), 1) if total_dictadas > 0 else 0

    # D. RENDIMIENTO AÑO ANTERIOR
    cur.execute("""
        SELECT 
            COUNT(*) as total_ant,
            SUM(CASE WHEN a.estado = 'PRESENTE' THEN 1 ELSE 0 END) as asistidas_ant
        FROM asistencia a
        JOIN eventos e ON a.evento_id = e.id
        WHERE a.legajo = %s AND e.tipo = 'CAPACITACION' 
        AND e.estado = 'FINALIZADO' AND YEAR(e.fecha) = %s
    """, (legajo, anio_anterior))
    res_ant = cur.fetchone()
    
    porcentaje_anterior = round((res_ant['asistidas_ant'] / res_ant['total_ant'] * 100), 1) if res_ant and res_ant['total_ant'] > 0 else "N/A"

    return render_template("detalle_capacitaciones.html", 
                           registros=registros, 
                           porcentaje=porcentaje,
                           porcentaje_anterior=porcentaje_anterior,
                           total_dictadas=total_dictadas,
                           asistidas=asistidas,
                           anio_actual=anio_actual)

@app.route('/mis-salidas')
@login_requerido
def ver_mis_salidas():
    db = get_db()
    cur = db.cursor(dictionary=True)
    
    # Obtenemos el legajo del usuario logueado
    mi_legajo = session.get('legajo') 

    # 1. Ejecutamos la consulta
    cur.execute("""
        SELECT 
            s.id, 
            s.fecha, 
            s.tipo_siniestro, 
            s.nro_part_ruba,
            p.rol, 
            p.movil,
            p.puntos_operativos
        FROM nexo_siniestros s
        JOIN nexo_personal p ON s.id = p.siniestro_id
        WHERE p.legajo = %s
        ORDER BY s.fecha DESC
    """, (mi_legajo,))
    
    # 2. AHORA DEFINIMOS salidas_raw
    salidas_raw = cur.fetchall()
    
    # 3. AHORA SÍ PODEMOS HACER EL PRINT
    print(f"DEBUG: Se encontraron {len(salidas_raw)} salidas para el legajo {mi_legajo}")

    # Procesamos las salidas (esto es lo que ya tenías)
    salidas_procesadas = []
    total_puntos = 0
    
    for s in salidas_raw:
        # Por ahora todos los pesos son 1.0
        s['puntaje_final'] = float(s['puntos_operativos'] or 0)
        total_puntos += s['puntaje_final']
        salidas_procesadas.append(s)

    cur.close()
    db.close()
    
    return render_template('mis_salidas.html', 
                           salidas=salidas_procesadas, 
                           total_puntos=total_puntos,
                           cantidad_salidas=len(salidas_procesadas))

@app.route("/mis-actividades-gestion")
@login_requerido
def mis_actividades_gestion():
    legajo = session.get('legajo')
    db = get_db()
    cur = db.cursor(dictionary=True)
    
    # Solo horas de gestión y prácticas (Pilar Vocación)
    cur.execute("""
        SELECT fecha_inicio as fecha, actividad, descripcion, horas
        FROM actividades 
        WHERE legajo = %s AND actividad NOT IN ('CAPACITACIÓN_OBLIGATORIA') AND anulada = 0
        ORDER BY fecha_inicio DESC
    """, (legajo,))
    registros = cur.fetchall()
    total_horas = sum(float(r['horas'] or 0) for r in registros)

    return render_template("detalle_actividades.html", 
                           registros=registros, 
                           total_horas=total_horas)

def calcular_pilar_vocacion(legajo, mes, anio):
    # Formateamos el filtro de fecha para el mes actual (ej: "04/2026")
    filtro_fecha = f"%/{mes:02d}/{anio}"
    
    cur.execute("""
        SELECT SUM(horas) as total_hs 
        FROM actividades 
        WHERE legajo = %s 
        AND fecha_inicio LIKE %s 
        AND estado = 'ACTIVA' 
        AND anulada = 0
    """, (legajo, filtro_fecha))
    
    resultado = cur.fetchone()
    horas_totales = float(resultado['total_hs'] or 0.0)
    
    # Ejemplo de escala: 10 horas mensuales = 5 puntos (máximo)
    puntaje = min(5.0, (horas_totales / 10) * 5)
    
    return puntaje, horas_totales

from datetime import datetime

@app.route('/planilla-nexo/nueva', methods=['GET', 'POST'])
@login_requerido
def nueva_planilla_nexo():
    db = get_db()
    cur = db.cursor(dictionary=True)

    if request.method == 'POST':
        try:
            nro_parte = request.form.get('nro_parte')
            # Ahora recibimos el ID del mapeo
            tipo_mapeo_id = request.form.get('tipo_siniestro_id') 
            lugar = request.form.get('lugar')
            fecha = request.form.get('fecha') 
            hora = request.form.get('hora_salida')

            # BUSCAMOS LOS DATOS DEL RUBA ANTES DE GUARDAR
            cur.execute("SELECT * FROM tipos_siniestros WHERE id = %s", (tipo_mapeo_id,))
            mapeo = cur.fetchone()

            if not mapeo:
                raise Exception("Tipo de siniestro no válido")

            # INSERTAMOS EN LA TABLA DE SALIDAS (Agregando los campos RUBA)
            # Asegúrate de haber agregado las columnas grupo_ruba y subtipo_ruba a tu tabla
            sql_siniestro = """
                INSERT INTO nexo_siniestros 
                (nro_part_ruba, fecha, hora_salida, tipo_siniestro, grupo_ruba, subtipo_ruba, lugar, estado)
                VALUES (%s, %s, %s, %s, %s, %s, %s, 'BORRADOR')
            """
            cur.execute(sql_siniestro, (
                nro_parte, 
                fecha, 
                hora, 
                mapeo['nombre_siab'],   # Detalle SIAB
                mapeo['grupo_ruba'],    # Grupo RUBA (Ej: Incendios)
                mapeo['subtipo_ruba'],  # Subtipo RUBA (Ej: Forestal)
                lugar
            ))
            
            siniestro_id = cur.lastrowid

            bomberos_seleccionados = request.form.getlist('bomberos_seleccionados')
            for legajo in bomberos_seleccionados:
                movil = request.form.get(f'movil_{legajo}')
                rol = request.form.get(f'rol_{legajo}')
                cur.execute("INSERT INTO nexo_personal (siniestro_id, legajo, movil, rol) VALUES (%s, %s, %s, %s)", 
                            (siniestro_id, legajo, movil, rol))

            db.commit()
            flash(f"Registro #{siniestro_id} guardado con éxito.", "success")
            return redirect(url_for('listado_siniestros'))
            
        except Exception as e:
            db.rollback()
            print(f"Error al guardar: {e}")
            flash(f"Error al guardar: {str(e)}", "danger")
            return redirect(url_for('nueva_planilla_nexo'))

    # --- MÉTODO GET: CARGA DE FORMULARIO ---
    
    # 1. Cargamos los tipos de siniestros mapeados (IMPORTANTE: incluir grupo_ruba y el ORDER BY)
    cur.execute("""
        SELECT id, nombre_siab, grupo_ruba 
        FROM tipos_siniestros 
        ORDER BY grupo_ruba ASC, nombre_siab ASC
    """)
    res_tipos = cur.fetchall()

    # 2. Cargamos el personal (el resto sigue igual)
    cur.execute("""
        SELECT legajo, nombre, apellido, situacion 
        FROM legajos 
        WHERE situacion LIKE '%ACTIVO%' OR situacion LIKE '%RESERVA%'
        ORDER BY apellido ASC
    """)
    res_bomberos = cur.fetchall()
    
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    hora_hoy = datetime.now().strftime('%H:%M')
    
    cur.close()
    db.close()
    
    return render_template('nexo_form.html', 
                           bomberos=res_bomberos, 
                           tipos_siniestros=res_tipos, # <-- Pasamos los tipos al HTML
                           fecha_hoy=fecha_hoy, 
                           hora_hoy=hora_hoy)
    
@app.route('/planilla-nexo/imprimir/<int:id>')
@login_requerido
def imprimir_nexo(id):
    db = get_db()
    cur = db.cursor(dictionary=True)
    
    # Traemos datos del siniestro
    cur.execute("SELECT * FROM nexo_siniestros WHERE id = %s", (id,))
    siniestro = cur.fetchone()
    
    # Traemos los bomberos asociados con sus nombres
    cur.execute("""
        SELECT p.*, l.nombre, l.apellido, l.grado as jerarquia 
        FROM nexo_personal p
        JOIN legajos l ON p.legajo = l.legajo
        WHERE p.siniestro_id = %s
    """, (id,))
    personal = cur.fetchall()
    
    return render_template('nexo_reporte.html', siniestro=siniestro, personal=personal)

    # ==================================================
    # PLANILLA NEXO PARA FIRMAS
    # ==================================================
    def exportar_planilla_nexo(self, file_path, siniestro, personal):
        from reportlab.platypus import Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.units import cm

        elementos = []
        estilo_texto = self.styles['Normal']
        estilo_texto.fontSize = 10

        # --- 1. BLOQUE DATOS DEL SINIESTRO ---
        datos_siniestro = [
            [f"Nro Parte RUBA: {siniestro['nro_part_ruba']}", f"Fecha: {siniestro['fecha']}"],
            [f"Tipo: {siniestro['tipo_siniestro']}", f"Lugar: {siniestro['lugar']}"],
        ]
        t_sin = Table(datos_siniestro, colWidths=[8*cm, 8*cm])
        t_sin.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold')]))
        elementos.append(t_sin)
        elementos.append(Spacer(1, 0.5 * cm))

        # --- 2. TABLA DE PERSONAL ---
        headers = ['LEGAJO', 'APELLIDO Y NOMBRE', 'MÓVIL', 'ROL', 'FIRMA']
        data = [headers]
        
        for p in personal:
            data.append([
                p['legajo'],
                f"{p['apellido']} {p['nombre']}",
                p['movil'],
                p['rol'],
                "________________" # Espacio para firma física
            ])

        tabla = Table(data, repeatRows=1, colWidths=[2*cm, 6*cm, 2.5*cm, 3.5*cm, 3*cm])
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#a50000")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8), # Más espacio para la firma
            ('TOPPADDING', (0,0), (-1,-1), 8),
        ]))
        elementos.append(tabla)
        elementos.append(Spacer(1, 1.5 * cm))

        # --- 3. BLOQUE DE FIRMAS AUTORIDAD ---
        firmas_data = [
            ["___________________________", "___________________________"],
            ["Firma Jefe de Dotación", "Firma Oficial de Turno"]
        ]
        t_firmas = Table(firmas_data, colWidths=[8*cm, 8*cm])
        t_firmas.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,1), (-1,1), 9),
        ]))
        elementos.append(t_firmas)

        # Usamos tu base unificada para generar el PDF con logo y header
        self.crear_pdf_unificado(file_path, elementos, "PLANILLA NEXO DE INTERVENCIÓN")

from pdf_manager import PDFManager

@app.route('/imprimir-nexo/<int:id>')
def ruta_imprimir_nexo(id):
    # 1. Obtener datos de la DB (siniestro y personal)
    # ... (tus consultas SQL aquí) ...
    
    # 2. Configurar PDFManager
    manager = PDFManager(session) # session tiene legajo, nombre, apellido
    
    filename = f"nexo_{id}.pdf"
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    # 3. Generar el archivo
    manager.exportar_planilla_nexo(file_path, siniestro_data, personal_data)
    
    return send_file(file_path, as_attachment=False)

@app.route('/registro-salidas/calificar/<int:id>', methods=['GET', 'POST']) # Agregado methods
@login_requerido
def calificar_salida(id):
    rol_usuario = session.get('rol')
    permisos_autorizados = ['ADMIN', 'ENCARGADO', 'SUPERVISOR']

    if rol_usuario not in permisos_autorizados:
        flash("No tenés permisos para calificar salidas.", "danger")
        return redirect(url_for('listado_siniestros'))
    
    db = get_db()
    cur = db.cursor(dictionary=True)

    if request.method == 'POST':
        try:
            legajos = request.form.getlist('legajos')
            # Capturamos el comentario general que agregaste abajo
            comentario_general = request.form.get('observaciones')
            
            for legajo in legajos:
                # Buscamos el puntaje específico de cada legajo
                puntos = request.form.get(f'puntos_{legajo}')
                
                cur.execute("""
                    UPDATE nexo_personal 
                    SET puntos_operativos = %s 
                    WHERE siniestro_id = %s AND legajo = %s
                """, (puntos, id, legajo))

            # Cerramos el siniestro y guardamos la observación general
            cur.execute("""
                UPDATE nexo_siniestros 
                SET estado = 'CALIFICADO', 
                    comentario_general = %s 
                WHERE id = %s
            """, (comentario_general, id))

            db.commit()
            flash("Calificación guardada. El personal ha recibido sus puntos operativos.", "success")
            return redirect(url_for('listado_siniestros'))
        except Exception as e:
            db.rollback()
            flash(f"Error al procesar puntos: {e}", "danger")

    # --- MÉTODO GET ---
    anio_siniestro = siniestro['fecha'].year # O usa datetime.now().year
    cur.execute("SELECT puntos_por_asistencia FROM config_puntos WHERE anio = %s", (anio_siniestro,))
    config = cur.fetchone()

    # Si no hay configuración para ese año, usamos un valor por defecto (ej: 1.0)
    puntos_sugeridos = config['puntos_por_asistencia'] if config else 1.0

    # Ajustado para que use tu tabla 'legajos'
    cur.execute("""
        SELECT p.*, l.apellido, l.nombre 
        FROM nexo_personal p
        JOIN legajos l ON p.legajo = l.legajo
        WHERE p.siniestro_id = %s
    """, (id,))
    personal = cur.fetchall()

    cur.close()
    return render_template('nexo_calificar.html', siniestro=siniestro, personal=personal)     

# ============================================================
# PANEL SISTEMA
# ============================================================
@app.route('/configuracion')
@login_requerido
@rol_requerido("ADMIN")
def panel_sistema():
    # Esta es la ruta que llama al nuevo template
    return render_template("admin_sistema.html")

@app.route('/admin/backup')
@login_requerido # Usando tus decoradores de seguridad
@rol_requerido("ADMIN")
def ejecutar_backup():
    try:
        db_user = DB_CONFIG['user']
        db_pass = DB_CONFIG['password']
        db_name = DB_CONFIG['database']
        
        # Ruta consistente con tu instalador SIAB
        folder = r"C:\SIAB\backups"
        if not os.path.exists(folder): 
            os.makedirs(folder)

        fecha = datetime.now().strftime("%Y-%m-%d_%H-%M")
        filename = f"backup_{db_name}_{fecha}.sql"
        filepath = os.path.join(folder, filename)

        # Buscador de binarios (Excelente lógica de portabilidad)
        posibles_rutas = [
            r"C:\xampp\mysql\bin\mysqldump.exe",
            r"C:\Program Files\MySQL\MySQL Server 8.0\bin\mysqldump.exe",
            r"C:\Program Files\MySQL\MySQL Server 8.1\bin\mysqldump.exe",
            "mysqldump"
        ]
        
        dump_exe = next((r for r in posibles_rutas if r == "mysqldump" or os.path.exists(r)), None)

        if not dump_exe:
            flash("No se encontró mysqldump.exe. Verificá la instalación.", "danger")
            return redirect(url_for('panel_sistema'))

        # Ejecución
        comando = [dump_exe, f"--user={db_user}", f"--password={db_pass}", db_name]

        with open(filepath, "w") as out_file:
            resultado = subprocess.run(comando, stdout=out_file, stderr=subprocess.PIPE, text=True)

        # Validamos éxito
        if resultado.returncode != 0:
            if os.path.exists(filepath): os.remove(filepath)
            flash(f"Error de MySQL: {resultado.stderr}", "danger")
        else:
            flash(f"¡Respaldo exitoso! Archivo: {filename}", "success")

    except Exception as e:
        flash(f"Error crítico: {str(e)}", "danger")
    
    # Redirigimos de vuelta al panel de control, no al inicio
    return redirect(url_for('panel_sistema'))

# ============================================================
# ACADEMIA BOMBERO
# ============================================================

@app.route("/academia/bombero/<int:legajo>")
@login_requerido
def ver_academia_bombero(legajo):
    conn = get_db()
    if not conn:
        flash("Error de conexión.", "danger")
        return redirect(url_for('inicio'))

    try:
        cur = conn.cursor(dictionary=True)

        # 1. Datos del bombero (Tabla: legajos)
        cur.execute("SELECT legajo, nombre, apellido, grado, cargo, foto FROM legajos WHERE legajo = %s", (legajo,))
        bombero = cur.fetchone()

        if not bombero:
            flash(f"Legajo {legajo} no encontrado.", "warning")
            return redirect(url_for('bomberos'))

        bombero['jerarquia'] = bombero['grado']

        # 2. Notas de ACADEMIA (Ajustado a tu estructura real)
        # Filtramos por 'PRESENTE' ya que tu ENUM no tiene 'ANULADA'
        cur.execute("""
            SELECT et.nombre AS descripcion, ant.nota, a.fecha_registro as fecha
            FROM asistencia_notas_temas ant
            JOIN evento_temas et ON ant.tema_id = et.id
            JOIN asistencia a ON ant.evento_id = a.evento_id AND ant.legajo = a.legajo
            WHERE ant.legajo = %s AND a.estado = 'PRESENTE'
            ORDER BY a.fecha_registro ASC 
        """, (legajo,))
        notas_academia = cur.fetchall()

        # 3. Puntos de SALIDAS (Intervenciones reales confirmadas)
        cur.execute("""
            SELECT s.fecha, s.tipo_siniestro as descripcion, p.rol, p.puntos_operativos as puntos
            FROM nexo_personal p
            JOIN nexo_siniestros s ON p.siniestro_id = s.id
            WHERE p.legajo = %s AND p.firma_confirmada = 1
            ORDER BY s.fecha DESC
        """, (legajo,))
        puntos_salidas = cur.fetchall()

        # 4. Cálculos de Totales y Preparación de Gráfica
        # Nota: Usamos float() para asegurar compatibilidad con Chart.js
        notas_validas = [float(n['nota']) for n in notas_academia if n['nota'] is not None]
        promedio = round(sum(notas_validas) / len(notas_validas), 2) if notas_validas else 0
        
        total_puntos_operativos = sum(p['puntos'] for p in puntos_salidas)
        total_salidas = len(puntos_salidas)

        # Listas para Chart.js
        fechas_grafica = [n['fecha'].strftime('%d/%m') for n in notas_academia if n['fecha']]
        valores_grafica = [float(n['nota']) for n in notas_academia if n['nota'] is not None]

        return render_template("academia_bombero.html", 
                               bombero=bombero, 
                               notas_academia=notas_academia[::-1], 
                               puntos_salidas=puntos_salidas,
                               promedio=promedio,
                               total_salidas=total_salidas,
                               total_puntos_operativos=total_puntos_operativos,
                               fechas_grafica=fechas_grafica,
                               valores_grafica=valores_grafica,
                               jefe_dotacion={"nombre": "Firma Autorizada", "jerarquia": "Jefatura"},
                               cuartelero={"nombre": "Cuartelero de Turno", "jerarquia": "Guardia"})

    except Exception as e:
        flash(f"Error técnico en Academia: {e}", "danger")
        return redirect(url_for('inicio'))
    finally:
        if conn: conn.close()

# ============================================================
# PERMISOS
# ============================================================
def tiene_permiso(area):
    # En lugar de current_user, usamos la session de Flask
    if "usuario_id" not in session:
        return False
    
    rol = str(session.get("rol", "")).upper()
    
    # El admin siempre tiene permiso
    if rol == 'ADMIN':
        return True
    
    # Verificamos si es encargado del área (usando el permiso de la sesión)
    # Suponiendo que guardas los permisos como 'es_encargado_moviles', etc.
    permiso_buscado = f'es_encargado_{area}'
    return session.get(permiso_buscado) == 1

# ============================================================
# CONTROL DE ACCESO Y PERMISOS
# ============================================================

def requerir_permiso(permiso_requerido):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if "usuario_id" not in session:
                flash("Debés iniciar sesión.", "warning")
                return redirect(url_for("login"))
            
            # CORRECCIÓN: Forzamos mayúsculas para que 'ADMIN' sea igual a 'admin'
            rol_actual = str(session.get("rol", "")).upper()
            
            # El administrador siempre tiene acceso total
            if rol_actual == 'ADMIN':
                return f(*args, **kwargs)
            
            # Verificación de permisos específicos para otros roles
            conn = get_db()
            cur = conn.cursor(dictionary=True)
            cur.execute(f"SELECT {permiso_requerido} FROM usuarios WHERE id = %s", (session["usuario_id"],))
            usuario = cur.fetchone()
            conn.close()

            if usuario and usuario.get(permiso_requerido) == 1:
                return f(*args, **kwargs)
            
            flash("Acceso restringido: No eres encargado de esta área.", "danger")
            return redirect(url_for('inicio'))
        return decorated_function
    return decorator

@app.route("/admin/asignar_cargo", methods=['POST'])
@requerir_permiso('es_encargado_admin')
def asignar_cargo():
    legajo = request.form.get('legajo')
    area = request.form.get('area')
    
    # 1. Actualizamos el permiso activo para el login
    # (Ej: SET es_encargado_moviles = 1)
    db.execute(f"UPDATE usuarios SET es_encargado_{area} = 1 WHERE legajo = {legajo}")
    
    # 2. Guardamos en el historial para el Legajo Digital
    db.execute("INSERT INTO historial_cargos (legajo, area, fecha_inicio) VALUES (%s, %s, CURDATE())", (legajo, area))
    
    flash("Cargo asignado y registrado en el historial.", "success")
    return redirect(url_for('gestion_personal'))

@app.route("/guardar_cargo", methods=['POST'])
@login_requerido
@requerir_permiso('es_encargado_admin')
def guardar_cargo():
    legajo = request.form.get('legajo')
    area = request.form.get('area')
    observaciones = request.form.get('observaciones')
    columna_permiso = f"es_encargado_{area}"
    
    conn = get_db()
    if conn:
        try:
            cur = conn.cursor()
            # Actualizar acceso
            cur.execute(f"UPDATE usuarios SET {columna_permiso} = 1 WHERE legajo = %s", (legajo,))
            # Registrar historial
            cur.execute("""
                INSERT INTO historial_cargos (legajo, area, fecha_inicio, observaciones) 
                VALUES (%s, %s, CURDATE(), %s)
            """, (legajo, area, observaciones))
            conn.commit()
            flash(f"Cargo de {area} asignado con éxito.", "success")
        except Error as e:
            conn.rollback()
            flash(f"Error: {str(e)}", "danger")
        finally:
            conn.close()
    return redirect(url_for('ver_gestion_cargos'))

@app.route("/finalizar_cargo", methods=['POST'])
@login_requerido
@requerir_permiso('es_encargado_admin')
def finalizar_cargo():
    id_historial = request.form.get('id_historial')
    legajo = request.form.get('legajo')
    area = request.form.get('area') # Ej: 'moviles'
    
    columna_permiso = f"es_encargado_{area}"
    
    try:
        # 1. QUITAMOS el acceso inmediato en la tabla USUARIOS
        query_user = f"UPDATE usuarios SET {columna_permiso} = 0 WHERE legajo = %s"
        db.execute(query_user, (legajo,))
        
        # 2. REGISTRAMOS la fecha de fin en el historial
        query_historial = """
            UPDATE historial_cargos 
            SET fecha_fin = CURDATE() 
            WHERE id = %s
        """
        db.execute(query_historial, (id_historial,))
        
        db.commit()
        flash(f"Se ha finalizado la tarea de {area} para el legajo {legajo}.", "success")
        
    except Exception as e:
        db.rollback()
        flash(f"Error al finalizar cargo: {str(e)}", "danger")

    return redirect(url_for('ver_gestion_cargos'))

@app.route("/admin/gestion_cargos")
@login_requerido
@requerir_permiso('es_encargado_admin')
def ver_gestion_cargos():
    conn = get_db()
    cargos_activos = []
    bomberos = []
    
    if conn:
        try:
            cur = conn.cursor(dictionary=True)
            
            # 1. Traer cargos actuales (donde fecha_fin es NULL)
            # Unimos con la tabla legajos para tener nombre y apellido reales
            query_activos = """
                SELECT h.id, h.legajo, h.area, h.fecha_inicio, l.nombre, l.apellido 
                FROM historial_cargos h
                JOIN legajos l ON h.legajo = l.legajo
                WHERE h.fecha_fin IS NULL
                ORDER BY l.apellido ASC
            """
            cur.execute(query_activos)
            cargos_activos = cur.fetchall()
            
            # 2. Traer lista de bomberos activos para el selector del formulario
            cur.execute("SELECT legajo, nombre, apellido FROM legajos WHERE situacion = 'ACTIVO' ORDER BY apellido ASC")
            bomberos = cur.fetchall()
            
        except Error as e:
            flash(f"Error al cargar datos: {str(e)}", "danger")
        finally:
            conn.close()
            
    return render_template("gestion_cargos.html", cargos_activos=cargos_activos, bomberos=bomberos)

# ============================================================
# MOVILES
# ============================================================

@app.route("/moviles")
@login_requerido
@requerir_permiso('es_encargado_moviles')
def gestion_moviles():
    # Eliminamos el check de 'moviles' si ya usas 'es_encargado_moviles' en el decorador
    
    db = get_db()
    # Usamos dictionary=True para que en el HTML podamos usar m['nombre']
    cursor = db.cursor(dictionary=True)
    
    # Traemos todos los móviles de la tabla
    cursor.execute("SELECT * FROM moviles")
    todos_los_moviles = cursor.fetchall()
    db.close()
    
    # Separamos las listas según el estado
    activos = [m for m in todos_los_moviles if m['estado'] in ['ACTIVO', 'REPARACION']]
    historicos = [m for m in todos_los_moviles if m['estado'] in ['HISTORICO', 'BAJA']]
    
    return render_template('gestion_moviles.html', 
                           moviles=activos, 
                           moviles_historicos=historicos)

@app.route("/moviles/crear", methods=["POST"])
@login_requerido
@rol_requerido('ADMIN', 'JEFATURA')
def crear_movil():
    if request.method == "POST":
        # Captura de datos del formulario
        datos = (
            request.form.get("nro_unidad"),
            request.form.get("dominio"),
            request.form.get("anio"),
            request.form.get("nombre_homenaje"),
            request.form.get("marca"),
            request.form.get("modelo"),
            request.form.get("tipo"),
            request.form.get("lugar_origen"),
            request.form.get("proveedor"),
            request.form.get("capacidad_agua"),
            request.form.get("tiene_espuma"),
            request.form.get("capacidad_personas"),
            request.form.get("fecha_vtv") or None, # Manejo de fechas vacías
            request.form.get("fecha_compra") or None,
            request.form.get("trailer_dominio"),
            request.form.get("trailer_ejes")
        )

        conn = get_db()
        cur = conn.cursor()
        
        sql = """INSERT INTO unidades_fisicas 
                 (nro_unidad, dominio, anio, nombre_homenaje, marca, modelo, tipo, 
                  lugar_origen, proveedor, capacidad_agua, tiene_espuma, 
                  capacidad_personas, fecha_vtv, fecha_compra, trailer_dominio, trailer_ejes) 
                 VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
        
        try:
            cur.execute(sql, datos)
            conn.commit()
            flash("Nueva unidad registrada exitosamente en el historial.", "success")
        except Exception as e:
            conn.rollback()
            flash(f"Error al registrar: {str(e)}", "danger")
        finally:
            conn.close()

        return redirect(url_for('gestion_moviles')) # O como se llame tu ruta de listado

@app.route('/moviles/editar/<int:id>', methods=['GET', 'POST'])
@login_requerido
@requerir_permiso('es_encargado_moviles')
def editar_movil(id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    km_ini = request.form.get('km_inicial') or 0

    if request.method == 'POST':
        # Captura masiva de datos para supervisores
        query = """
            UPDATE moviles SET 
                nro_unidad=%s, dominio=%s, nombre_homenaje=%s, marca=%s, modelo=%s, 
                anio=%s, tipo=%s, lugar_origen=%s, proveedor=%s, capacidad_agua=%s, 
                tiene_espuma=%s, capacidad_personas=%s, fecha_vtv=%s, fecha_compra=%s, 
                trailer_dominio=%s, trailer_ejes=%s, estado=%s, km_inicial=%s
            WHERE id=%s
        """
        valores = (
            request.form.get('nro_unidad'), request.form.get('dominio'), 
            request.form.get('nombre_homenaje'), request.form.get('marca'), 
            request.form.get('modelo'), request.form.get('anio') or None, 
            request.form.get('tipo'), request.form.get('lugar_origen'), 
            request.form.get('proveedor'), request.form.get('capacidad_agua') or 0, 
            request.form.get('tiene_espuma'), request.form.get('capacidad_personas') or 0, 
            request.form.get('fecha_vtv') or None, request.form.get('fecha_compra') or None, 
            request.form.get('trailer_dominio'), request.form.get('trailer_ejes') or 0, 
            request.form.get('estado'), id
        )
        
        cursor.execute(query, valores)
        db.commit()
        db.close()
        return redirect(url_for('gestion_moviles'))

    cursor.execute("SELECT * FROM moviles WHERE id = %s", (id,))
    movil = cursor.fetchone()
    db.close()
    return render_template('editar_movil.html', movil=movil)

@app.route('/moviles/mantenimiento/registrar/<int:id_movil>', methods=['POST'])
@login_requerido
@requerir_permiso('es_encargado_moviles')
def registrar_mantenimiento(id_movil):
    if request.method == 'POST':
        fecha = request.form.get('fecha')
        tipo = request.form.get('tipo')
        desc = request.form.get('descripcion')
        prov = request.form.get('proveedor')
        km = request.form.get('km') or 0
        importe = request.form.get('importe') or 0
        proximo = request.form.get('proximo_vence') or None
        obs = request.form.get('observaciones')

        db = get_db()
        cursor = db.cursor()
        # 1. Insertamos el historial
        query = """INSERT INTO historial_mantenimiento 
                (id_movil, fecha_reparacion, tipo_mantenimiento, descripcion, 
                    proveedor, km_unidad, importe_total, fecha_proximo_mantenimiento, observaciones) 
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"""
        cursor.execute(query, (id_movil, fecha, tipo, desc, prov, km, importe, proximo, obs))
        
        # 2. AGREGADO: Actualizamos el KM en la tabla principal del móvil
        cursor.execute("UPDATE moviles SET km_inicial = %s WHERE id = %s", (km, id_movil))
        db.commit()
        db.close()
        return redirect(url_for('ficha_integral_movil', id=id_movil))

@app.route('/mantenimiento/actualizar/<int:id_reg>', methods=['POST'])
@login_requerido  # Usando tu decorador correcto
def actualizar_mantenimiento(id_reg):
    # Capturamos TODOS los campos del formulario
    fecha = request.form.get('fecha')
    km = request.form.get('km')
    tipo = request.form.get('tipo')
    desc = request.form.get('descripcion')
    prov = request.form.get('proveedor')
    importe = request.form.get('importe')
    obs = request.form.get('observaciones')

    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    # Buscamos el id_movil para saber a qué ficha volver
    cursor.execute("SELECT id_movil FROM historial_mantenimiento WHERE id = %s", (id_reg,))
    m = cursor.fetchone()
    id_movil = m['id_movil']

    # Actualizamos el registro completo en la tabla
    sql = """
        UPDATE historial_mantenimiento 
        SET fecha_reparacion=%s, km_unidad=%s, tipo_mantenimiento=%s, 
            descripcion=%s, proveedor=%s, importe_total=%s, observaciones=%s
        WHERE id=%s
    """
    cursor.execute(sql, (fecha, km, tipo, desc, prov, importe, obs, id_reg))
    
    # Opcional: También podrías actualizar el KM en la tabla moviles si es el registro más nuevo
    
    db.commit()
    db.close()

    return redirect(url_for('ficha_integral_movil', id=id_movil))
    
# --- ELIMINAR MANTENIMIENTO ---
@app.route('/moviles/mantenimiento/eliminar/<int:id_reg>')
@login_requerido
@requerir_permiso('es_encargado_moviles')
def eliminar_mantenimiento(id_reg):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    # Primero obtenemos el id_movil para saber a dónde volver
    cursor.execute("SELECT id_movil FROM historial_mantenimiento WHERE id = %s", (id_reg,))
    registro = cursor.fetchone()
    
    if registro:
        id_movil = registro['id_movil']
        cursor.execute("DELETE FROM historial_mantenimiento WHERE id = %s", (id_reg,))
        db.commit()
        flash("Registro de mantenimiento eliminado.", "warning")
    
    db.close()
    return redirect(url_for('ficha_integral_movil', id=id_movil))

# --- EDITAR MANTENIMIENTO (VISTA) ---
@app.route('/moviles/mantenimiento/editar/<int:id_reg>')
@login_requerido
def editar_mantenimiento(id_reg):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT * FROM historial_mantenimiento WHERE id = %s", (id_reg,))
    mantenimiento = cursor.fetchone()
    db.close()
    
    return render_template('editar_mantenimiento.html', m=mantenimiento)    

@app.route("/moviles/historial/<int:id>")
@login_requerido
@requerir_permiso('es_encargado_moviles')
def historial_movil(id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    # Obtenemos los datos del móvil para el encabezado
    cursor.execute("SELECT nro_unidad, marca, modelo FROM moviles WHERE id = %s", (id,))
    movil = cursor.fetchone()
    
    # Obtenemos todo el historial de reparaciones y servicios
    cursor.execute("""
        SELECT * FROM historial_mantenimiento 
        WHERE id_movil = %s 
        ORDER BY fecha_reparacion DESC
    """, (id,))
    historial = cursor.fetchall()
    
    db.close()
    return render_template('historial_movil.html', movil=movil, historial=historial)

@app.route("/moviles/ficha/<int:id>")
@login_requerido
@requerir_permiso('es_encargado_moviles')
def ficha_integral_movil(id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    # 1. DATOS PATRIMONIALES
    cursor.execute("SELECT * FROM moviles WHERE id = %s", (id,))
    movil = cursor.fetchone()
    
    if not movil:
        db.close()
        return "Móvil no encontrado", 404

    # Obtenemos el número de unidad (ej: 43/5) para buscarlo en nexo_personal
    nro_unidad_texto = movil['nro_unidad']

    # 2. FICHA TÉCNICA (Mantenimiento)
    cursor.execute("""
        SELECT * FROM historial_mantenimiento 
        WHERE id_movil = %s 
        ORDER BY fecha_reparacion DESC
    """, (id,))
    mantenimiento = cursor.fetchall()
    
    # 3. REGISTRO OPERATIVO (Siniestros/Salidas)
    # Usamos nexo_siniestros y nexo_personal que son tus tablas reales
    cursor.execute("""
        SELECT DISTINCT s.nro_part_ruba, s.fecha, s.tipo_siniestro, s.lugar
        FROM nexo_personal np
        JOIN nexo_siniestros s ON np.siniestro_id = s.id
        WHERE np.movil = %s
        ORDER BY s.fecha DESC
    """, (nro_unidad_texto,))
    salidas = cursor.fetchall()
    
    db.close()
    return render_template('ficha_integral_movil.html', 
                           movil=movil, 
                           mantenimiento=mantenimiento, 
                           salidas=salidas)

@app.route("/moviles/mantenimiento/nuevo/<int:id>")
@login_requerido
def nuevo_mantenimiento_especifico(id):
    db = get_db()
    cursor = db.cursor(dictionary=True)
    
    # Buscamos el móvil para pasar sus datos al formulario
    cursor.execute("SELECT id, nro_unidad, marca, modelo FROM moviles WHERE id = %s", (id,))
    movil = cursor.fetchone()
    
    db.close()
    
    if not movil:
        flash("Móvil no encontrado", "danger")
        return redirect(url_for('gestion_moviles'))
        
    # Aquí lo mandamos al mismo template que ya usas para registrar mantenimientos
    # Pero le pasamos el objeto 'movil_preseleccionado'
    return render_template('registrar_mantenimiento.html', movil_preseleccionado=movil)

import csv
from flask import Response

@app.route('/moviles/reporte-seguro')
@login_requerido
@requerir_permiso('es_encargado_moviles')
def reporte_seguro():
    db = get_db()
    cursor = db.cursor(dictionary=True)
    # Traemos solo los móviles activos para el seguro
    cursor.execute("SELECT nro_unidad, dominio, nro_chasis, nro_motor, marca, modelo, anio, aseguradora FROM moviles WHERE estado = 'ACTIVO'")
    moviles = cursor.fetchall()
    
    def generate():
        data = [['Unidad', 'Dominio', 'Chasis', 'Motor', 'Marca', 'Modelo', 'Año', 'Seguro']]
        yield ','.join(data[0]) + '\n'
        for m in moviles:
            row = [str(m['nro_unidad']), m['dominio'], m['nro_chasis'], m['nro_motor'], m['marca'], m['modelo'], str(m['anio']), m['aseguradora']]
            yield ','.join(row) + '\n'

    return Response(generate(), mimetype='text/csv', headers={"Content-disposition": "attachment; filename=listado_seguro_almafuerte.csv"})

@app.route('/moviles/reporte-excel')
@login_requerido
@requerir_permiso('es_encargado_moviles')
def reporte_excel():
    db = get_db()
    query = "SELECT nro_unidad, dominio, marca, modelo, nro_chasis, nro_motor, km_inicial, aseguradora, estado FROM moviles"
    df = pd.read_sql(query, db)
    db.close()

    df.columns = ['Unidad', 'Patente', 'Marca', 'Modelo', 'Chasis', 'Motor', 'KM inicial', 'Seguro', 'Estado actual']

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Parque Automotor')
        
        # Ajuste automático de columnas
        worksheet = writer.sheets['Parque Automotor']
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            worksheet.column_dimensions[column].width = max_length + 2

    output.seek(0)
    return send_file(output, download_name="SIAB_Moviles_Almafuerte.xlsx", as_attachment=True)

@app.route('/moviles/reporte-pdf')
@login_requerido
@requerir_permiso('es_encargado_moviles')
def reporte_pdf():
    try:
        db = get_db()
        cursor = db.cursor(dictionary=True)
        cursor.execute("SELECT nro_unidad, dominio, marca, modelo, nro_chasis, nro_motor, km_inicial FROM moviles WHERE estado != 'BAJA'")
        moviles = cursor.fetchall()
        db.close()

        output = BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4, topMargin=1*cm, bottomMargin=2.5*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
        elementos = []
        styles = getSampleStyleSheet()
        
        # --- ESTILOS PERSONALIZADOS ---
        estilo_entidad = styles['Normal']
        estilo_entidad.fontSize = 14
        estilo_entidad.fontName = 'Helvetica-Bold'
        estilo_entidad.alignment = 1 
        
        estilo_sistema = styles['Normal']
        estilo_sistema.fontSize = 9
        estilo_sistema.fontName = 'Helvetica' 
        estilo_sistema.alignment = 1
        estilo_sistema.leading = 12 

        # AGREGA ESTE BLOQUE QUE FALTABA:
        estilo_reporte = styles['Normal']
        estilo_reporte.fontSize = 12
        estilo_reporte.fontName = 'Helvetica-Bold'
        estilo_reporte.alignment = 1

        # --- 1. ENCABEZADO MEJORADO ---
        logo_path = os.path.join(base_dir, "static", "img", "Bomberos.png")        
        
        # --- ENCABEZADO ---
        col_textos = [
            Paragraph("SOCIEDAD BOMBEROS VOLUNTARIOS DE ALMAFUERTE", estilo_entidad),
            Spacer(1, 0.15*cm),
            Paragraph("<font name='Helvetica'>SIAB - Sistema Informático Automatizado de Bomberos</font>", estilo_sistema),
            Spacer(1, 0.15*cm),
            Paragraph("PARQUE AUTOMOTOR", estilo_reporte) # <--- Ahora sí encontrará la variable
        ]
        
        if os.path.exists(logo_path):
            img = Image(logo_path, 2.3*cm, 2.3*cm)
            header_data = [[img, col_textos]]
        else:
            header_data = [[ "", col_textos]]

        # Tabla de encabezado: Columna de logo fija, columna de texto centrada
        header_tab = Table(header_data, colWidths=[2.5*cm, 15.5*cm])
        header_tab.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'), # Centrar el contenido de la celda de texto
            ('LEFTPADDING', (1, 0), (1, 0), 0),
        ]))
        elementos.append(header_tab)
        
        elementos.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#a50000"), spaceAfter=2))
        
        # --- DATOS DE SESIÓN (USUARIO LOGEADO) ---
        # Se asume que guardas el nombre en session['user_nombre'] o session['nombre']
        usuario_logeado = session.get('user_nombre') or session.get('nombre') or "Usuario No Identificado"
        fecha_hora = datetime.now().strftime('%d/%m/%Y %H:%M')
        
        estilo_meta = styles['Normal']
        estilo_meta.fontSize = 8
        estilo_meta.textColor = colors.grey
        elementos.append(Paragraph(f"Generado por: {usuario_logeado} | Fecha y Hora: {fecha_hora}", estilo_meta))
        elementos.append(Spacer(1, 0.6 * cm))

        # --- 2. TABLA DE DATOS (SIN "NONE") ---
        headers = ['UNIDAD', 'PATENTE', 'MARCA / MODELO', 'CHASIS / MOTOR', 'KM INIC.']
        data = [headers]
        
        for m in moviles:
            patente = m['dominio'] if m['dominio'] else ""
            marca_mod = f"{m['marca'] if m['marca'] else ''} {m['modelo'] if m['modelo'] else ''}".strip()
            chasis = m['nro_chasis'] if m['nro_chasis'] else ""
            motor = m['nro_motor'] if m['nro_motor'] else ""
            
            data.append([
                m['nro_unidad'],
                patente,
                marca_mod,
                f"CH: {chasis}\nMOT: {motor}",
                m['km_inicial'] if m['km_inicial'] is not None else 0
            ])

        tabla = Table(data, repeatRows=1, colWidths=[2*cm, 3*cm, 5.5*cm, 5.5*cm, 2*cm])
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#a50000")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))
        elementos.append(tabla)

        # --- 3. PIE DE PÁGINA FIJO ---
        def footer(canvas, doc):
            canvas.saveState()
            
            # CAMBIA ESTAS DOS LÍNEAS:
            canvas.setStrokeColor(colors.HexColor("#a50000")) # Cambia colors.grey por el rojo
            canvas.setLineWidth(1.5)                         # Cambia 0.5 por 1.5 para que sea igual a la de arriba
            
            # Línea al final de la hoja (fija)
            canvas.line(1.5*cm, 1.5*cm, 19.5*cm, 1.5*cm)
            
            # El resto queda igual
            canvas.setFont('Helvetica', 8)
            canvas.setFillColor(colors.grey)
            canvas.drawCentredString(A4[0]/2, 1.1*cm, "Fin del Reporte Oficial - SIAB Almafuerte")
            canvas.restoreState()

        doc.build(elementos, onFirstPage=footer, onLaterPages=footer)
        
        output.seek(0)
        return send_file(output, download_name="Parque_Automotor_Almafuerte.pdf", as_attachment=True)

    except Exception as e:
        return f"Error en el sistema: {str(e)}"
          
# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)