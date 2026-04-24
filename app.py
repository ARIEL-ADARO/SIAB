"""
SIAB - Sistema Informático Automatizado de Bomberos
====================================================
App Flask principal - Etapa 2 v2
"""

from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import mysql.connector
from mysql.connector import Error
from datetime import datetime
import hashlib
import os
import hashlib
import hmac
from werkzeug.security import check_password_hash

app = Flask(__name__)
app.secret_key = "siab_bomberos_2026_secretkey"

# ============================================================
# CONFIGURACIÓN BASE DE DATOS
# ============================================================

DB_CONFIG = {
    "host":     "localhost",
    "port":     3306,
    "user":     "root",
    "password": "siab1234",
    "database": "siab"
}

def get_db():
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        return conn
    except Error as e:
        print(f"Error de conexión: {e}")
        return None

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
            SELECT u.*, l.nombre, l.apellido, l.grado
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
# INICIO
# ============================================================

@app.route("/inicio")
@login_requerido
def inicio():
    conn = get_db()
    stats = {}
    borradores = []
    if conn:
        cur = conn.cursor(dictionary=True)

        cur.execute("SELECT COUNT(*) as total FROM legajos WHERE situacion = 'ACTIVO'")
        stats["bomberos_activos"] = cur.fetchone()["total"]

        cur.execute("""SELECT COUNT(*) as total FROM eventos
                       WHERE MONTH(fecha) = MONTH(CURDATE())
                       AND YEAR(fecha) = YEAR(CURDATE())
                       AND estado = 'FINALIZADO'""")
        stats["eventos_mes"] = cur.fetchone()["total"]

        cur.execute("""SELECT COUNT(*) as total FROM asistencia a
                       JOIN eventos e ON a.evento_id = e.id
                       WHERE a.estado = 'PRESENTE'
                       AND MONTH(e.fecha) = MONTH(CURDATE())
                       AND e.estado = 'FINALIZADO'""")
        stats["asistencias_mes"] = cur.fetchone()["total"]

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
            LIMIT 5        
        """, (depto_id,))
    else:
        # Si no se eligió departamento, trae a TODOS los activos
        cur.execute("""
            SELECT legajo, apellido, nombre, grado, cargo
            FROM legajos
            WHERE situacion = 'ACTIVO'
            ORDER BY apellido, nombre
            LIMIT 5        
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
            LIMIT 5
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
    legajo_usuario = session.get('legajo') # <--- PASO 1: Obtener quién es el usuario
    
    conn = get_db()
    eventos = []
    
    if conn:
        cur = conn.cursor(dictionary=True)
        
        # PASO 2: La consulta ahora tiene el "mi_estado"
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
        
        if not ver_anulados:
            query += " AND e.estado != 'ANULADO'"
            
        query += " ORDER BY e.fecha DESC, e.id DESC LIMIT 50"
        
        cur.execute(query, (legajo_usuario,)) # <--- PASO 3: Pasamos el legajo aquí
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
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT d.*, COUNT(bd.id) as miembros
            FROM departamentos d
            LEFT JOIN bombero_departamento bd ON d.id = bd.departamento_id AND bd.activo = 1
            WHERE d.activo = 1
            GROUP BY d.id
            ORDER BY d.nombre
        """)
        lista = cur.fetchall()
        conn.close()
    return render_template("departamentos.html", departamentos=lista)


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
            LIMIT 5        
        """)
        lista = cur.fetchall()
        conn.close()
    return render_template("bomberos.html", bomberos=lista)


# ============================================================
# CONFIGURACIÓN DE PUNTOS
# ============================================================

@app.route("/config/puntos")
@login_requerido
@rol_requerido("ADMIN")
def config_puntos():
    conn = get_db()
    registros = []
    if conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM config_puntos ORDER BY anio DESC")
        registros = cur.fetchall()
        conn.close()
    return render_template("config_puntos.html", registros=registros)


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
    
    # Perfil por defecto (esto es lo que verá el ADMIN si no está en la tabla legajos)
    datos = {
        'legajo': legajo,
        'apellido': 'ADMINISTRADOR',
        'nombre': 'SISTEMA',
        'grado': 'SOPORTE',
        'cargo': 'ADMIN',
        'dni': '0',
        'situacion': 'ACTIVO',
        'email': 'admin@sistema.com',
        'nro_cel': 'N/A',
        'asistencias_anio': 0,
        'promedio_general': 0.0,
        'pilar_vocacion': 5.0,
        'pilar_capacidad': 5.0,
        'pilar_asistencia': 5.0,
        'pilar_cualidades': 5.0,
        'puntaje_final': 5.0,
        'calif_letra': "EXCELENTE"
    }
    historial = []

    if not conn:
        return datos, [] # Devolvemos el perfil genérico en lugar de None

    try:
        cur = conn.cursor(dictionary=True)
        from datetime import datetime
        mes_actual = datetime.now().strftime('%Y-%m')

        # 1. Datos básicos
        cur.execute("SELECT * FROM legajos WHERE legajo = %s", (legajo,))
        perfil_base = cur.fetchone()
        if perfil_base: datos.update(perfil_base)

        # 2. CÁLCULO DE INDICADORES DE IMPACTO
        # -----------------------------------------------------------
        
        # CLASES OBLIGATORIAS (Pilar Capacitación)
        # Contamos asistencias presentes en el mes. Cada una vale 2.5 pts.
        cur.execute("""
            SELECT COUNT(*) as total 
            FROM asistencia 
            WHERE legajo = %s AND estado = 'PRESENTE' 
            AND fecha_registro LIKE %s
        """, (legajo, f"{mes_actual}%"))
        clases_asistidas = cur.fetchone()['total'] or 0
        
        datos['puntos_capacitacion'] = min(clases_asistidas * 2.5, 5.0)
        datos['clases_conteo'] = clases_asistidas

        # HORAS DE ACTIVIDAD (Gestión/Dedicación) - ACUMULADO TOTAL
        cur.execute("""
            SELECT SUM(horas) as total_horas 
            FROM actividades 
            WHERE legajo = %s AND actividad NOT IN ('PRÁCTICA', 'CAPACITACIÓN')
            AND anulada = 0
        """, (legajo,)) # Quitamos el parámetro del mes
        horas_gestion = cur.fetchone()['total_horas'] or 0

        datos['horas_actividad_reales'] = round(horas_gestion, 1)
        # Aplicamos el tope: 10hs = 5pts.
        datos['puntos_actividad'] = min((horas_gestion / 10) * 5, 5.0)

        # EMERGENCIAS (Pilar Operativo - Conectado a Registro de Salidas)
        # -----------------------------------------------------------
        # Sumamos los puntos de las intervenciones calificadas del mes actual
        cur.execute("""
            SELECT 
                COUNT(p.id) as total_salidas,
                SUM(p.puntos_operativos) as puntos_totales
            FROM nexo_personal p
            JOIN nexo_siniestros s ON p.siniestro_id = s.id
            WHERE p.legajo = %s 
            AND s.estado = 'CALIFICADO'
            AND s.fecha LIKE %s
        """, (legajo, f"{mes_actual}%"))
        
        resumen_operativo = cur.fetchone()
        
        # Guardamos los valores en el diccionario 'datos' para el radar
        total_salidas = resumen_operativo['total_salidas'] or 0
        puntos_ope = float(resumen_operativo['puntos_totales'] or 0.0)

        datos['total_salidas'] = total_salidas
        # El radar suele tener un tope (ej: 5.0) para no deformarse
        datos['puntos_operativo'] = min(puntos_ope, 5.0) 
        # Guardamos el real por si queremos mostrarlo en texto
        datos['puntos_operativo_reales'] = puntos_ope

        # FIRMAS PENDIENTES (Solo borradores del bombero)
        cur.execute("""
            SELECT COUNT(*) as total FROM actividades 
            WHERE legajo = %s AND firma_bombero_fecha IS NULL AND anulada = 0
        """, (legajo,))
        datos['pendientes_firma_bombero'] = cur.fetchone()['total']

        # 3. Historial para los botones de "Ver más"
        cur.execute("""
            SELECT *, fecha_inicio AS fecha, actividad AS tipo 
            FROM actividades WHERE legajo = %s AND anulada = 0
            ORDER BY fecha_inicio DESC, hora_inicio DESC LIMIT 20
        """, (legajo,))
        historial_raw = cur.fetchall()
        
        historial = []
        for h in historial_raw:
            if h.get('fecha_inicio'):
                try:
                    h['fecha_display'] = h['fecha_inicio'].strftime('%d/%m/%Y')
                except:
                    h['fecha_display'] = str(h['fecha_inicio'])
            historial.append(h)

        # --- AQUÍ EMPIEZA LA SUGERENCIA: INTEGRACIÓN DE SALIDAS ---
        cur.execute("""
            SELECT 
                s.fecha as fecha, 
                s.tipo_siniestro as tipo, 
                CONCAT('Móvil: ', p.movil, ' - Rol: ', p.rol) as descripcion, 
                p.puntos_operativos as horas,
                s.fecha as fecha_inicio
            FROM nexo_personal p
            JOIN nexo_siniestros s ON p.siniestro_id = s.id
            WHERE p.legajo = %s AND s.estado = 'CALIFICADO'
            ORDER BY s.fecha DESC LIMIT 10
        """, (legajo,))
        
        intervenciones = cur.fetchall()
        for i in intervenciones:
            if i.get('fecha'):
                try:
                    i['fecha_display'] = i['fecha'].strftime('%d/%m/%Y')
                except:
                    i['fecha_display'] = str(i['fecha'])
            historial.append(i)

        # Volvemos a ordenar la lista completa por fecha para que no queden las salidas todas al final
        historial.sort(key=lambda x: x.get('fecha_inicio') if x.get('fecha_inicio') else x.get('fecha'), reverse=True)
        # --- FIN DE LA SUGERENCIA ---

    except Exception as e:
        print(f"Error en el perfil: {e}")
        return datos, [] 
    
    finally:
        if 'cur' in locals():
            cur.close()

    return datos, historial

@app.route("/ver-perfil/<int:legajo_id>")
@rol_requerido('ADMIN', 'JEFATURA')
def ver_perfil_ajeno(legajo_id):
    datos, historial = obtener_datos_completos_perfil(legajo_id)
    if not datos:
        flash("No se encontró el legajo.", "warning")
        return redirect(url_for('inicio'))
    return render_template("mi_perfil.html", datos=datos, historial=historial)

@app.route("/mi-perfil")
@login_requerido
def mi_perfil():
    legajo = session.get("legajo")
    datos, historial = obtener_datos_completos_perfil(legajo)
    if not datos:
        flash("Error al cargar tu perfil.", "danger")
        return redirect(url_for('inicio'))
    return render_template("mi_perfil.html", datos=datos, historial=historial)

@app.route("/mesa-calificadora")
@login_requerido
@rol_requerido('ADMIN', 'JEFATURA')
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

@app.route("/mesa-calificadora/cerrar-ciclo", methods=["POST"])
@login_requerido
def cerrar_ciclo_anual():
    # Verificación de seguridad: Solo Jefatura o Admin
    if session.get('grado') != 'JEFATURA' and session.get('rol') != 'ADMIN':
        return "Acceso denegado", 403
    
    conn = get_db()
    cur = conn.cursor()
    try:
        # Pasamos la nota que sacaron este año a la columna de año anterior
        cur.execute("UPDATE calificaciones_cualidades SET anio_anterior_puntos = nota_cualidades")
        # Opcional: Podrías resetear la nota actual a 0 para empezar el nuevo año
        # cur.execute("UPDATE calificaciones_cualidades SET nota_cualidades = 0")
        conn.commit()
        # Aquí podrías agregar un mensaje de éxito con flash
    except Exception as e:
        print(f"Error al cerrar ciclo: {e}")
        conn.rollback()
    finally:
        conn.close()
        
    return redirect(url_for('mesa_calificadora'))

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
    
    # AGREGÁ 'id' AL PRINCIPIO DEL SELECT
    cur.execute("""
        SELECT id, nro_part_ruba, fecha, hora_salida, tipo_siniestro, lugar, estado 
        FROM nexo_siniestros 
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

@app.route("/mis-capacitaciones")
@login_requerido
def mis_capacitaciones():
    legajo = session.get('legajo')
    db = get_db()
    if not db:
        return "Error de conexión a la base de datos", 500
        
    try:
        cur = db.cursor(dictionary=True)
        cur.execute("""
            SELECT *, fecha_inicio as fecha 
            FROM actividades 
            WHERE legajo = %s AND actividad IN ('PRÁCTICA', 'CAPACITACIÓN') AND anulada = 0
            ORDER BY fecha_inicio DESC
        """, (legajo,))
        registros = cur.fetchall()
        
        # 1. PRIMERO HACEMOS LA SUMA
        total_horas = sum(float(r['horas'] or 0) for r in registros)

        # 2. DESPUÉS HACEMOS EL RETURN (enviando todas las variables)
        return render_template("detalle_actividades.html", 
                               titulo="Mis Capacitaciones", 
                               registros=registros, 
                               total_horas=total_horas)
    
    finally:
        cur.close()
        db.close()

@app.route("/mis-actividades-gestion")
@login_requerido
def mis_actividades_gestion():
    legajo = session.get('legajo')
    db = get_db()
    if not db:
        return "Error de conexión a la base de datos", 500
        
    try:
        cur = db.cursor(dictionary=True)
        cur.execute("""
            SELECT *, fecha_inicio as fecha 
            FROM actividades 
            WHERE legajo = %s AND actividad NOT IN ('PRÁCTICA', 'CAPACITACIÓN') AND anulada = 0
            ORDER BY fecha_inicio DESC
        """, (legajo,))
        registros = cur.fetchall()

        # 1. PRIMERO HACEMOS LA SUMA
        total_horas = sum(float(r['horas'] or 0) for r in registros)

        # 2. DESPUÉS HACEMOS EL RETURN
        return render_template("detalle_actividades.html", 
                               titulo="Gestión y Dedicación", 
                               registros=registros, 
                               total_horas=total_horas)

    finally:
        cur.close()
        db.close()

from datetime import datetime

@app.route('/planilla-nexo/nueva', methods=['GET', 'POST'])
@login_requerido
def nueva_planilla_nexo():
    db = get_db()
    cur = db.cursor(dictionary=True)

    if request.method == 'POST':
        try:
            nro_parte = request.form.get('nro_parte')
            tipo = request.form.get('tipo_siniestro')
            lugar = request.form.get('lugar')
            fecha = request.form.get('fecha') 
            hora = request.form.get('hora_salida')

            sql_siniestro = """
                INSERT INTO nexo_siniestros (nro_part_ruba, fecha, hora_salida, tipo_siniestro, lugar, estado)
                VALUES (%s, %s, %s, %s, %s, 'BORRADOR')
            """
            cur.execute(sql_siniestro, (nro_parte, fecha, hora, tipo, lugar))
            siniestro_id = cur.lastrowid

            bomberos_seleccionados = request.form.getlist('bomberos_seleccionados')
            for legajo in bomberos_seleccionados:
                movil = request.form.get(f'movil_{legajo}')
                rol = request.form.get(f'rol_{legajo}')
                cur.execute("INSERT INTO nexo_personal (siniestro_id, legajo, movil, rol) VALUES (%s, %s, %s, %s)", 
                            (siniestro_id, legajo, movil, rol))

            db.commit()
            flash(f"Registro #{siniestro_id} guardado.", "success")
            return redirect(url_for('listado_siniestros'))
        except Exception as e:
            db.rollback()
            print(f"Error al guardar: {e}")
            flash("Error al guardar.", "danger")
            return redirect(url_for('nueva_planilla_nexo'))

    # --- MÉTODO GET: CARGA DE FORMULARIO ---
    # Usamos LIKE para que si hay espacios locos o caracteres raros, los encuentre igual
    cur.execute("""
        SELECT legajo, nombre, apellido, situacion 
        FROM legajos 
        WHERE situacion LIKE '%ACTIVO%' 
           OR situacion LIKE '%RESERVA%'
        ORDER BY apellido ASC
    """)
    res_bomberos = cur.fetchall()
    
    # DEBUG: Esto te dirá en la consola negra cuántos cargó realmente
    print(f"DEBUG SIAB: Enviando {len(res_bomberos)} bomberos al HTML")

    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    hora_hoy = datetime.now().strftime('%H:%M')
    
    cur.close()
    db.close()
    
    # IMPORTANTE: El nombre a la IZQUIERDA del igual debe ser 'bomberos'
    return render_template('nexo_form.html', 
                           bomberos=res_bomberos, 
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
        SELECT p.*, b.nombre, b.apellido 
        FROM nexo_personal p
        JOIN bomberos b ON p.legajo = b.legajo
        WHERE p.siniestro_id = %s
    """, (id,))
    personal = cur.fetchall()
    
    return render_template('nexo_print.html', siniestro=siniestro, personal=personal)

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

@app.route('/registro-salidas/calificar/<int:id>')
@login_requerido
def calificar_salida(id):
    # Verificamos si el rol tiene permiso
    rol_usuario = session.get('rol')
    permisos_autorizados = ['ADMIN', 'ENCARGADO', 'SUPERVISOR']

    if rol_usuario not in permisos_autorizados:
        flash("No tenés permisos para calificar salidas.", "danger")
        return redirect(url_for('listado_siniestros'))
    
    db = get_db()
    cur = db.cursor(dictionary=True)

    if request.method == 'POST':
        # 1. Recibimos los puntos de cada bombero
        # El formulario enviará un diccionario con legajo: puntaje
        for legajo in request.form.getlist('legajos'):
            puntos = request.form.get(f'puntos_{legajo}')
            
            # Actualizamos la tabla nexo_personal
            cur.execute("""
                UPDATE nexo_personal 
                SET puntos_operativos = %s 
                WHERE siniestro_id = %s AND legajo = %s
            """, (puntos, id, legajo))

        # 2. Marcamos el siniestro como CALIFICADO
        cur.execute("UPDATE nexo_siniestros SET estado = 'CALIFICADO' WHERE id = %s", (id,))
        
        db.commit()
        flash("Puntajes asignados correctamente. El radar de los bomberos ha sido actualizado.", "success")
        return redirect(url_for('listado_siniestros'))

    # GET: Datos para la pantalla
    cur.execute("SELECT * FROM nexo_siniestros WHERE id = %s", (id,))
    siniestro = cur.fetchone()

    cur.execute("""
        SELECT p.*, b.apellido, b.nombre 
        FROM nexo_personal p
        JOIN bomberos b ON p.legajo = b.legajo
        WHERE p.siniestro_id = %s
    """, (id,))
    personal = cur.fetchall()

    return render_template('nexo_calificar.html', siniestro=siniestro, personal=personal)        

@app.route('/admin/backup')
def ejecutar_backup():
    if session.get('rol') != 'ADMIN':
        return redirect(url_for('inicio'))

    try:
        db_user = DB_CONFIG['user']
        db_pass = DB_CONFIG['password']
        db_name = DB_CONFIG['database']
        
        folder = r"C:\SIAB\backups"
        if not os.path.exists(folder): os.makedirs(folder)

        fecha = datetime.now().strftime("%Y-%m-%d_%H-%M")
        filename = f"backup_{db_name}_{fecha}.sql"
        filepath = os.path.join(folder, filename)

        # --- BUSCADOR DEL EJECUTABLE ---
        posibles_rutas = [
            r"C:\xampp\mysql\bin\mysqldump.exe",
            r"C:\Program Files\MySQL\MySQL Server 8.0\bin\mysqldump.exe",
            r"C:\Program Files\MySQL\MySQL Server 8.1\bin\mysqldump.exe",
            "mysqldump" # Si está en el PATH
        ]
        
        dump_exe = None
        for ruta in posibles_rutas:
            if ruta == "mysqldump" or os.path.exists(ruta):
                dump_exe = ruta
                break

        if not dump_exe:
            flash("No se encontró mysqldump.exe. Verificá la instalación de MySQL/XAMPP.", "danger")
            return redirect(url_for('inicio'))
        # -------------------------------

        comando = [dump_exe, f"--user={db_user}", f"--password={db_pass}", db_name]

        with open(filepath, "w") as out_file:
            resultado = subprocess.run(comando, stdout=out_file, stderr=subprocess.PIPE, text=True)

        if resultado.returncode != 0:
            if os.path.exists(filepath): os.remove(filepath)
            flash(f"Error de MySQL: {resultado.stderr}", "danger")
        else:
            flash(f"¡Respaldo exitoso! Guardado en {folder}", "success")

    except Exception as e:
        flash(f"Error crítico: {str(e)}", "danger")
    
    return redirect(url_for('inicio'))

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)