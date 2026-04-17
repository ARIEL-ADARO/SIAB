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

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

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

        if usuario and usuario["password_hash"] == hash_password(password):
            apellido = usuario.get("apellido") or ""
            nombre   = usuario.get("nombre") or ""
            nombre_completo = f"{apellido} {nombre}".strip() or username

            session["usuario_id"] = usuario["id"]
            session["username"]   = usuario["username"]
            session["rol"]        = usuario["rol"]
            session["legajo"]     = usuario["legajo"]
            session["nombre"]     = nombre_completo

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
                       AND estado = 'CONFIRMADO'""")
        stats["eventos_mes"] = cur.fetchone()["total"]

        cur.execute("""SELECT COUNT(*) as total FROM asistencia a
                       JOIN eventos e ON a.evento_id = e.id
                       WHERE a.estado = 'PRESENTE'
                       AND MONTH(e.fecha) = MONTH(CURDATE())
                       AND e.estado = 'CONFIRMADO'""")
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

@app.route("/asistencia")
@login_requerido
def asistencia():
    conn = get_db()
    conceptos     = []
    departamentos = []
    config_puntos = None
    if conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT id, concepto FROM conceptos WHERE activo = 1 ORDER BY concepto")
        conceptos = cur.fetchall()
        cur.execute("SELECT id, nombre FROM departamentos WHERE activo = 1 ORDER BY nombre")
        departamentos = cur.fetchall()
        cur.execute("SELECT puntos_por_asistencia FROM config_puntos WHERE anio = YEAR(CURDATE()) LIMIT 1")
        config_puntos = cur.fetchone()
        conn.close()
    return render_template("asistencia.html",
                           conceptos=conceptos,
                           departamentos=departamentos,
                           config_puntos=config_puntos)

@app.route("/asistencia/guardar", methods=["POST"])
@login_requerido
def guardar_asistencia():
    data = request.get_json()
    # Capturamos el departamento del payload (si es "todos" lo guardamos como None)
    depto_id = data.get("departamento_id")
    if depto_id == "todos" or depto_id == "":
        depto_id = None

    tipo        = data.get("tipo")
    descripcion = data.get("descripcion", "")
    fecha       = data.get("fecha")
    hora_inicio = data.get("hora_inicio") or None
    hora_fin    = data.get("hora_fin") or None
    concepto_id = data.get("concepto_id") or None
    asistencias = data.get("asistencias", [])
    confirmar   = data.get("confirmar", False)
    evento_id   = data.get("evento_id")

    if not tipo or not fecha or not asistencias:
        return jsonify({"ok": False, "error": "Faltan datos obligatorios."})

    estado = "CONFIRMADO" if confirmar else "BORRADOR"

    conn = get_db()
    if not conn:
        return jsonify({"ok": False, "error": "Error de conexión."})

    try:
        cur = conn.cursor()

        # --- A. GUARDAR O ACTUALIZAR EVENTO ---
        if evento_id:
            cur.execute("""
                UPDATE eventos SET tipo=%s, descripcion=%s, fecha=%s,
                hora_inicio=%s, hora_fin=%s, concepto_id=%s, estado=%s, departamento_id=%s
                WHERE id=%s
            """, (tipo, descripcion, fecha, hora_inicio, hora_fin,
                  concepto_id, estado, depto_id, evento_id))
            cur.execute("DELETE FROM asistencia WHERE evento_id = %s", (evento_id,))
            # TAMBIÉN LIMPIAMOS TEMAS ANTERIORES
            cur.execute("DELETE FROM evento_temas WHERE evento_id = %s", (evento_id,))
        else:
            cur.execute("""
                INSERT INTO eventos (tipo, descripcion, fecha, hora_inicio, hora_fin,
                                     concepto_id, estado, creado_por, departamento_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (tipo, descripcion, fecha, hora_inicio, hora_fin,
                  concepto_id, estado, session["usuario_id"], depto_id))
            evento_id = cur.lastrowid

        # --- B. GUARDAR TEMAS/POSTAS (NUEVO) ---
        temas = data.get("temas", []) # Capturamos la lista de temas del JSON
        for i, t in enumerate(temas):
            if t.get("nombre") and t.get("nombre").strip():
                cur.execute("""
                    INSERT INTO evento_temas (evento_id, nombre, calificador_legajo, orden)
                    VALUES (%s, %s, %s, %s)
                """, (evento_id, t["nombre"], t.get("calificador_legajo") or None, i + 1))

        # --- C. GUARDAR ASISTENCIAS ---
        for a in asistencias:
            cur.execute("""
                INSERT INTO asistencia (evento_id, legajo, estado, observacion, 
                                        calificacion, registrado_por)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE 
                    estado = VALUES(estado),
                    observacion = VALUES(observacion),
                    calificacion = VALUES(calificacion),
                    registrado_por = VALUES(registrado_por)
            """, (evento_id, a["legajo"], a["estado"], 
                  a.get("observacion", ""), 
                  a.get("calificacion") or None, 
                  session["usuario_id"]))
                  
        conn.commit()
        return jsonify({"ok": True, "evento_id": evento_id, "estado": estado})

    except Error as e:
        conn.rollback()
        conn.close()
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

    # 3. Traer asistencias previas (Bomberos marcados)
    cur.execute("""
        SELECT a.legajo, a.estado, a.observacion, a.calificacion,
               l.apellido, l.nombre, l.grado
        FROM asistencia a
        JOIN legajos l ON a.legajo = l.legajo
        WHERE a.evento_id = %s
    """, (evento_id,))
    asistencias = cur.fetchall()

    dict_asistencias = {str(a['legajo']): {
        'estado': a['estado'], 
        'observacion': a['observacion'],
        'calificacion': a['calificacion'],
        'apellido': a['apellido'],
        'nombre': a['nombre'],
        'grado': a['grado']
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
    conn = get_db()
    eventos = []
    if conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT 
                e.id, e.tipo, e.descripcion, e.fecha, e.hora_inicio, e.estado, 
                c.concepto, 
                d.nombre AS nombre_departamento,
                COUNT(a.id) as total,
                SUM(CASE WHEN a.estado = 'PRESENTE' THEN 1 ELSE 0 END) as presentes,
                SUM(CASE WHEN a.estado = 'AUSENTE' THEN 1 ELSE 0 END) as ausentes,
                SUM(CASE WHEN a.estado = 'JUSTIFICADO' THEN 1 ELSE 0 END) as justificados
            FROM eventos e
            LEFT JOIN conceptos c ON e.concepto_id = c.id
            LEFT JOIN departamentos d ON e.departamento_id = d.id
            LEFT JOIN asistencia a ON e.id = a.evento_id
            GROUP BY e.id  -- Agrupamos solo por ID de evento para mayor seguridad
            ORDER BY e.fecha DESC, e.id DESC
            LIMIT 100
        """)
        eventos = cur.fetchall()
        conn.close()
    return render_template("historial_asistencia.html", eventos=eventos)

@app.route("/asistencia/detalle/<int:evento_id>")
@login_requerido
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


@app.route("/departamentos/<int:depto_id>/miembros")
@login_requerido
def miembros_departamento(depto_id):
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
    if conn:
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT c.*, COUNT(cp.id) as participantes
            FROM cursos c
            LEFT JOIN curso_participantes cp ON c.id = cp.curso_id
            GROUP BY c.id
            ORDER BY c.fecha_inicio DESC
        """)
        lista = cur.fetchall()
        conn.close()
    return render_template("cursos.html", cursos=lista)


@app.route("/cursos/nuevo", methods=["GET", "POST"])
@login_requerido
def nuevo_curso():
    if request.method == "POST":
        nombre      = request.form.get("nombre")
        institucion = request.form.get("institucion", "")
        fecha_ini   = request.form.get("fecha_inicio") or None
        fecha_fin   = request.form.get("fecha_fin") or None
        horas       = request.form.get("horas") or None
        descripcion = request.form.get("descripcion", "")
        legajos     = request.form.getlist("participantes")

        conn = get_db()
        if conn:
            try:
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO cursos (nombre, institucion, fecha_inicio, fecha_fin,
                                        horas, descripcion, creado_por)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (nombre, institucion, fecha_ini, fecha_fin,
                      horas, descripcion, session["usuario_id"]))
                curso_id = cur.lastrowid
                for legajo in legajos:
                    cur.execute("""
                        INSERT INTO curso_participantes (curso_id, legajo)
                        VALUES (%s, %s)
                    """, (curso_id, legajo))
                conn.commit()
                flash(f"Curso '{nombre}' registrado con {len(legajos)} participantes.", "success")
                return redirect(url_for("cursos"))
            except Error as e:
                conn.rollback()
                flash(f"Error al guardar: {e}", "danger")
            finally:
                conn.close()

    conn = get_db()
    bomberos      = []
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
    """Muestra la cuadrícula de notas para las postas de una capacitación"""
    conn = get_db()
    if not conn: return redirect(url_for("inicio"))
    
    cur = conn.cursor(dictionary=True)
    
    # Traer el evento
    cur.execute("SELECT * FROM eventos WHERE id = %s", (evento_id,))
    evento = cur.fetchone()
    
    # Traer los temas/postas
    cur.execute("SELECT * FROM evento_temas WHERE evento_id = %s ORDER BY orden", (evento_id,))
    temas = cur.fetchall()
    
    # Traer solo a los que están PRESENTES
    cur.execute("""
        SELECT a.legajo, l.apellido, l.nombre, l.grado
        FROM asistencia a
        JOIN legajos l ON a.legajo = l.legajo
        WHERE a.evento_id = %s AND a.estado = 'PRESENTE'
        ORDER BY l.apellido, l.nombre
    """, (evento_id,))
    presentes = cur.fetchall()
    
    conn.close()
    return render_template("cargar_notas.html", evento=evento, temas=temas, presentes=presentes)

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
# MAIN
# ============================================================

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)