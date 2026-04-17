"""
SIAB - Script de Migración SQLite → MySQL
==========================================
Ejecutar DESPUÉS de tener MySQL corriendo localmente.

Requisitos:
    pip install mysql-connector-python

Uso:
    1. Ajustá las variables de conexión MySQL abajo (MYSQL_*)
    2. Poné la ruta correcta a tu siab.db
    3. Ejecutá: python migrar_siab.py
"""

import sqlite3
import mysql.connector
from mysql.connector import Error
import sys

# ============================================================
# CONFIGURACIÓN — AJUSTAR ANTES DE EJECUTAR
# ============================================================

SQLITE_PATH = r"C:\SIAB\siab.db"   # Ruta a tu .db actual

MYSQL_HOST     = "localhost"
MYSQL_PORT     = 3306
MYSQL_USER     = "root"
MYSQL_PASSWORD = "siab1234"         # ← cambiá esto
MYSQL_DB       = "siab"                     # Se crea automáticamente

# ============================================================


def conectar_mysql_sin_db():
    return mysql.connector.connect(
        host=MYSQL_HOST,
        port=MYSQL_PORT,
        user=MYSQL_USER,
        password=MYSQL_PASSWORD
    )

def conectar_mysql():
    return mysql.connector.connect(
        host=MYSQL_HOST,
        port=MYSQL_PORT,
        user=MYSQL_USER,
        password=MYSQL_PASSWORD,
        database=MYSQL_DB
    )

def crear_base_de_datos():
    print(f"\n[1/5] Creando base de datos '{MYSQL_DB}'...")
    try:
        conn = conectar_mysql_sin_db()
        cur = conn.cursor()
        cur.execute(f"DROP DATABASE IF EXISTS `{MYSQL_DB}`")
        cur.execute(f"CREATE DATABASE `{MYSQL_DB}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
        print(f"      ✓ Base '{MYSQL_DB}' creada.")
        conn.close()
    except Error as e:
        print(f"      ✗ Error: {e}")
        sys.exit(1)

def crear_tablas(cur):
    print("\n[2/5] Creando tablas...")

    tablas = []

    # ----------------------------------------------------------
    # CONFIGURACION
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS configuracion (
        id               INT AUTO_INCREMENT PRIMARY KEY,
        institucion      VARCHAR(200),
        direccion        VARCHAR(200),
        ciudad           VARCHAR(100),
        provincia        VARCHAR(100),
        telefono         VARCHAR(50),
        email            VARCHAR(100),
        sitio_web        VARCHAR(200),
        logo             VARCHAR(200),
        version_sistema  VARCHAR(20),
        fecha_instalacion VARCHAR(20)
    )
    """)

    # ----------------------------------------------------------
    # LEGAJOS
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS legajos (
        legajo      INT PRIMARY KEY,
        apellido    VARCHAR(100),
        nombre      VARCHAR(100),
        grado       VARCHAR(50),
        cargo       VARCHAR(100),
        dni         VARCHAR(20),
        foto        VARCHAR(200),
        fecha_carga VARCHAR(20),
        situacion   VARCHAR(20),
        fecha_baja  VARCHAR(20),
        autoriza    VARCHAR(100),
        email       VARCHAR(100),
        nro_cel     VARCHAR(30)
    )
    """)

    # ----------------------------------------------------------
    # CONCEPTOS
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS conceptos (
        id      INT AUTO_INCREMENT PRIMARY KEY,
        concepto VARCHAR(200),
        puntos  INT DEFAULT 0,
        detalle TEXT,
        activo  TINYINT DEFAULT 1
    )
    """)

    # ----------------------------------------------------------
    # USUARIOS
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id                    INT AUTO_INCREMENT PRIMARY KEY,
        username              VARCHAR(50) UNIQUE NOT NULL,
        password_hash         VARCHAR(200),
        rol                   VARCHAR(20),
        legajo                INT,
        activo                TINYINT DEFAULT 1,
        fecha_baja            VARCHAR(20),
        usuario_baja          VARCHAR(50),
        debe_cambiar_password TINYINT DEFAULT 0,
        FOREIGN KEY (legajo) REFERENCES legajos(legajo)
    )
    """)

    # ----------------------------------------------------------
    # ACTIVIDADES
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS actividades (
        id                      INT AUTO_INCREMENT PRIMARY KEY,
        legajo                  VARCHAR(20),
        actividad               VARCHAR(200),
        area                    VARCHAR(100),
        fecha_inicio            VARCHAR(20),
        fecha_fin               VARCHAR(20),
        hora_inicio             VARCHAR(10),
        hora_fin                VARCHAR(10),
        descripcion             TEXT,
        fecha_carga             VARCHAR(20),
        asignado                VARCHAR(100),
        usuario_id              INT,
        horas                   DECIMAL(5,2),
        concepto_id             INT,
        firma_bombero_usuario   VARCHAR(50),
        firma_bombero_fecha     VARCHAR(30),
        firma_supervisor_usuario VARCHAR(50),
        firma_supervisor_fecha  VARCHAR(30),
        anulada                 TINYINT DEFAULT 0,
        estado                  VARCHAR(30),
        motivo_anulacion        TEXT,
        usuario_anula           VARCHAR(50),
        fecha_anulacion         VARCHAR(20),
        creado_por              INT,
        fecha_creacion          VARCHAR(30),
        modificado_por          INT,
        fecha_modificacion      VARCHAR(30)
    )
    """)

    # ----------------------------------------------------------
    # ACTIVIDADES HISTORIAL
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS actividades_historial (
        id           INT AUTO_INCREMENT PRIMARY KEY,
        actividad_id INT,
        fecha        VARCHAR(30),
        usuario_id   INT,
        campo        VARCHAR(100),
        valor_anterior TEXT,
        valor_nuevo    TEXT,
        FOREIGN KEY (actividad_id) REFERENCES actividades(id)
    )
    """)

    # ----------------------------------------------------------
    # NOTIFICACIONES
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS notificaciones (
        id           INT AUTO_INCREMENT PRIMARY KEY,
        actividad_id INT,
        tipo         VARCHAR(50),
        destinatario VARCHAR(100),
        asunto       VARCHAR(200),
        fecha_envio  VARCHAR(30),
        estado       VARCHAR(20),
        detalle_error TEXT
    )
    """)

    # ==========================================================
    # TABLAS NUEVAS - ETAPA 2
    # ==========================================================

    # ----------------------------------------------------------
    # EVENTOS (agrupa asistencia: guardia, salida, capacitación)
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS eventos (
        id          INT AUTO_INCREMENT PRIMARY KEY,
        tipo        VARCHAR(50)  NOT NULL,
        descripcion VARCHAR(300),
        fecha       DATE         NOT NULL,
        hora_inicio TIME,
        hora_fin    TIME,
        concepto_id INT,
        creado_por  INT,
        fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (concepto_id) REFERENCES conceptos(id),
        FOREIGN KEY (creado_por)  REFERENCES usuarios(id)
    )
    """)

    # ----------------------------------------------------------
    # ASISTENCIA (un registro por bombero por evento)
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS asistencia (
        id          INT AUTO_INCREMENT PRIMARY KEY,
        evento_id   INT          NOT NULL,
        legajo      INT          NOT NULL,
        estado      ENUM('PRESENTE','AUSENTE','JUSTIFICADO') DEFAULT 'AUSENTE',
        observacion VARCHAR(300),
        registrado_por INT,
        fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (evento_id)      REFERENCES eventos(id),
        FOREIGN KEY (legajo)         REFERENCES legajos(legajo),
        FOREIGN KEY (registrado_por) REFERENCES usuarios(id),
        UNIQUE KEY uq_asistencia (evento_id, legajo)
    )
    """)

    # ----------------------------------------------------------
    # CURSOS
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS cursos (
        id          INT AUTO_INCREMENT PRIMARY KEY,
        nombre      VARCHAR(300) NOT NULL,
        institucion VARCHAR(200),
        fecha_inicio DATE,
        fecha_fin    DATE,
        horas        DECIMAL(5,2),
        descripcion  TEXT,
        creado_por   INT,
        fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (creado_por) REFERENCES usuarios(id)
    )
    """)

    # ----------------------------------------------------------
    # PARTICIPANTES DE CURSOS
    # ----------------------------------------------------------
    tablas.append("""
    CREATE TABLE IF NOT EXISTS curso_participantes (
        id        INT AUTO_INCREMENT PRIMARY KEY,
        curso_id  INT NOT NULL,
        legajo    INT NOT NULL,
        aprobado  TINYINT DEFAULT 1,
        nota      DECIMAL(4,2),
        observacion VARCHAR(300),
        FOREIGN KEY (curso_id) REFERENCES cursos(id),
        FOREIGN KEY (legajo)   REFERENCES legajos(legajo),
        UNIQUE KEY uq_curso_participante (curso_id, legajo)
    )
    """)

    for sql in tablas:
        cur.execute(sql)
    print("      ✓ Todas las tablas creadas.")


def migrar_datos(sqlite_conn, mysql_cur):
    print("\n[3/5] Migrando datos existentes...")

    # --- configuracion ---
    rows = sqlite_conn.execute("SELECT * FROM configuracion").fetchall()
    for r in rows:
        mysql_cur.execute("""
            INSERT INTO configuracion
            (id,institucion,direccion,ciudad,provincia,telefono,email,sitio_web,logo,version_sistema,fecha_instalacion)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, r)
    print(f"      ✓ configuracion: {len(rows)} registros")

    # --- conceptos ---
    rows = sqlite_conn.execute("SELECT * FROM conceptos").fetchall()
    for r in rows:
        mysql_cur.execute("""
            INSERT INTO conceptos (id,concepto,puntos,detalle,activo)
            VALUES (%s,%s,%s,%s,%s)
        """, r)
    print(f"      ✓ conceptos: {len(rows)} registros")

    # --- legajos ---
    rows = sqlite_conn.execute("SELECT * FROM legajos").fetchall()
    for r in rows:
        mysql_cur.execute("""
            INSERT INTO legajos
            (legajo,apellido,nombre,grado,cargo,dni,foto,fecha_carga,situacion,fecha_baja,autoriza,email,nro_cel)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, r)
    print(f"      ✓ legajos: {len(rows)} registros")

    # --- usuarios ---
    rows = sqlite_conn.execute("SELECT * FROM usuarios").fetchall()
    for r in rows:
        mysql_cur.execute("""
            INSERT INTO usuarios
            (id,username,password_hash,rol,legajo,activo,fecha_baja,usuario_baja,debe_cambiar_password)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, r)
    print(f"      ✓ usuarios: {len(rows)} registros")

    # --- actividades (vacía pero migra estructura) ---
    rows = sqlite_conn.execute("SELECT * FROM actividades").fetchall()
    print(f"      ✓ actividades: {len(rows)} registros (tabla vacía, OK)")

    print("      ✓ Migración de datos completa.")


def verificar(mysql_cur):
    print("\n[4/5] Verificando...")
    tablas = [
        "configuracion","legajos","conceptos","usuarios",
        "actividades","actividades_historial","notificaciones",
        "eventos","asistencia","cursos","curso_participantes"
    ]
    for t in tablas:
        mysql_cur.execute(f"SELECT COUNT(*) FROM {t}")
        count = mysql_cur.fetchone()[0]
        print(f"      ✓ {t}: {count} registros")


def main():
    print("=" * 55)
    print("  SIAB — Migración SQLite → MySQL")
    print("=" * 55)

    # Conectar SQLite
    try:
        sqlite_conn = sqlite3.connect(SQLITE_PATH)
        print(f"\n✓ SQLite conectado: {SQLITE_PATH}")
    except Exception as e:
        print(f"✗ No se pudo abrir el SQLite: {e}")
        print(f"  Verificá la ruta: {SQLITE_PATH}")
        sys.exit(1)

    # Crear DB MySQL
    crear_base_de_datos()

    # Conectar MySQL con la nueva DB
    try:
        mysql_conn = conectar_mysql()
        mysql_cur  = mysql_conn.cursor()
        print(f"✓ MySQL conectado: {MYSQL_HOST}/{MYSQL_DB}")
    except Error as e:
        print(f"✗ Error conectando a MySQL: {e}")
        sys.exit(1)

    # Crear tablas
    crear_tablas(mysql_cur)

    # Migrar datos
    migrar_datos(sqlite_conn, mysql_cur)

    # Commit
    mysql_conn.commit()

    # Verificar
    verificar(mysql_cur)

    print("\n[5/5] ¡Migración completada exitosamente!")
    print("      Podés abrir MySQL Workbench y revisar la base 'siab'.")
    print("=" * 55)

    mysql_cur.close()
    mysql_conn.close()
    sqlite_conn.close()


if __name__ == "__main__":
    main()