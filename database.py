import mysql.connector
from mysql.connector import Error

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