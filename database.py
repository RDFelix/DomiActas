import sqlite3
import os

database = "domiactas.db"

def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        print(f"Conexión a la base de datos '{db_file}' establecida con éxito.")
    except sqlite3.Error as e:
        print(f"Error al conectar a la base de datos: {e}")
    return conn

def create_database():
    """"
    if os.path.exists(database):
        os.remove(database)
        print(f"Archivo de base de datos '{database}' existente eliminado para un inicio limpio.")
    """

    conn = create_connection(database)

    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS actas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cedula TEXT NOT NULL,
                    nombre_completo TEXT NOT NULL,
                    zona TEXT NOT NULL,
                    formato TEXT NOT NULL,
                    fecha TEXT NOT NULL,
                    observaciones TEXT NOT NULL,
                    estado TEXT NOT NULL
                );
            ''')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS descuentos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    empleado TEXT NOT NULL,
                    fecha TEXT NOT NULL,
                    cantidad TEXT NOT NULL,
                    valor_total TEXT NOT NULL,
                    numero_cuotas TEXT NOT NULL,
                    valor_cuota TEXT NOT NULL,
                    saldo_novedad TEXT NOT NULL,
                    observacion TEXT NOT NULL,
                    zona TEXT NOT NULL
                );
            ''')

            conn.commit()
        except sqlite3.Error as e:
            print(f"Error al crear la tabla: {e}")
        finally:
            conn.close()
            print(f"\nConexión a la base de datos '{database}' cerrada.")
    else:
        print("No se pudo establecer la conexión a la base de datos. Saliendo.")

def insert_acta(values):
    conn = create_connection(database)

    if conn:
        sql = "INSERT INTO actas (cedula, nombre_completo, zona, formato, fecha, observaciones, estado)VALUES (?, ?, ?, ?, ?, ?, ?)"

        try:
            cursor = conn.cursor()
            cursor.execute(sql, values)
            conn.commit()
            print(f"Registro insertado con éxito en 'actas'. ID: {cursor.lastrowid}")
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al insertar en 'actas': {e}")
            conn.close()
            print(f"\nConexión a la base de datos '{database}' cerrada.")
            return None
    else:
        print("No se pudo establecer la conexión a la base de datos. Saliendo.")
    

def insert_descuento(values):
    conn = create_connection(database)

    if conn:
        sql = "INSERT INTO descuentos (empleado, fecha, cantidad, valor_total, numero_cuotas, valor_cuota, saldo_novedad, observacion, zona) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"

        try:
            cursor = conn.cursor()
            cursor.execute(sql, values)
            conn.commit()
            print(f"Registro insertado con éxito en 'descuentos'. ID: {cursor.lastrowid}")
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al insertar en 'descuentos': {e}")
            conn.close()
            print(f"\nConexión a la base de datos '{database}' cerrada.")
            return None
    else:
        print("No se pudo establecer la conexión a la base de datos. Saliendo.")
    
def get_all_actas():
    conn = create_connection(database)
    actas = []
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM actas")
            actas = cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Error al obtener actas: {e}")
        finally:
            conn.close()
    return actas

def get_all_descuentos():
    conn = create_connection(database)
    descuentos = []
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM descuentos")
            descuentos = cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Error al obtener descuentos: {e}")
        finally:
            conn.close()
    return descuentos

def delete_acta(acta_id):
    conn = create_connection(database)
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM actas WHERE id = ?", (acta_id,))
            conn.commit()
            print(f"Acta con ID {acta_id} eliminada con éxito.")
        except sqlite3.Error as e:
            print(f"Error al eliminar acta: {e}")
        finally:
            conn.close()
    else:
        print("No se pudo establecer la conexión a la base de datos. Saliendo.")

def delete_descuento(descuento_id):
    conn = create_connection(database)
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM descuentos WHERE id = ?", (descuento_id,))
            conn.commit()
            print(f"Descuento con ID {descuento_id} eliminado con éxito.")
        except sqlite3.Error as e:
            print(f"Error al eliminar descuento: {e}")
        finally:
            conn.close()
    else:
        print("No se pudo establecer la conexión a la base de datos. Saliendo.")