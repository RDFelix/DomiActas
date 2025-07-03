import sqlite3
import os

def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        print(f"Conexión a la base de datos '{db_file}' establecida con éxito.")
    except sqlite3.Error as e:
        print(f"Error al conectar a la base de datos: {e}")
    return conn

def create_table(conn):
    try:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS actas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cedula TEXT NOT NULL,
                nombres TEXT NOT NULL,
                apellidos TEXT NOT NULL,
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

def insert(conn, libro):
    sql = ''' INSERT INTO libros(titulo,autor,anio_publicacion)
              VALUES(?,?,?) '''
    try:
        cursor = conn.cursor()
        cursor.execute(sql, libro)
        conn.commit()
        print(f"Libro '{libro[0]}' insertado con éxito. ID: {cursor.lastrowid}")
        return cursor.lastrowid
    except sqlite3.Error as e:
        print(f"Error al insertar el libro '{libro[0]}': {e}")
        return None

def select_all(conn, table):
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM {table}")
        rows = cursor.fetchall()

        if rows:
            # Obtener los nombres de las columnas
            col_names = [description[0] for description in cursor.description]
            print(f"\n----------------- TABLA {table.upper()} -----------------")
            print(" | ".join(col_names))
            print("-" * (len(" | ".join(col_names)) + 5))
            for row in rows:
                print(" | ".join(str(item) for item in row))
            print("-" * (len(" | ".join(col_names)) + 5))
        else:
            print(f"\nNo hay registros en la tabla '{table}'.")
    except sqlite3.Error as e:
        print(f"Error al seleccionar: {e}")

def main():
    database = "biblioteca.db"

    if os.path.exists(database):
        os.remove(database)
        print(f"Archivo de base de datos '{database}' existente eliminado para un inicio limpio.")


    conn = create_connection(database)

    if conn:
        create_table(conn)

        select_all_books(conn)

        conn.close()
        print(f"\nConexión a la base de datos '{database}' cerrada.")
    else:
        print("No se pudo establecer la conexión a la base de datos. Saliendo.")

if __name__ == '__main__':
    main()
