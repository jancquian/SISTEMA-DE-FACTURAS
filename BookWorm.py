"""
El presente código “BookWorm.py” fue elaborado por el pasante de Ingeniería en Sistemas Computacionales (ISC)
Juan Carlos Garcia Jimenez (juancarlosgarciajimenez123@gmail.com) para la empresa
Distribución e Ingeniería VAR S.A. de C.V en los años 2023-2024.
"""

import os
import shutil
import xml.etree.ElementTree as Xml
import pandas as panda
import zipfile
import sqlite3
# import psycopg2


class BookWorm:

    def __init__(self, ruta):
        self._conexion = sqlite3.connect(ruta)
        self._cursor = self._conexion.cursor()
        self._ruta_origen = ""
        self._ruta = ruta
        self._pword = "653510d0eddd9282ed16ce4f0ee9e906242f2c0a1bbc60985682e323bd346002"

    # El método estático “convertir_mes” recibe como parámetro una cadena de dos caracteres
    # simbolizando el mes numérico y retorna una cadena con la representación textual resumida
    # de tres caracteres.
    @staticmethod
    def convertir_mes(mes_cadena):

        conversion = [("01", "ENE"), ("02", "FEB"), ("03", "MAR"), ("04", "ABR"), ("05", "MAY"), ("06", "JUN"),
                      ("07", "JUL"), ("08", "AGO"), ("09", "SEP"), ("10", "OCT"), ("11", "NOV"), ("12", "DIC")]

        for dupla in conversion:
            if mes_cadena == dupla[0]:
                return dupla[1]

        return "ERR"

    # El método "crear_directorio" crea un directorio en la ruta especificada si no existe.
    @staticmethod
    def crear_directorio(ruta):
        if not os.path.exists(ruta):
            os.makedirs(ruta)
            return ruta
        else:
            return ruta

    # El método estático “listar_archivos” retorna un objeto tipo lista con el nombre de
    # todos los archivos encontrados en un directorio dado (ruta).
    @staticmethod
    def listar_archivos(ruta):
        try:
            archivos = os.listdir(ruta)
            # Descomentar para ver lista de archivos en consola
            '''
            # Imprimir los nombres de los archivos
            contador = 0

            for archivo in archivos:
                contador = contador + 1
                print("{0} : {1}".format(contador, archivo))
            '''
            return archivos

        except FileNotFoundError:
            print(f"El directorio '{ruta}' no existe.")
        except PermissionError:
            print(f"No tienes permisos para acceder al directorio '{ruta}'.")
        except Exception as e:
            print(f"Ocurrió un error: {e}")

    # El método estático “descomprimir_zip” descomprime un archivo formato “.zip” de una
    # ruta dada y vuelca sus archivos en un directorio destino dado.
    @staticmethod
    def descomprimir_zip(ruta_zip, directorio_destino):

        with zipfile.ZipFile(ruta_zip, 'r') as zip_ref:
            # Extraer todos los contenidos en el directorio de destino
            zip_ref.extractall(directorio_destino)

    # El método “authenticate” compara una cadena de entrada “pwd” con una cadena atributo
    # que funge como contraseña “_pword”; cabe mencionar que el atributo es la representación
    # de la contraseña real, pues el atributo es en realidad el digesto de la contraseña generado
    # con la función hash SHA256.
    def authenticate(self, pwd):
        return True if pwd == self._pword else False

    # El método “crecon_db” genera la base de datos con las tablas necesarias para utilizar el programa; debido
    # a que la creación de este programa se realizo previamente al diseño de la base de datos del servidor, será
    # necesario hacer los ajustes necesarios para su incorporación al sistema final.
    def crecon_db(self):

        # El siguiente código comentado fue implementado para la modificación de la base de datos
        # ante un cambio de requerimientos durante el desarrollo; se mantiene para futuras
        # modificaciones y/o correcciones.

        # self._cursor.execute('''ALTER TABLE CONCEPTOS_NUEVOS ADD COLUMN CLIENTE TEXT DEFAULT "DEF_CLE/C;"''')
        # self._cursor.execute('''ALTER TABLE CONCEPTOS ADD COLUMN CLIENTE TEXT DEFAULT "DEF_CLE/C;"''')

        # self._cursor.execute('''ALTER TABLE FACTURAS ADD COLUMN NO_COTIZACION TEXT DEFAULT "NO_COT/C";''')
        # self._cursor.execute('''ALTER TABLE FACTURAS ADD COLUMN CARPETA TEXT DEFAULT "CARPETA/C";''')
        # self._cursor.execute('''ALTER TABLE FACTURAS_NUEVAS ADD COLUMN NO_COTIZACION TEXT DEFAULT "NO_COT/C";''')
        # self._cursor.execute('''ALTER TABLE FACTURAS_NUEVAS ADD COLUMN CARPETA TEXT DEFAULT "CARPETA/C";''')
        # self._cursor.execute('''ALTER TABLE FACTURAS RENAME COLUMN OBSERVACIONES TO RECIBIO_DIVAR;''')
        # self._cursor.execute('''ALTER TABLE FACTURAS_NUEVAS RENAME COLUMN OBSERVACIONES TO RECIBIO_DIVAR;''')

        self._cursor.execute('''
            CREATE TABLE IF NOT EXISTS CLIENTES (
                CLAVE TEXT PRIMARY KEY
            );
        ''')

        self._cursor.execute('''
            CREATE TABLE IF NOT EXISTS CATEGORIA (
                CAT TEXT PRIMARY KEY
            );
        ''')

        self._cursor.execute('''
            CREATE TABLE IF NOT EXISTS PERSONAL (
                EMPLEADO CHARVAR(30) PRIMARY KEY
            );
        ''')

        self._cursor.execute('''
            CREATE TABLE IF NOT EXISTS FACTURAS (
                NOMBRE_XML TEXT PRIMARY KEY,
                NOMBRE TEXT,
                FOLIO TEXT,
                SERIE TEXT,
                FECHA DATE,
                TOTAL FLOAT,
                SUBTOTAL FLOAT,
                REGIMEN_FISCAL TEXT,
                RFC TEXT,
                IMPUESTOS TEXT,
                TIPO_COMPROBANTE TEXT,
                FORMA_PAGO TEXT,
                METODO_PAGO TEXT,
                TRANSFERENCIA TEXT DEFAULT "TRANN/A",
                NO_COTIZACION TEXT DEFAULT "NO_COT/C",
                CATEGORIA_FACTURA TEXT DEFAULT "CAT_FN/A",
                NUMERO_CLIENTE TEXT DEFAULT "NO_CN/A",
                RECIBIO_DIVAR TEXT DEFAULT "RES/A",
                CARPETA TEXT DEFAULT "CARPETA/C",
                FOREIGN KEY (NUMERO_CLIENTE) REFERENCES CLIENTES(CLAVE),
                FOREIGN KEY (CATEGORIA_FACTURA) REFERENCES CATEGORIA(CAT)
                
            );
        ''')

        self._cursor.execute('''
            CREATE TABLE IF NOT EXISTS FACTURAS_NUEVAS (
                NOMBRE_XML TEXT PRIMARY KEY,
                NOMBRE TEXT,
                FOLIO TEXT,
                SERIE TEXT,
                FECHA DATE,
                TOTAL FLOAT,
                SUBTOTAL FLOAT,
                REGIMEN_FISCAL TEXT,
                RFC TEXT,
                IMPUESTOS TEXT,
                TIPO_COMPROBANTE TEXT,
                FORMA_PAGO TEXT,
                METODO_PAGO TEXT,
                TRANSFERENCIA TEXT DEFAULT "TRANN/A",
                NO_COTIZACION TEXT DEFAULT "NO_COT/C",
                CATEGORIA_FACTURA TEXT DEFAULT "CAT_FN/A",
                NUMERO_CLIENTE TEXT DEFAULT "NO_CN/A",
                RECIBIO_DIVAR TEXT DEFAULT "RES/A",
                CARPETA TEXT DEFAULT "CARPETA/C",
                FOREIGN KEY (NUMERO_CLIENTE) REFERENCES CLIENTES(CLAVE),
                FOREIGN KEY (CATEGORIA_FACTURA) REFERENCES CATEGORIA(CAT)
            );
        ''')

        self._cursor.execute('''
            CREATE TABLE IF NOT EXISTS CONCEPTOS (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                NOMBRE_XML TEXT,
                CLAVE_PROD_SERV TEXT,
                CLIENTE TEXT DEFAULT "DEF_CLE/C",
                DESCRIPCION TEXT,
                CANTIDAD FLOAT,
                VALOR_UNITARIO FLOAT,
                FOREIGN KEY (NOMBRE_XML) REFERENCES FACTURAS(NOMBRE_XML)
                -- Agrega más columnas según tus necesidades
            );
        ''')

        self._cursor.execute('''
            CREATE TABLE IF NOT EXISTS CONCEPTOS_NUEVOS (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                NOMBRE_XML TEXT,
                CLAVE_PROD_SERV TEXT,
                CLIENTE TEXT DEFAULT "DEF_CLE/C",
                DESCRIPCION TEXT,
                CANTIDAD FLOAT,
                VALOR_UNITARIO FLOAT,
                FOREIGN KEY (NOMBRE_XML) REFERENCES FACTURAS(NOMBRE_XML)
                -- Agrega más columnas según tus necesidades
            );
        ''')

    # El método “agregar_factura” hace INSERT de un registro a la tabla de FACTURAS o FACTURAS_NUEVAS; esto
    # último depende del parámetro “tabla” proporcionado.
    def agregar_factura(self, nom_xml, nom, folio, serie, fecha, total, stotal, r_fis, rfc, impu, tipo_com,
                        forma_pago, metodo_pago, tabla):
        # Utiliza parámetros en la consulta para evitar SQL injection
        self._cursor.execute('''
            INSERT INTO {0} (
                NOMBRE_XML,
                NOMBRE,
                FOLIO,
                SERIE,
                FECHA,
                TOTAL,
                SUBTOTAL,
                REGIMEN_FISCAL,
                RFC,
                IMPUESTOS,
                TIPO_COMPROBANTE,
                FORMA_PAGO,
                METODO_PAGO
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        '''.format(tabla), (nom_xml, nom, folio, serie, fecha,
                            total, stotal, r_fis, rfc, impu, tipo_com, forma_pago, metodo_pago))

        self._conexion.commit()

    # El método “agregar_concepto” hace INSERT de un registro a la tabla de CONCEPTOS o CONCEPTOS_NUEVOS; esto
    # último depende del parámetro “tabla” proporcionado.
    def agregar_concepto(self, nom_xml, clave, descripcion, cantidad, valor_uni, tabla):
        # Utiliza parámetros en la consulta para evitar SQL injection
        self._cursor.execute('''
            INSERT INTO {0} (
                NOMBRE_XML,
                CLAVE_PROD_SERV,
                DESCRIPCION,
                CANTIDAD,
                VALOR_UNITARIO
            )
            VALUES (?, ?, ?, ?, ?)
        '''.format(tabla), (nom_xml, clave, descripcion, cantidad, valor_uni))

    # El método estático  “clasificar_factura_p”  a través de un archivo “.xml” verifica si la factura
    # dada es del tipo “P” y retorna un diccionario con los datos de relevancia para la empresa.
    @staticmethod
    def clasificar_factura_p(root, nombre_archivo):

        diccionario = {}
        conceptos = []

        namespace = '{' + root.tag.split('}')[0][1:] + '}'
        child_emisor = root.find(namespace + "Emisor")
        child_conceptos = root.find(namespace + "Conceptos")

        diccionario['NombreXML'] = nombre_archivo

        lista_de_campos = ['Nombre', 'RegimenFiscal', 'Rfc',]

        for campo in lista_de_campos:
            try:
                diccionario[campo] = child_emisor.attrib[campo]
            except (KeyError, AttributeError):
                diccionario[campo] = "N/A"

        lista_de_campos = ['Folio', 'Serie', 'Total', 'SubTotal']
        for campo in lista_de_campos:
            try:
                diccionario[campo] = root.attrib[campo]
            except (KeyError, AttributeError):
                diccionario[campo] = "N/A"

        try:
            diccionario['Fecha'] = root.attrib['Fecha']
        except (KeyError, AttributeError):
            diccionario['Fecha'] = "N/A"

        diccionario['Impuestos'] = "N/A"
        diccionario['TipoDeComprobante'] = "P"
        diccionario['FormaPago'] = "N/A"
        diccionario['MetodoPago'] = "N/A"

        id_aux = 0
        for concepto in child_conceptos:

            dicc_aux = dict()
            dicc_aux['id'] = str(id_aux)
            id_aux = id_aux + 1

            lista_campos_conceptos = ['ClaveProdServ', 'Descripcion', 'Cantidad', 'ValorUnitario']

            for campo_c in lista_campos_conceptos:
                try:
                    dicc_aux[campo_c] = concepto.attrib[campo_c]
                except (KeyError, AttributeError):
                    dicc_aux[campo_c] = "N/A"

            conceptos.append(dicc_aux)

        diccionario['Conceptos'] = conceptos

        # Descomentar para ver detalles de factura en consola
        '''
        print("_________________________________________________________")
        for key in diccionario.keys():
            print(key + ': ', diccionario[key])
        '''

        return diccionario

    # El método estático  “clasificar_factura_i_e”  a través de un archivo “.xml” verifica si la factura
    # dada es del tipo “I” o "E" y retorna un diccionario con los datos de relevancia para la empresa.
    @staticmethod
    def clasificar_factura_i_e(root, nombre_archivo):

        diccionario = {}
        conceptos = []

        namespace = '{' + root.tag.split('}')[0][1:] + '}'

        child_emisor = root.find(namespace + "Emisor")
        child_impuesto = root.find(namespace + "Impuestos")
        child_concepto = root.find(namespace + "Conceptos")

        diccionario['NombreXML'] = nombre_archivo

        lista_de_campos = ['Nombre', 'RegimenFiscal', 'Rfc']

        for campo in lista_de_campos:
            try:
                diccionario[campo] = child_emisor.attrib[campo]
            except (KeyError, AttributeError):
                diccionario[campo] = "N/A"

        lista_de_campos = ['Folio', 'Serie', 'Total', 'SubTotal', 'TipoDeComprobante', 'FormaPago', 'MetodoPago']

        for campo in lista_de_campos:
            try:
                diccionario[campo] = root.attrib[campo]
            except (KeyError, AttributeError):
                diccionario[campo] = "N/A"

        try:
            diccionario['Fecha'] = root.attrib['Fecha']
        except (KeyError, AttributeError):
            diccionario['Fecha'] = "N/A"

        try:
            diccionario["Impuestos"] = child_impuesto.attrib['TotalImpuestosTrasladados']
        except (KeyError, AttributeError):
            diccionario["Impuestos"] = "N/A"

        id_aux = 0
        for concepto in child_concepto:

            dicc_aux = dict()
            dicc_aux['id'] = str(id_aux)
            id_aux = id_aux + 1

            lista_campos_conceptos = ['ClaveProdServ', 'Descripcion', 'Cantidad', 'ValorUnitario']

            for campo_c in lista_campos_conceptos:
                try:
                    dicc_aux[campo_c] = concepto.attrib[campo_c]
                except (KeyError, AttributeError):
                    dicc_aux[campo_c] = "N/A"

            conceptos.append(dicc_aux)

        diccionario['Conceptos'] = conceptos

        # Descomentar para ver los detalles de la factura en consola
        '''
        print("_________________________________________________________")
        for key in diccionario.keys():
            print(key + ': ', diccionario[key])
        '''

        return diccionario

    # El método "get_column_desc" devuelve una lista que contiene el primer elemento de cada tupla
    # en self._cursor.description.
    def get_column_desc(self):
        return [nombre[0] for nombre in self._cursor.description]

    # El método “exportar_xlsx” genera un archivo .xlsx dada una tabla, si existe una lista de columnas
    # no deseadas para la generación, estas son excluidas.
    def exportar_xlsx(self, tabla, columnas, nombre, columnas_excluidas=None):

        self.crear_directorio("xlsx")
        # print("C_DTF:", columnas)
        data_frame = panda.DataFrame(tabla, columns=columnas)
        if columnas_excluidas is not None:
            data_frame = data_frame.drop(columnas_excluidas, axis=1)
        data_frame.to_excel("xlsx/" + nombre + ".xlsx", index=False)

    def consulta_sql_rango_fecha(self, tabla, criterio, anio):

        fecha_inicial = criterio[6:10] + '-' + criterio[3:5] + '-' + criterio[0:2]
        fecha_final = criterio[17:21] + '-' + criterio[14:16] + '-' + criterio[11:13]
        # print(fecha_inicial)
        # print(fecha_final)

        if anio == "0000":
            # Se hace la evaluacion a dos fechas.
            consulta = "SELECT * FROM {0} WHERE FECHA BETWEEN ? AND ? ORDER BY FECHA".format(tabla)
            self._cursor.execute(consulta, (fecha_inicial, fecha_final))
            resultado = self._cursor.fetchall()

            # Si la consulta está vacia, entonces intenta solo con la primera fecha
            if resultado:
                return resultado
            else:
                consulta = "SELECT * FROM {0} WHERE FECHA = ?".format(tabla)
                self._cursor.execute(consulta, (fecha_inicial,))
                resultado = self._cursor.fetchall()
                return resultado
        else:
            # Se hace la evaluacion a dos fechas.
            consulta = '''SELECT * FROM {0} WHERE FECHA BETWEEN ? AND ? 
            AND FECHA LIKE '%{1}%' ORDER BY FECHA'''.format(tabla, anio)
            self._cursor.execute(consulta, (fecha_inicial, fecha_final))
            resultado = self._cursor.fetchall()

            # Si la consulta está vacia, entonces intenta solo con la primera fecha
            if resultado:
                return resultado
            else:
                consulta = '''SELECT * FROM {0} WHERE FECHA = ? 
                AND FECHA LIKE '%{1}%' ORDER BY FECHA'''.format(tabla, anio)
                self._cursor.execute(consulta, (fecha_inicial,))
                resultado = self._cursor.fetchall()
                return resultado

    # El método “consulta_sql_concepto_por_xml”  hace una consulta de los conceptos asociados al nombre
    # de un XML; básicamente recupera los conceptos asociados a una factura. NO ESTÁ PROTEGIDO CONTRA INYECCION SQL.
    def consulta_sql_concepto_por_xml(self, tabla_conceptos, criterio):
        consulta = '''SELECT * FROM {0} WHERE NOMBRE_XML = ? '''.format(tabla_conceptos)
        self._cursor.execute(consulta, (criterio,))
        return self._cursor.fetchall()

    # El método “consulta_sql_concepto”  hace una consulta de los conceptos dado un criterio de busqueda; básicamente
    # recupera las facturas que inciden con el criterio. NO ESTÁ PROTEGIDO CONTRA INYECCION SQL.
    def consulta_sql_concepto(self, tabla, tabla_conceptos, criterio):

        consulta = '''SELECT DISTINCT F.* 
        FROM {0} F JOIN {1} C ON F.NOMBRE_XML = C.NOMBRE_XML 
        WHERE C.DESCRIPCION LIKE ?'''.format(tabla, tabla_conceptos)
        self._cursor.execute(consulta, ['%' + criterio + '%'])
        return self._cursor.fetchall()

    # El método “consulta_sql_no_pagadas” consulta las facturas no pagadas tomándose como criterio la cadena
    # por defecto “TRANN/A” para indicar que no ha sido pagada. ESTE MÉTODO NO ESTÁ PROTEGIDO CONTRA INYECCION SQL.
    def consulta_sql_no_pagadas(self, tabla, anio):

        if anio == "0000":
            consulta = "SELECT * FROM {0} WHERE TRANSFERENCIA = 'TRANN/A'".format(tabla)
            self._cursor.execute(consulta)
        else:
            consulta = "SELECT * FROM {0} WHERE TRANSFERENCIA = 'TRANN/A' AND FECHA LIKE ?".format(tabla)
            self._cursor.execute(consulta, (['%' + anio + '%']))

        return self._cursor.fetchall()

    # El método “consulta_sql_por_campo” hace una consulta basica dando la tabla y el campo a consultar junto
    # a un criterio de busqueda. ESTE MÉTODO NO ESTÁ PROTEGIDO CONTRA INYECCION SQL.
    def consulta_sql_por_campo(self, tabla, campo, criterio):

        consulta = "SELECT * FROM {0} WHERE {1} = ?".format(tabla, campo)
        self._cursor.execute(consulta, (criterio,))
        return self._cursor.fetchall()

    # El método “consulta_panda” hace una consulta sucesiva a un conjunto de datos (dados en formato de tabla
    # indicando los datos "tabla" y las columnas que la forman "columnas") aplicandole consecutivamente una
    # serie de filtros ("criteria") hasta obtener la tabla deseada; esta se da en forma de lista en un retorno.
    def consulta_panda(self, tabla, columnas, criteria):

        data_frame = panda.DataFrame(tabla, columns=columnas)
        # Inicializar el DataFrame de resultados con el DataFrame original
        resultado_consulta = data_frame

        # Iterar sobre cada criterio en la lista de criterios
        for crit in criteria:
            columna = crit[0]  # Nombre de la columna
            cadena_criterio = crit[1]  # Criterio de búsqueda

            # Aplicar el filtro sobre el DataFrame acumulado de resultados

            # Verificar si el criterio es un entero
            print("CADENA_CRITERIO_EN_TUNRO: ", cadena_criterio, type(cadena_criterio))
            if columna in ['TOTAL', 'SUBTOTAL']:
                try:
                    cadena_criterio = float(cadena_criterio)
                    resultado_consulta = resultado_consulta[(resultado_consulta[columna] <= cadena_criterio*1.10) &
                                                            (resultado_consulta[columna] >= cadena_criterio*0.90)]
                except ValueError:
                    print("omitiendo operacion")

                # Aplicar filtro para igualdad numérica
                # resultado_consulta = resultado_consulta[resultado_consulta[columna] == cadena_criterio]
            else:
                # Aplicar filtro para cadenas
                resultado_consulta = resultado_consulta[
                    resultado_consulta[columna].str.contains(cadena_criterio, case=False)]

        # Eliminar duplicados del DataFrame de resultados (si es necesario)
        resultado_consulta = resultado_consulta.drop_duplicates()
        print("RESULTADO: ", list(resultado_consulta.to_records(index=False)))
        return list(resultado_consulta.to_records(index=False))

    def consulta_sql(self, tabla, condiciones, anio=None):

        # Construir la parte del WHERE dinámicamente
        where_clause = ""
        for condicion in condiciones:
            print("CONDICION BACK: ", condicion)
            columna, criterio = condicion
            if where_clause:
                where_clause += " AND "
            where_clause += f"{columna} LIKE '%{criterio}%'"

        # Construir la consulta completa
        if anio is None:
            consulta = f"SELECT * FROM {tabla} WHERE {where_clause};"
        else:
            consulta = f"SELECT * FROM {tabla} WHERE {where_clause} AND FECHA LIKE '%{anio}%';"

        # Ejecutar la consulta
        self._cursor.execute(consulta)

        # Retornar los resultados
        return self._cursor.fetchall()

    # El método “consulta_sql_completa" realiza una consulta SQL completa de una tabla dada. NO SE ENCUENTRA PROTEGIDA
    # DE INYECCION SQL.
    def consulta_sql_completa(self, tabla):
        self._cursor.execute("SELECT * FROM {0};".format(tabla))
        return self._cursor.fetchall()

    # El método “consulta_sql_faltantes" realiza una consulta SQL completa de una tabla dada. NO SE ENCUENTRA
    # PROTEGIDA DE INYECCION SQL.
    def consulta_sql_faltantes(self, tabla, anio):
        if tabla == "FACTURAS":
            if anio == "0000":
                self._cursor.execute('''SELECT * FROM FACTURAS WHERE
                    TRANSFERENCIA LIKE "TRANN/A" OR
                    CATEGORIA_FACTURA LIKE "CAT_FN/A" OR
                    NUMERO_CLIENTE LIKE "NO_CN/A"OR
                    RECIBIO_DIVAR LIKE "RES/A"
                ''')
            else:
                self._cursor.execute('''SELECT * FROM FACTURAS WHERE
                    (TRANSFERENCIA LIKE "TRANN/A" OR
                    CATEGORIA_FACTURA LIKE "CAT_FN/A" OR
                    NUMERO_CLIENTE LIKE "NO_CN/A"OR
                    RECIBIO_DIVAR LIKE "RES/A") AND
                    FECHA LIKE '%{0}%'
                '''.format(anio))
        else:
            if anio == "0000":
                self._cursor.execute('''SELECT * FROM FACTURAS_NUEVAS WHERE
                    TRANSFERENCIA LIKE "TRANN/A" OR
                    CATEGORIA_FACTURA LIKE "CAT_FN/A" OR
                    NUMERO_CLIENTE LIKE "NO_CN/A"OR
                    RECIBIO_DIVAR LIKE "RES/A"
                ''')
            else:
                self._cursor.execute('''SELECT * FROM FACTURAS_NUEVAS WHERE
                    (TRANSFERENCIA LIKE "TRANN/A" OR
                    CATEGORIA_FACTURA LIKE "CAT_FN/A" OR
                    NUMERO_CLIENTE LIKE "NO_CN/A"OR
                    RECIBIO_DIVAR LIKE "RES/A") AND
                    FECHA LIKE '%{0}%'
                '''.format(anio))
        return self._cursor.fetchall()

    # El método “consulta_sql_actualizar_registro" realiza un UPDATE de los registros de las facturas.
    def consulta_sql_actualizar_registro(self, trans, numero_cliente, r_divar, categoria, no_cot, carpeta, nombre_xml,
                                         tabla):
        if tabla == "FACTURAS":
            update = '''UPDATE FACTURAS
                SET TRANSFERENCIA = ?,
                CATEGORIA_FACTURA = ?,
                NUMERO_CLIENTE = ?,
                RECIBIO_DIVAR = ?,
                NO_COTIZACION = ?,
                CARPETA = ?
                WHERE NOMBRE_XML = ?
            '''
        else:
            update = '''UPDATE FACTURAS_NUEVAS
                SET TRANSFERENCIA = ?,
                CATEGORIA_FACTURA = ?,
                NUMERO_CLIENTE = ?,
                RECIBIO_DIVAR = ?,
                NO_COTIZACION = ?,
                CARPETA = ?
                WHERE NOMBRE_XML = ?
            '''

        self._cursor.execute(update, (trans.upper(), categoria.upper(), numero_cliente.upper(),
                                      r_divar.upper(), no_cot, carpeta, nombre_xml))
        self._conexion.commit()

    # El método “consulta_sql_actualizar_concepto" realiza un UPDATE de los registros de los conceptos de una factura.
    def consulta_sql_actualizar_concepto(self, identificador, cliente, tabla):
        if tabla == "FACTURAS":
            update = '''UPDATE CONCEPTOS SET CLIENTE = ? WHERE ID = ?'''
        else:
            update = '''UPDATE CONCEPTOS_NUEVOS SET CLIENTE = ? WHERE ID = ?'''

        self._cursor.execute(update, (cliente.upper(), identificador))
        self._conexion.commit()

    # El método “consulta_sql_mover_registros_index” transfiere los registros de las facturas de la tabla
    # FACTURAS_NUEVAS a FACTURAS y sus respectivos conceptos de CONCEPTOS_NUEVOS a CONCEPTOS.
    def consulta_sql_mover_registros_index(self):

        self._cursor.execute("SELECT * FROM FACTURAS_NUEVAS")
        facturas = self._cursor.fetchall()

        insert = '''INSERT INTO FACTURAS (
            NOMBRE_XML, NOMBRE, FOLIO, SERIE, FECHA, TOTAL, SUBTOTAL, REGIMEN_FISCAL, RFC, IMPUESTOS, TIPO_COMPROBANTE,
            FORMA_PAGO, METODO_PAGO, TRANSFERENCIA, NO_COTIZACION, CATEGORIA_FACTURA, NUMERO_CLIENTE, RECIBIO_DIVAR, 
            CARPETA) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''

        insert_c = '''INSERT INTO CONCEPTOS (
                NOMBRE_XML,CLAVE_PROD_SERV, CLIENTE, DESCRIPCION, CANTIDAD, VALOR_UNITARIO)
                VALUES (?, ?, ?, ?, ?, ?)'''

        for factura in facturas:
            try:
                # print(factura)
                self._cursor.execute(insert, factura)
                get_c = "SELECT * FROM CONCEPTOS_NUEVOS WHERE NOMBRE_XML = ?"
                self._cursor.execute(get_c, (factura[0],))
                conceptos = self._cursor.fetchall()

                for concepto in conceptos:
                    self._cursor.execute(insert_c, concepto[1:])

            except sqlite3.IntegrityError as e:
                print("FACTURA YA PRESENTE EN 'FACTURAS' (MOV): ", e)

        self._cursor.execute("DELETE FROM FACTURAS_NUEVAS")
        self._cursor.execute("DELETE FROM CONCEPTOS_NUEVOS")
        self._conexion.commit()

    # El método “consulta_sql_obtener_anios” obtiene los años en los que se encuentran
    # las facturas; es decir, retorna una lista con los años disponibles en las facturas.
    def consulta_sql_obtner_anios(self, tabla_actual):

        consulta = "SELECT DISTINCT strftime('%Y', FECHA) FROM {0}".format(tabla_actual)
        self._cursor.execute(consulta)
        anios = list()
        for tupla in self._cursor.fetchall():
            anios.append(tupla[0])
        # Linea de visualizacion en consola
        return anios

    # El método “consulta_sql_por_anio” hace una consulta SQL de una tabla con base en el
    # año que se le proporcione.
    def consulta_sql_por_anio(self, anio, tabla_actual):
        consulta = "SELECT * FROM {0} WHERE strftime('%Y', FECHA) = '{1}'".format(tabla_actual, anio)
        self._cursor.execute(consulta)
        return self._cursor.fetchall()

    '''
    def obtener_registos(self):
        return self._cursor.fetchall()

    def imprimir_consulta(self):

        rows = self._cursor.fetchall()
        # print(" ".join([description[0] for description in self._cursor.description]))
        
        #for row in rows:
        #    print(" ".join(map(str, row)))
        
    '''
    def agregar_a_reg_tab_simple(self, tabla, valor):
        if tabla == "CLIENTES":
            try:
                self._cursor.execute('''INSERT INTO CLIENTES (CLAVE) VALUES (?)''', (valor,))
            except sqlite3.IntegrityError as e:
                print("CLIENTE YA PRESENTE EN CLIENTES: ", e)
        elif tabla == "CATEGORIA":
            try:
                self._cursor.execute('''INSERT INTO CATEGORIA (CAT) VALUES (?)''', (valor,))
            except sqlite3.IntegrityError as e:
                print("CATEGORIA YA PRESENTE EN CATEGORIAS: ", e)

    # El método “agregar_facturas_bd”  agrega las facturas en formato .xml de una ruta dada a la base de datos
    # junto con sus conceptos asociados.
    def agregar_facturas_bd(self, ruta_origen, nombre_tabla_facturas, nombre_tabla_conceptos):

        # Se obtiene una lista de cadenas con los nombres de los archivos presentes en la ruta origen
        archivos = self.listar_archivos(ruta_origen)

        # Para cada uno de los nombres de archivo...

        for archivo in archivos:
            # GENERAL...
            # Se abre el archivo XML con el nombre del archivo y se obtiene su root.
            tree = Xml.parse(ruta_origen + '/' + archivo)
            root = tree.getroot()

            # AGREGAR A BASE DE DATOS DEL PROGRAMA A BASE DE DATOS DEL PROGRAMA

            if root.attrib['TipoDeComprobante'] == 'I' or root.attrib['TipoDeComprobante'] == 'E':
                factura = self.clasificar_factura_i_e(root, archivo)
            elif root.attrib['TipoDeComprobante'] == 'P':
                factura = self.clasificar_factura_p(root, archivo)
            else:

                factura = {}

            try:
                self.agregar_factura(factura['NombreXML'], factura['Nombre'], factura['Folio'], factura['Serie'],
                                     factura['Fecha'][0:10], factura['Total'], factura['SubTotal'],
                                     factura['RegimenFiscal'], factura['Rfc'], factura['Impuestos'],
                                     factura['TipoDeComprobante'], factura['FormaPago'], factura['MetodoPago'],
                                     nombre_tabla_facturas)

                for concepto in factura['Conceptos']:
                    self.agregar_concepto(factura['NombreXML'], concepto['ClaveProdServ'], concepto['Descripcion'],
                                          concepto['Cantidad'], concepto['ValorUnitario'], nombre_tabla_conceptos)

            except sqlite3.IntegrityError as e:
                print("")
                print("REGISTRO YA PRESENTE EN FACTURAS: ", e)

            except KeyError as e:
                # Manejo específico para errores de acceso a claves
                print(archivo)
                print("TIPO DE FACTURA '{0}'".format(root.attrib['TipoDeComprobante']))
                os.remove(ruta_origen + '/' + archivo)
                print("Error de acceso a la clave:", e)

    # El método “clasificar_facturas_archivos”  organiza a nivel de archivos las facturas en formato .xml
    # de una ruta origen en una ruta destino; la organización se hace en años mediante carpetas.
    def clasificar_facturas_archivos(self, ruta_origen, ruta_destino):

        # Se crea la carpeta destino en caso de no existir
        self.crear_directorio(ruta_destino)

        # Se obtiene una lista de cadenas con los nombres de los archivos presentes en la ruta origen
        archivos = self.listar_archivos(ruta_origen)

        # Para cada uno de los nombres de archivo...

        for archivo in archivos:
            # GENERAL...
            # Se abre el archivo XML con el nombre del archivo y se obtiene su root.
            try:
                tree = Xml.parse(ruta_origen + '/' + archivo)
                root = tree.getroot()

                # CLASIFICACIÓN A NIVEL FICHEROS...
                # A traves del diccionaro (obtenido con el root) se obtiene el dato relativo a la fecha(str).

                diccionario = root.attrib

                mes = self.convertir_mes(diccionario['Fecha'][5:7])
                year = diccionario['Fecha'][0:4]

                # Se crea el directorio destino al cual va ir el archivo en caso de no existir.

                ruta_move = self.crear_directorio(ruta_destino + '/' + year + '/' + "CDFI " + mes + " " + year)

                if os.path.exists(ruta_move + '/' + archivo):
                    os.remove(ruta_origen + '/' + archivo)
                else:
                    shutil.move(ruta_origen + '/' + archivo, ruta_move)
            except Xml.ParseError as e:
                print(f"Error al analizar XML Parser: {e}")
                print(archivo)
            except KeyError as e:
                print(f"Error al analizar XML KeyError: {e}")
                print(archivo)

    # El método "cerrar_conexion" cierra la conexion con la base de datos local
    def cerrar_conexion(self):
        self._conexion.close()
