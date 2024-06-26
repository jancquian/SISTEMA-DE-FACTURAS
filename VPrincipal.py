"""
El presente código “VPrincipal.py” fue elaborado por el pasante de Ingeniería en Sistemas Computacionales (ISC)
Juan Carlos Garcia Jimenez (juancarlosgarciajimenez123@gmail.com) para la empresa
Distribución e Ingeniería VAR S.A. de C.V en los años 2023-2024.
"""

import tkinter as tk
import ToolTip as Toip
import hashlib
from tkinter import ttk
from tkinter import filedialog
from ttkwidgets.autocomplete import AutocompleteCombobox
# from ttkwidgets import CheckboxTreeview
# from ttkwidgets import ScrolledListbox


class VPrincipal:
    def __init__(self, groot, worm):
        self._worm = worm
        self._factura_actual = ''
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS")
        self._columnas_actual = self._worm.get_column_desc()
        self._anio_actual = "0000"
        self._tabla_actual = "FACTURAS"
        self._anios = self._worm.consulta_sql_obtner_anios(self._tabla_actual)
        self._anios_frames = {}
        self._root = groot
        self._root.protocol("WM_DELETE_WINDOW", self.abortar)
        self._root.iconbitmap('sources/divar.ico')
        self._root.title("SISTEMA FACTURAS INDICE GENERAL")
        self._root.geometry("1300x600")
        self._root.resizable(False, False)
        self._columnas_ocultas = []
        self._top_levels_activos = []
        self._search_level = 1
        self._imagen = tk.PhotoImage(file="sources/warning.png")
        self._imagen_correct = tk.PhotoImage(file="sources/correct.png")
        self._imagen_fp = tk.PhotoImage(file="sources/first_p.png")
        self._imagen_add = tk.PhotoImage(file="sources/add.png")
        self._imagen_del = tk.PhotoImage(file="sources/del.png")

        self._widget_collection = list()

        # PARTE IZQUIERDA DE LA PANTALLA; OPERACIONES Y SELECCION POR AÑO:

        frame = tk.Frame(groot, bg='#387AB2', bd=0, relief=tk.RAISED)

        # Ajustar las coordenadas para colocar el rectángulo en el borde izquierdo
        # 387AB2
        portafolio_l = tk.Label(frame, text="FACTURAS DIVAR", bg='#387AB2', fg="white",
                                font=("Comic Sans MS", 10))
        portafolio_l.place(x=30, y=5)

        # Boton para cerrar el programa

        button_salir = tk.Button(frame, text="SALIR", command=self.abortar, width=20, height=1,
                                 font=("Comic Sans MS", 8, "bold"), fg='#387AB2', relief=tk.RAISED)
        button_salir.place(x=5, y=565)

        button_export = tk.Button(frame, text="EXPORTAR A EXCEL", command=self.exportar_excel, width=20, height=1,
                                  font=("Comic Sans MS", 8, "bold"), fg='#387AB2', relief=tk.RAISED)
        button_export.place(x=5, y=423)

        # Frame de botones para cambiar de tabla.
        self._frame_switch = tk.Frame(frame, width=149, height=91, bg='#387AB2')
        self.actualiza_button_switch("FACTURAS NUEVAS")
        self._frame_switch.place(x=5, y=454)

        # Inicio de la parte del scroll de las plataformas.

        frame_s = tk.Frame(frame, bg='#387AB2', width=175, height=340)
        frame_s.place(x=5, y=40)

        scrollbar = tk.Scrollbar(frame_s, orient='vertical', width=15)
        scrollbar.place(x=160, y=0, relheight=1)
        # '#387AB2'
        self.canvas = tk.Canvas(frame_s, bg='#387AB2', width=155, height=336, yscrollcommand=scrollbar.set)
        self.canvas.place(x=0, y=0)

        scrollbar.config(command=self.canvas.yview)

        self.inner_frame = tk.Frame(self.canvas, bg='#387AB2', width=155, height=336)
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor='nw')
        self.actualizar_anios()

        # Fin del area del scroll de botones.
        # Configurar el tamaño y posición del frame para que funcione como un rectángulo

        frame.place(x=0, y=0, width=185, height=600)

        # FIN PARTE IZQUIERDA DE LA PANTALLA.

        # PARTE ACTUALIZACION POR DEFECTO EN HOME

        frame_dos = tk.Frame(groot, bd=0, relief=tk.RAISED, width=1280, height=600, bg="white")
        frame_dos.place(x=187, y=0)

        # self._frame_search = tk.Frame(frame_dos, bd=0, relief=tk.RAISED, width=1097, height=20, bg="red")
        self._frame_search = tk.Frame(frame_dos, bd=0, relief=tk.RAISED, width=1050, height=20, bg="white")

        self._titulo_tabla = tk.Label(frame_dos, text="FACTURAS   INDICE   GENERAL", bg='#387AB2', fg="white",
                                      font=("Comic Sans MS", 10))
        self._titulo_tabla.place(relx=0.5, y=8, anchor='center')

        self._anio_tabla = tk.Label(frame_dos, text="0000", bg='#387AB2', fg="white", font=("Comic Sans MS", 10))

        self.search_section_update()
        self._frame_search.place(x=52, y=20)

        self.frame_botones_tabla = tk.Frame(frame_dos, bg='white', bd=0, relief=tk.RAISED, width=1100, height=46)

        no_pagadas_button = tk.Button(frame_dos, image=self._imagen_fp, command=self.act_mostrar_no_pagadas,
                                      height=17, width=17)
        no_pagadas_button.place(x=99, y=63)

        Toip.create_tooltip(no_pagadas_button, "Facturas no pagadas")

        inc_button = tk.Button(frame_dos, image=self._imagen, command=self.act_mostrar_faltantes, height=17, width=17)
        inc_button.place(x=122, y=63)

        Toip.create_tooltip(inc_button, "Facturas incompletas")

        add_button = tk.Button(frame_dos, image=self._imagen_add, command=self.increase_search_level,
                               height=15, width=17)
        add_button.place(x=5, y=19)

        Toip.create_tooltip(add_button, "Añadir filtro")

        del_button = tk.Button(frame_dos, image=self._imagen_del, command=lambda: self.decrease_search_level(),
                               height=15, width=17)
        del_button.place(x=28, y=19)

        Toip.create_tooltip(del_button, "Eliminar filtro")

        self.actualizar_bottones()
        self.frame_botones_tabla.place(x=5, y=40)

        # Frame y funciones para mostrar la tabla de datos

        self.frame_t = tk.Frame(frame_dos, width=1100, height=500, bg='red')
        self.tabular()
        self.frame_t.place(x=5, y=88)

    # METODOS PARA EL FRAME DE CAMBIO DE TABLA (INDICE-NUEVAS)

    # El método “abrir_archivo” controla la adición de facturas mediante el archivo comprimido en formato .zip dado
    # por el SAT; hace uso de los métodos definidos en la clase “BookWorm” para la descompresión del archivo, la
    # adición de los datos de los archivos a la base de datos y la consulta de la tabla de facturas. Además, se encarga
    # de tabular una vez que se han incluido las facturas y actualiza los años disponibles para su búsqueda mediante
    # la sección de botones.
    def abrir_archivo(self):
        # Abre el cuadro de diálogo de selección de archivo
        archivo = filedialog.askopenfilename(title="Selecciona un archivo",
                                             filetypes=[("Zip", "*.zip"), ("Todos los archivos", "*.*")])

        self._worm.descomprimir_zip(archivo, "Facturas Descargadas")
        self._worm.agregar_facturas_bd("Facturas Descargadas", "FACTURAS_NUEVAS", "CONCEPTOS_NUEVOS")
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS_NUEVAS")
        self.tabular()
        self.actualizar_anios()

    # El método “switch_to_index” se encarga de cambiar la consulta actual hacia la tabla de FACTURAS; básicamente
    # cambia el contexto del programa para trabajar con la tabla de FACTURAS y actualiza la interfaz para mostrar
    # los datos de dicha tabla.
    def switch_to_index(self):
        self._tabla_actual = "FACTURAS"
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS")
        self._titulo_tabla.configure(text="FACTURAS   INDICE   GENERAL")
        self._anio_actual = "0000"
        self._anio_tabla.place_forget()
        self.tabular()
        self.actualiza_button_switch("FACTURAS NUEVAS")
        self.actualizar_anios()

    # El método “switch_to_new” se encarga de cambiar la consulta actual hacia la tabla de FACTURAS_NUEVAS; básicamente
    # cambia el contexto del programa para trabajar con la tabla de FACTURAS_NUEVAS y actualiza la interfaz para mostrar
    # los datos de dicha tabla.
    def switch_to_new(self):
        self._tabla_actual = "FACTURAS_NUEVAS"
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS_NUEVAS")
        self._titulo_tabla.configure(text="FACTURAS RECIEN AGREGADAS")
        self._anio_actual = "0000"
        self._anio_tabla.place_forget()
        self.tabular()
        self.actualiza_button_switch("INDICE GENERAL")
        self.actualizar_anios()

    # Los métodos “increase_search_level” y “decrease_search_level” se encargan de
    # aumentar y disminuir (respectivamente) el criterio que determina el número de campos que se implementará
    # en las búsquedas de registros; esto es los cuadros de búsqueda. Siendo cero el mínimo y el máximo de tres
    # cuadros de búsqueda (filtros).
    def increase_search_level(self):
        if self._search_level < 3:
            self._search_level = self._search_level + 1
            self.search_section_update()

    def decrease_search_level(self):
        if self._search_level > 0:
            self._search_level = self._search_level - 1
            self.search_section_update()

    # 
    def search_section_update(self):

        for widget in self._frame_search.winfo_children():
            widget.destroy()

        self._widget_collection = list()

        cox = [0, 90, 195, 220]
        inc = [85, 105, 25, 122]

        # range(0, self._search_level)
        for x in range(0, self._search_level):
            print("X:", x)
            # Se almacena en widget label de busqueda, campo de busqueda, combo, seleccion label, coordenadas
            widget = list()

            if x != 0:
                cox[0] = cox[3] + inc[3]
                cox[1] = cox[0] + inc[0]
                cox[2] = cox[1] + inc[1]
                cox[3] = cox[2] + inc[2]
            else:
                cox = [0, 90, 195, 220]

            busqueda_label = tk.Label(self._frame_search, text="BUSQUEDA DE ", bg='#DCDAD5')
            campo_busqueda = tk.Entry(self._frame_search, bg='#DCDAD5', width=17)
            seleccion_label = tk.Label(self._frame_search, text="EN ", bg='#DCDAD5')
            combo = ttk.Combobox(self._frame_search, values=['NOMBRE XML', 'RFC', 'PROOVEDOR', 'REGIMEN FISCAL',
                                                             'FOLIO', 'SERIE', 'FORMA DE PAGO', 'TIPO DE COMPRO.',
                                                             'METODO DE PAGO', 'FECHA', 'SUBTOTAL', 'IMPUESTOS',
                                                             'TOTAL', 'TRANSFERENCIA', 'NO. DE CLIENTE',
                                                             'RECIBIO_DIVAR', 'CATEGORIA', 'CONCEPTO',
                                                             'NO.COTIZACION', 'CARPETA'],
                                 state="readonly", width=17)
            combo.set('PROOVEDOR')

            campo_busqueda.bind("<KeyRelease>", lambda event: self.search_registers(campo_busqueda.get().upper(),
                                                                                    combo.get()))

            widget.append(busqueda_label)
            widget.append(campo_busqueda)
            widget.append(seleccion_label)
            widget.append(combo)
            widget.append([cox[0], cox[1], cox[2], cox[3]])
            self._widget_collection.append(widget)

        for widget_item in self._widget_collection:
            print("Entre")
            widget_item[0].place(x=widget_item[4][0], y=0)
            widget_item[1].place(x=widget_item[4][1], y=0)
            widget_item[2].place(x=widget_item[4][2], y=0)
            widget_item[3].place(x=widget_item[4][3], y=0)

    def actualiza_button_switch(self, nombre):

        for widget in self._frame_switch.winfo_children():
            widget.destroy()

        if nombre == "FACTURAS NUEVAS":
            button_facturas = tk.Button(self._frame_switch, text=nombre, width=20, command=self.switch_to_new,
                                        height=1, font=("Comic Sans MS", 8, "bold"), fg='#387AB2', relief=tk.RAISED)

        else:
            button_facturas = tk.Button(self._frame_switch, text=nombre, width=20, command=self.switch_to_index,
                                        height=1, font=("Comic Sans MS", 8, "bold"), fg='#387AB2', relief=tk.RAISED)

            button_volcado = tk.Button(self._frame_switch, text="MOVER A INDICE", width=20,
                                       command=self.mov_reg_to_index, height=1, font=("Comic Sans MS", 8, "bold"),
                                       fg='#387AB2', relief=tk.RAISED)

            button_volcado.place(x=75, y=45, anchor="center")

            button_add_zip = tk.Button(self._frame_switch, text="AGREGAR ZIP", width=20, command=self.abrir_archivo,
                                       height=1, font=("Comic Sans MS", 8, "bold"), fg='#387AB2', relief=tk.RAISED)

            button_add_zip.place(x=75, y=77, anchor="center")

        button_facturas.place(x=75, y=13, anchor="center")

    # METODOS DE LA TABLA, RELATIVOS A ELLA Y DEL SWITCH.

    # El método “tabular” se encarga de mostrar en pantalla las consultas SQL que se realicen
    # a la base de datos, dentro del método se definen los Scrollbar para la visualización completa
    # de la consulta y se asigna a cada registro los eventos para desplegar el menú contextual para las
    # distintas opciones para cada registro.
    def tabular(self):

        # Eliminar los widgets anteriores dentro del scroll de la tabla
        for widget in self.frame_t.winfo_children():
            widget.destroy()
            del widget
        self.frame_t.winfo_children().clear()

        scrollbar_v = tk.Scrollbar(self.frame_t, orient='vertical', width=15)
        scrollbar_v.place(x=0, y=0, height=500)

        scrollbar_h = tk.Scrollbar(self.frame_t, orient='horizontal')
        scrollbar_h.place(x=15, y=0, height=15, width=1085)

        table_frame = tk.Frame(self.frame_t, width=1085, height=485, bg='green')

        stl = ttk.Style()
        stl.theme_use('clam')
        stl.configure('Treeview.Heading', background='#387AB2', foreground='white')

        headers = [["C_RFC", 110, "#1"], ["C_PROVEEDOR", 620, "#2"], ["C_REG.FISCAL", 80, "#3"], ["C_FOLIO", 250, "#4"],
                   ["C_SERIE", 100, "#5"], ["C_F.PAGO", 60, "#6"], ["C_T.COMPRO", 70, "#7"], ["C_M.PAGO", 60, "#8"],
                   ["C_FECHA", 80, "#9"], ["C_SUBTOTAL", 150, "#10"], ["C_IMPUESTOS", 150, "#11"],
                   ["C_TOTAL", 150, "#12"], ["C_TRANS./CH.", 150, "#13"], ["C_#CLIENTE", 150, "#14"],
                   ["C_RECIBIO_DIVAR", 200, "#15"], ["C_CATEGORIA", 150, "#16"], ["C_NO.COTIZACION", 150, "#17"],
                   ["C_CARPETA", 150, "#18"]]

        tree = ttk.Treeview(table_frame, height=23, columns=[header[0] for header in headers])

        tree.heading("#0", text="NOMBRE XML", anchor="w")
        if "C_XML" in self._columnas_ocultas:
            tree.column("#0", width=0, stretch=False)
        else:
            tree.column("#0", width=250, stretch=False)

        for tupla in headers:
            tree.heading(tupla[0], text=tupla[0][2:], anchor="w")
            if tupla[0] not in self._columnas_ocultas:
                tree.column(tupla[0], width=tupla[1], stretch=False)
            else:
                tree.column(tupla[0], width=0, stretch=False)

        for registro in self._consulta_actual:

            tag = "white"
            # tag = self.get_color(registro[10])
            # print(tag)
            tree.insert("", "end", text=registro[0], values=(registro[8], registro[1], registro[7],
                                                             registro[2], registro[3], registro[11], registro[10],
                                                             registro[12], registro[4][8:10] + '-' + registro[4][5:7] +
                                                             '-' + registro[4][0:4], registro[6], registro[9],
                                                             registro[5], registro[13], registro[15], registro[16],
                                                             registro[14], registro[17], registro[18]), tags=("cl"))

            tree.tag_configure("cl", background=tag)

        scrollbar_v.config(command=tree.yview)
        scrollbar_h.config(command=tree.xview)
        tree.configure(xscrollcommand=scrollbar_h.set)
        tree.configure(yscrollcommand=scrollbar_v.set)

        tree.bind("<Motion>", lambda event: self.on_mouse_hover(event, tree))
        tree.bind("<Button-3>", lambda event: self.show_context_menu(event, tree))

        tree.place(x=0, y=0, width=1085)
        table_frame.place(x=15, y=15)

    # Destinado a determinar un color para los registros del TView de acuerdo con el
    # tipo de factura; no se ha implementado
    def get_color(self, tipo_factura):
        if tipo_factura == 'P':
            bg_color = "green"
        elif tipo_factura == 'I':
            bg_color = "blue"
        elif tipo_factura == 'E':
            bg_color = "yellow"
        elif tipo_factura == 'T':
            bg_color = "pink"
        elif tipo_factura == 'N':
            bg_color = "orange"
        else:
            bg_color = "white"
        return bg_color

    # Encargado de la busqueda de los Entry de las consulatas; uno de ellos usa pandas a manera de subconsulta
    def search_registers(self, criterio, columna):

        date_flag = False
        concept_flag = False
        date_criteria = list()
        concept_criteria = list()
        consulta_tab = list()

        condiciones = list()
        for widget_set in self._widget_collection:
            condiciones.append([widget_set[3].get(), widget_set[1].get()])

        claves = [('NOMBRE XML', 'NOMBRE_XML'), ('RFC', 'RFC'), ('PROOVEDOR', 'NOMBRE'),
                  ('REGIMEN FISCAL', 'REGIMEN_FISCAL'), ('FOLIO', 'FOLIO'), ('SERIE', 'SERIE'),
                  ('FORMA DE PAGO', 'FORMA_PAGO'), ('TIPO DE COMPRO.', 'TIPO_COMPROBANTE'),
                  ('METODO DE PAGO', 'METODO_PAGO '), ('FECHA', 'FECHA'), ('SUBTOTAL', 'SUBTOTAL'),
                  ('IMPUESTOS', 'IMPUESTOS'), ('TOTAL', 'TOTAL'), ('TRANSFERENCIA', 'TRANSFERENCIA '),
                  ('NO. DE CLIENTE', 'NUMERO_CLIENTE'), ('RECIBIO_DIVAR', 'RECIBIO_DIVAR'),
                  ('CATEGORIA', 'CATEGORIA_FACTURA'), ('NO_COTIZACION', 'NO.COTIZACION'), ('CARPETA', 'CARPETA')]

        columnas = ['NOMBRE_XML', 'NOMBRE', 'FOLIO', 'SERIE', 'FECHA',  'TOTAL', 'SUBTOTAL', 'REGIMEN_FISCAL', 'RFC',
                    'IMPUESTOS', 'TIPO_COMPROBANTE', 'FORMA_PAGO', 'METODO_PAGO ', 'TRANSFERENCIA ',
                    'CATEGORIA_FACTURA', 'NUMERO_CLIENTE', 'RECIBIO_DIVAR', 'NO.COTIZACION', 'CARPETA']

        for clave in claves:
            for cond in condiciones:
                if cond[0] == clave[0]:
                    cond[0] = clave[1]

        for cond in condiciones.copy():

            if cond[0] == "CONCEPTO":
                concept_flag = True
                index_concept = condiciones.index(cond)
                concept_criteria = cond[1]
                condiciones.pop(index_concept)

            if cond[0] == "FECHA":
                date_flag = True
                index_date = condiciones.index(cond)
                date_criteria = cond[1]
                condiciones.pop(index_date)

        if date_flag is True and concept_flag is True:
            if len(date_criteria) == 10:
                date_criteria = (str(date_criteria[6:10]) + '-' + str(date_criteria[3:5]) + '-' +
                                 str(date_criteria[0:2]))

        if concept_flag:
            if not isinstance(date_criteria, list):
                condiciones.append(["FECHA", date_criteria])

            if self._tabla_actual == "FACTURAS":
                consulta_tab = self._worm.consulta_sql_concepto("FACTURAS", "CONCEPTOS", concept_criteria)
            else:
                consulta_tab = self._worm.consulta_sql_concepto("FACTURAS_NUEVAS", "CONCEPTOS_NUEVOS", concept_criteria)

        elif date_flag:
            if self._tabla_actual == "FACTURAS":
                consulta_tab = self._worm.consulta_sql_rango_fecha("FACTURAS", date_criteria, self._anio_actual)

            else:
                consulta_tab = self._worm.consulta_sql_rango_fecha("FACTURAS_NUEVAS", date_criteria, self._anio_actual)

        if condiciones:
            if date_flag or concept_flag:
                consulta_tab = self._worm.consulta_panda(consulta_tab, columnas, condiciones)
            else:
                if self._anio_actual == "0000":
                    consulta_tab = self._worm.consulta_sql(self._tabla_actual, condiciones)
                    # print(self._consulta_actual)
                else:
                    consulta_tab = self._worm.consulta_sql(self._tabla_actual, condiciones, self._anio_actual)
        self._consulta_actual = consulta_tab
        self.tabular()

    # Encargado de identificar el objeto sobre el cual se posiciona el cursor; es usado para identificar los registros
    # en el TView.
    @ staticmethod
    def on_mouse_hover(event, tree):
        item = tree.identify_row(event.y)
        tree.selection_set(item)

    # Muestra el menú contextual de un registro del TView al hacer clic derecho en él (conceptos de factura).
    def show_context_menu_concept(self, event, tree):
        # Obtener el item en el cual se hizo click derecho
        item_id = tree.identify('item', event.x, event.y)
        # print(tree.item(item_id, 'text'))
        item_values = tree.item(item_id, 'values')
        # print(item_values)
        # Creacion y despliegue del menu contextual
        c_menu = tk.Menu(self.frame_t, tearoff=0)
        c_menu.add_command(label="Editar campos",
                           command=lambda: self.pop_editar_concepto(tree.item(item_id, 'text'), item_values[-1]))
        c_menu.post(event.x_root, event.y_root)

    # Muestra el menú contextual de un registro del TView al hacer clic derecho en él (factura).
    def show_context_menu(self, event, tree):
        # Obtener el item en el cual se hizo click derecho
        item_id = tree.identify('item', event.x, event.y)
        item_values = tree.item(item_id, 'values')
        # Guardado auxiliar de la factura actual
        self._factura_actual = tree.item(item_id, 'text')
        # Creacion y despliegue del menu contextual
        c_menu = tk.Menu(self.frame_t, tearoff=0)
        # print("valores", item_values)
        # print("no_se que onda:", *item_values[12:18])
        c_menu.add_command(label="Editar campos",
                           command=lambda: self.pop_editar_registro(tree.item(item_id, 'text'), *item_values[12:18]))
        c_menu.add_command(label="Mostrar concepto",
                           command=lambda: self.pop_concepto(tree.item(item_id, 'text')))
        c_menu.post(event.x_root, event.y_root)

    # Encargado de gestionar los botones que muestran/ocultan campos del TView de facturas.
    def actualizar_bottones(self):

        # Eliminar los widgets anteriores dentro del frame de botones
        for widget in self.frame_botones_tabla.winfo_children():
            widget.destroy()

        lista_de_botones = list()

        columnas = [["C_XML", "XML", 4, (0, 0)],
                    ["C_RFC", "RFC", 4, (34, 0)],
                    ["C_PROVEEDOR", "PROVEEDOR", 10, (68, 0)],
                    ["C_REG.FISCAL", "REG.FIS.", 8, (138, 0)],
                    ["C_FOLIO", "FOLIO", 6, (196, 0)],
                    ["C_SERIE", "SERIE", 5, (242, 0)],
                    ["C_F.PAGO", "F.PAGO", 7, (282, 0)],
                    ["C_T.COMPRO", "T.COMPR.", 9, (334, 0)],
                    ["C_M.PAGO", "M.PAGO", 8, (398, 0)],
                    ["C_FECHA", "FECHA", 7, (456, 0)],
                    ["C_SUBTOTAL", "SUBTOTAL", 10, (508, 0)],
                    ["C_IMPUESTOS", "IMPUESTOS", 10, (578, 0)],
                    ["C_TOTAL", "TOTAL", 6, (648, 0)],
                    ["C_TRANS./CH.", "TRANSFERENCIA", 14, (694, 0)],
                    ["C_#CLIENTE", "#CLIENTE", 9, (788, 0)],
                    ["C_RECIBIO_DIVAR", "RECIBIO_DIVAR", 14, (852, 0)],
                    ["C_CATEGORIA", "CATEGORIA", 12, (946, 0)],
                    ["C_CARPETA", "CARPETA", 10, (1028, 0)],
                    ["C_NO.COTIZACION", "NO.COTIZACION", 14, (0, 23)]]

        for columna in columnas:
            # print(columna[0])
            if columna[0] not in self._columnas_ocultas:
                button = tk.Button(self.frame_botones_tabla, text=columna[1],
                                   command=lambda col=columna[0]: self.ocultar_campo(col), width=columna[2], height=1,
                                   font=("Comic Sans MS", 7, "bold"), fg='#387AB2', relief=tk.RAISED)
            else:
                button = tk.Button(self.frame_botones_tabla, text=columna[1],
                                   command=lambda col=columna[0]: self.mostrar_campo(col), width=columna[2],
                                   bg='#646464', height=1, font=("Comic Sans MS", 7, "bold"), fg='white',
                                   relief=tk.RAISED)
            lista_de_botones.append(button)

            button.place(x=columna[3][0], y=columna[3][1])
            # button.pack(side='left')

    # Llamado para ocultar un campo del TView de facturas.
    def ocultar_campo(self, campo):
        self._columnas_ocultas.append(campo)
        # print(self._columnas_ocultas)
        self.actualizar_bottones()
        self.tabular()

    # Llamado para mostrar un campo del TView de facturas.
    def mostrar_campo(self, campo):
        self._columnas_ocultas.pop(self._columnas_ocultas.index(campo))
        self.actualizar_bottones()
        self.tabular()

    # Encargado de actualizar la tabla de facturas de acuerdo con el año seleccionado.
    def act_anios(self, anio, tabla):
        self._consulta_actual = self._worm.consulta_sql_por_anio(anio, tabla)
        self._anio_actual = str(anio)
        self._anio_tabla.configure(text=str(anio))
        self._anio_tabla.place(relx=0.5, y=25, anchor='center')
        self.tabular()

    # Encargado de actualizar la tabla de facturas aun estado predeterminado.
    def act_gen(self, tabla):
        self._consulta_actual = self._worm.consulta_sql_completa(tabla)
        self._anio_actual = "0000"
        self._anio_tabla.place_forget()
        self.tabular()

    # Encargado de actualizar los botones de los años disponibles.
    def actualizar_anios(self):

        for widget in self.inner_frame.winfo_children():
            widget.destroy()

        self._anios = self._worm.consulta_sql_obtner_anios(self._tabla_actual)

        button = tk.Button(self.inner_frame, text="GENERAL", width=18, height=1,
                           font=("Comic Sans MS", 10, "bold"), fg='#387AB2', relief=tk.RAISED,
                           command=lambda: self.act_gen(self._tabla_actual))
        button.pack(pady=5, padx=1, anchor="center")

        for anio in self._anios:
            button = tk.Button(self.inner_frame, text=anio, width=18, height=1, font=("Comic Sans MS", 10, "bold"),
                               fg='#387AB2', relief=tk.RAISED,
                               command=lambda an=anio: self.act_anios(an, self._tabla_actual))
            button.pack(pady=5, padx=1, anchor="center")

            frame_anual = tk.Frame(self.inner_frame, bg='#387AB2', width=175, height=340)

            for mes in ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre",
                        "Octubre", "Noviembre", "Diciembre"]:

                button = tk.Button(frame_anual, text=mes, width=18, height=1, font=("Comic Sans MS", 10, "bold"),
                                   fg='#387AB2', relief=tk.RAISED)
                button.pack(pady=5, padx=1, anchor="center")

            self._anios_frames[anio] = frame_anual

        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    # Es la ventana pop que muestra los conceptos de la factura.
    def pop_concepto(self, value):

        if self._tabla_actual == "FACTURAS":

            consulta = self._worm.consulta_sql_por_campo("CONCEPTOS", "NOMBRE_XML", value)

        else:

            consulta = self._worm.consulta_sql_por_campo("CONCEPTOS_NUEVOS", "NOMBRE_XML", value)

        for pop_up in self._top_levels_activos:
            pop_up.destroy()

        pop = tk.Toplevel()
        self._top_levels_activos.append(pop)
        pop.title("CONCEPTOS DE {0}".format(value.upper()))
        pop.geometry("900x210")
        pop.resizable(False, False)

        label_head = tk.Label(pop, text="CONCEPTOS", fg='#387AB2', font=("Comic Sans MS", 12))
        label_head.place(relx=0.5, y=20, anchor="center")

        # F0F0F0
        general_frame = tk.Frame(pop, width=900, height=180, bg="#F0F0F0")
        scrollbar_v = tk.Scrollbar(general_frame, orient='vertical', width=15)

        canvas = tk.Canvas(general_frame, bg="#F0F0F0", width=840, height=170, yscrollcommand=scrollbar_v.set)

        inner_frame = tk.Frame(canvas, width=840, height=171, bg='#F0F0F0')

        stl = ttk.Style()
        stl.theme_use('clam')

        stl.configure('Treeview.Heading', background='#387AB2', foreground='white')

        headers = [["C_CLAVE_PRO/SER", 100], ["C_DESCRIPCION", 250], ["C_CANTIDAD", 80], ["C_VALOR_U.", 100],
                   ["C_CON_CLIENTE", 200]]

        tree = ttk.Treeview(inner_frame, height=7, columns=[header[0] for header in headers])

        tree.heading("#0", text="ID", anchor="w")
        tree.column("#0", width=100)

        for tupla in [header for header in headers]:
            tree.heading(tupla[0], text=tupla[0][2:], anchor="w")
            tree.column(tupla[0], minwidth=tupla[1], width=tupla[1])

        # for registro in self._worm._cursor.fetchall():
        for registro in consulta:
            tree.insert("", "end", text=registro[0], values=(registro[2], registro[3], registro[4],
                                                             registro[5], registro[6]))

        scrollbar_v.config(command=tree.yview)
        tree.configure(yscrollcommand=scrollbar_v.set)
        tree.place(relx=0.5, rely=0.5, anchor="center")
        # tree.place(x=0, y=0)
        inner_frame.place(x=0, y=0)
        scrollbar_v.place(x=10, y=10, height=150)
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.place(x=30, y=0)
        general_frame.place(x=0, y=30)

        tree.bind("<Motion>", lambda event: self.on_mouse_hover(event, tree))
        tree.bind("<Button-3>", lambda event: self.show_context_menu_concept(event, tree))

        pop.mainloop()

    # Pop que solicita una contraseña antes de acceder a una sección del programa (pop).
    def authenticate(self):

        pop_pass = tk.Toplevel()
        self._top_levels_activos.append(pop_pass)
        pop_pass.title("AUTENTICACIÓN")
        pop_pass.geometry("290x130")
        pop_pass.resizable(False, False)

        variable = tk.StringVar()
        condicion = False
        contador = 0

        label_cont = tk.Label(pop_pass, text="", fg='red', font=("Comic Sans MS", 10))
        label_cont.place(x=20, y=65)

        label_head = tk.Label(pop_pass, text="CONTRASEÑA", fg='#387AB2', font=("Comic Sans MS", 15))
        label_head.place(relx=0.5, y=20, anchor="center")

        campo_busqueda = AutocompleteCombobox(pop_pass, width=24, textvariable=variable, font=('Comic Sans MS', 12),
                                              show='*')
        campo_busqueda.place(x=20, y=40)

        def authentication():
            nonlocal pop_pass
            nonlocal condicion
            nonlocal contador
            nonlocal label_cont
            hash_obj = hashlib.sha256(variable.get().encode())
            condicion = self._worm.authenticate(hash_obj.hexdigest())
            if condicion is True:
                pop_pass.destroy()
            else:
                contador = contador + 1
                if contador >= 3:
                    pop_pass.destroy()
                else:
                    label_cont.configure(text="Intento {0} de 3".format(contador))

        button_aceptar = tk.Button(pop_pass, text="ACEPTAR",
                                   command=authentication,
                                   width=15, height=1, font=("Comic Sans MS", 10, "bold"),
                                   fg='#208000', relief=tk.RAISED)
        button_aceptar.place(x=80, y=100, anchor="center")

        button_cancelar = tk.Button(pop_pass, text="CANCELAR", command=self.act_regresar_uno, width=15,
                                    height=1, font=("Comic Sans MS", 10, "bold"), fg='#A21B00', relief=tk.RAISED)
        button_cancelar.place(x=215, y=100, anchor="center")
        pop_pass.wait_window()
        return condicion

    # Pop que sirve para editar un concepto de una factura.
    def pop_editar_concepto(self, value, c_cliente):

        if not self.authenticate():
            return

        entradas = []
        variables = []

        pop = tk.Toplevel()
        self._top_levels_activos.append(pop)
        pop.title("EDITAR CONCEPTO NO. {0}".format(value))
        pop.geometry("500x130")
        pop.resizable(False, False)

        coy = 35

        for parametro in [c_cliente]:
            variable = tk.StringVar()
            variable.set(parametro)
            variables.append(variable)

        label_head = tk.Label(pop, text="MODIFICACIÓN DE CAMPOS", fg='#387AB2', font=("Comic Sans MS", 20))
        label_head.place(relx=0.5, y=20, anchor="center")

        for campo in [('CLIENTE:', 0)]:

            label = tk.Label(pop, text=campo[0], fg='#387AB2', font=("Comic Sans MS", 12))
            label.place(x=20, y=coy+7)

            opciones = list()

            for opcion in self._worm.consulta_sql_completa('CLIENTES'):
                opciones.append(opcion[0])

            campo_busqueda = AutocompleteCombobox(pop, width=37, textvariable=variables[campo[1]],
                                                  font=('Comic Sans MS', 12), completevalues=opciones)

            campo_busqueda.place(x=100, y=coy + 7)
            coy = coy + 30
            entradas.append(campo_busqueda)

        button_aceptar = tk.Button(pop, text="ACEPTAR",
                                   command=lambda: self.act_aceptar_ec(*(var.get() for var in entradas), value),
                                   width=27, height=1, font=("Comic Sans MS", 10, "bold"),
                                   fg='#208000', relief=tk.RAISED)
        button_aceptar.place(x=130, y=100, anchor="center")

        button_cancelar = tk.Button(pop, text="CANCELAR", command=self.act_regresar_uno, width=27,
                                    height=1, font=("Comic Sans MS", 10, "bold"), fg='#A21B00', relief=tk.RAISED)
        button_cancelar.place(x=360, y=100, anchor="center")

        pop.mainloop()

    # Pop que sirve para editar un registro (factura).
    def pop_editar_registro(self, value, c_transferencia, c_cliente, c_divar, c_categoria, c_no_cotizacion, c_carpeta):

        if not self.authenticate():
            return

        entradas = []
        variables = []

        for pop_up in self._top_levels_activos:
            pop_up.destroy()

        pop = tk.Toplevel()
        self._top_levels_activos.append(pop)
        pop.title("EDITAR {0}".format(value.upper()))
        pop.geometry("680x270")
        pop.resizable(False, False)

        coy = 35

        for parametro in [c_transferencia, c_cliente, c_divar, c_categoria, c_no_cotizacion, c_carpeta]:
            variable = tk.StringVar()
            variable.set(parametro)
            variables.append(variable)

        label_head = tk.Label(pop, text="MODIFICACIÓN DE CAMPOS", fg='#387AB2', font=("Comic Sans MS", 20))
        label_head.place(relx=0.5, y=20, anchor="center")

        for campo in [('TRANSFERENCIA:', 0), ('#CLIENTE:', 1), ('RECIBIO_DIVAR:', 2), ('CATEGORIA:', 3),
                      ('NO.COTIZACIÓN:', 4), ('CARPETA', 5)]:

            label = tk.Label(pop, text=campo[0], fg='#387AB2', font=("Comic Sans MS", 12))
            label.place(x=20, y=coy)
            if campo[0] in ['TRANSFERENCIA:', 'NO.COTIZACIÓN:']:
                campo_busqueda = tk.Entry(pop, bg='#DCDAD5', textvariable=variables[campo[1]], width=80)
            else:
                opciones = list()
                if campo[0] == '#CLIENTE:':
                    for opcion in self._worm.consulta_sql_completa('CLIENTES'):
                        opciones.append(opcion[0])
                elif campo[0] == 'CATEGORIA:':
                    for opcion in self._worm.consulta_sql_completa('CATEGORIA'):
                        opciones.append(opcion[0])
                elif campo[0] == 'RECIBIO_DIVAR:':
                    for opcion in self._worm.consulta_sql_completa('PERSONAL'):
                        opciones.append(opcion[0])
                elif campo[0] == 'CARPETA':
                    opciones = ["ENERO", "FEBRERO", "MARZO", "JUNIO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
                                "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

                campo_busqueda = AutocompleteCombobox(pop, textvariable=variables[campo[1]], width=77,
                                                      completevalues=opciones)
                campo_busqueda.set(variables[campo[1]].get())
            campo_busqueda.place(x=180, y=coy + 7)
            coy = coy + 30
            entradas.append(campo_busqueda)

        button_aceptar = tk.Button(pop, text="ACEPTAR",
                                   command=lambda: self.act_aceptar_er(*(var.get() for var in entradas), value),
                                   width=38, height=1, font=("Comic Sans MS", 10, "bold"),
                                   fg='#208000', relief=tk.RAISED)
        button_aceptar.place(x=180, y=240, anchor="center")

        button_cancelar = tk.Button(pop, text="CANCELAR", command=self.act_regresar, width=38,
                                    height=1, font=("Comic Sans MS", 10, "bold"), fg='#A21B00', relief=tk.RAISED)
        button_cancelar.place(x=505, y=240, anchor="center")

        pop.mainloop()

    # Pop que sirve para indicar que una modificacion a los registros ha sido exitosa; se cierra automaticamente.
    def pop_modificaion_exitosa(self):
        pop = tk.Toplevel()
        self._top_levels_activos.append(pop)
        pop.title("MODIFICACIÓN EXISTOSA")
        pop.geometry("380x170")
        pop.resizable(False, False)
        label_image = tk.Label(pop, image=self._imagen_correct, text="________________", fg='#208000')
        label_image.place(relx=0.5, y=70, anchor="center")
        label_nombre_perfil_a = tk.Label(pop, text="LA MODIFICACIÓN SE HA HECHO CORRECTAMENTE", bg='#F0F0F0',
                                         fg='#208000', font=("Comic Sans MS", 10))
        label_nombre_perfil_a.place(relx=0.5, y=150, anchor="center")
        # pop.after(2000, pop.destroy)
        pop.after(2000, self.act_regresar)

    # Pop que sirve para indicar que una modificacion a los conceptos ha sido exitosa; se cierra automaticamente.
    def pop_modificaion_exitosa_conceptos(self):
        pop = tk.Toplevel()
        # self._top_levels_activos.append(pop)
        pop.title("MODIFICACIÓN EXISTOSA")
        pop.geometry("380x170")
        pop.resizable(False, False)
        label_image = tk.Label(pop, image=self._imagen_correct, text="________________", fg='#208000')
        label_image.place(relx=0.5, y=70, anchor="center")
        label_nombre_perfil_a = tk.Label(pop, text="LA MODIFICACIÓN SE HA HECHO CORRECTAMENTE", bg='#F0F0F0',
                                         fg='#208000', font=("Comic Sans MS", 10))
        label_nombre_perfil_a.place(relx=0.5, y=150, anchor="center")
        pop.after(2000, pop.destroy)
        for pop_up in self._top_levels_activos:
            pop_up.destroy()
        self._top_levels_activos = list()
        self.pop_concepto(self._factura_actual)

    # Destruye todos los pop activos
    def act_regresar(self):
        for pop_up in self._top_levels_activos:
            pop_up.destroy()
            del pop_up

    # Destruye el ultimo pop activo (en ultimo en pila; el del tope)
    def act_regresar_uno(self):
        self._top_levels_activos[-1].destroy()

    # Hace una consulta y tabula las facturas a las cuales les falta uno de los tres campos añadidos por la empresa.
    def act_mostrar_faltantes(self):
        # print(self._anio_actual)
        self._consulta_actual = self._worm.consulta_sql_faltantes(self._tabla_actual, self._anio_actual)
        self.tabular()

    # Hace una consulta y tabula las facturas que no se han pagado.
    def act_mostrar_no_pagadas(self):
        self._consulta_actual = self._worm.consulta_sql_no_pagadas(self._tabla_actual, self._anio_actual)
        self.tabular()

    # Da de alta un registro (factura).
    def act_aceptar_er(self, trans, numero_cliente, r_divar, categoria, n_cot, carpeta, nombre_xml):

        self._worm.agregar_a_reg_tab_simple("CLIENTES", numero_cliente.upper())
        self._worm.agregar_a_reg_tab_simple("CATEGORIA", categoria.upper())
        self._worm.agregar_a_reg_tab_simple("PERSONAL", r_divar.upper())
        self._worm.consulta_sql_actualizar_registro(trans, numero_cliente, r_divar, categoria, n_cot, carpeta,
                                                    nombre_xml, self._tabla_actual)
        self.pop_modificaion_exitosa()
        self._consulta_actual = self._worm.consulta_sql_completa(self._tabla_actual)
        self.tabular()

    # Asociado a la modificación de conceptos; agrega un cliente si es necesario.
    def act_aceptar_ec(self, cliente, identificador):

        self._worm.agregar_a_reg_tab_simple("CLIENTES", cliente.upper())
        self._worm.consulta_sql_actualizar_concepto(identificador, cliente, self._tabla_actual)
        self.pop_modificaion_exitosa_conceptos()
        # self._consulta_actual = self._worm.consulta_sql_completa(self._tabla_actual)
        # self.tabular()

    # Mueve las facturas de la tabla temporal a la tabla index de facturas.
    def mov_reg_to_index(self):
        self._worm.consulta_sql_mover_registros_index()
        self._worm.clasificar_facturas_archivos("Facturas Descargadas", "Facturas Clasificadas")
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS_NUEVAS")
        self._columnas_actual = self._worm.get_column_desc()
        self.tabular()
        self.actualizar_anios()

    # Se encarga de hacer la exoprtación a archivo Excel.
    def exportar_excel(self):

        dicc = {"C_XML": "NOMBRE_XML", "C_PROVEEDOR": "NOMBRE", "C_FOLIO": "FOLIO", "C_SERIE": "SERIE",
                "C_FECHA": "FECHA", "C_TOTAL": "TOTAL", "C_SUBTOTAL": "SUBTOTAL", "C_REG.FISCAL": "REGIMEN_FISCAL",
                "C_RFC": "RFC", "C_IMPUESTOS": "IMPUESTOS", "C_T.COMPRO": "TIPO_COMPROBANTE", "C_F.PAGO": "FORMA_PAGO",
                "C_M.PAGO": "METODO_PAGO", "C_TRANS./CH.": "TRANSFERENCIA", "C_CATEGORIA": "CATEGORIA_FACTURA",
                "C_#CLIENTE": "NUMERO_CLIENTE", "C_RECIBIO_DIVAR": "RECIBIO_DIVAR", "C_CARPETA": "CARPETA",
                "C_NO.COTIZACION": "NO_COTIZACION"}

        if self._columnas_ocultas:
            columnas = list()
            for columna in self._columnas_ocultas:
                columnas.append(dicc[columna])
            # print("C_OCULTAS: ", columnas)
            self._worm.exportar_xlsx(self._consulta_actual, self._columnas_actual, "facturas", columnas)
        else:
            self._worm.exportar_xlsx(self._consulta_actual, self._columnas_actual, "facturas")

        consulta_de_conceptos = list()

        for tupla in self._consulta_actual:

            if self._tabla_actual == "FACTURAS":
                aux = self._worm.consulta_sql_concepto_por_xml("CONCEPTOS", tupla[0])
            else:
                aux = self._worm.consulta_sql_concepto_por_xml("CONCEPTOS_NUEVOS", tupla[0])

            for tupla_c in aux:
                consulta_de_conceptos.append(tupla_c)

        columnas_con = ["ID", "NOMBRE XML", "CLAVE PROD/SERV", "DESCRIPCION", "CANTIDAD", "VALOR UNITARIO", "CLIENTE"]
        self._worm.exportar_xlsx(consulta_de_conceptos, columnas_con, "conceptos_asociados")

    # Usado para cerrar los widgets, cerrar la conexión con la base de datos y el programa.
    def abortar(self):
        for pop_up in self._top_levels_activos:
            pop_up.destroy()
        self._root.destroy()
        self._worm.cerrar_conexion()
