import tkinter as tk
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
        self._root = groot
        self._root.protocol("WM_DELETE_WINDOW", self.abortar)
        self._root.iconbitmap('sources/divar.ico')
        # self._divar_icon = tk.PhotoImage('sources/divar.png')
        # self._root.iconphoto(True, self._divar_icon)
        self._root.title("SISTEMA FACTURAS INDICE GENERAL")
        self._root.geometry("1300x600")
        self._root.resizable(False, False)
        self._columnas_ocultas = []
        self._top_levels_activos = []

        self._imagen = tk.PhotoImage(file="sources/warning.png")
        self._imagen_correct = tk.PhotoImage(file="sources/correct.png")
        self._imagen_fp = tk.PhotoImage(file="sources/first_p.png")

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

        self._titulo_tabla = tk.Label(frame_dos, text="FACTURAS   INDICE   GENERAL", bg='#387AB2', fg="white",
                                      font=("Comic Sans MS", 10))
        self._titulo_tabla.place(relx=0.5, y=8, anchor='center')

        self._anio_tabla = tk.Label(frame_dos, text="0000", bg='#387AB2', fg="white", font=("Comic Sans MS", 10))

        busqueda_label = tk.Label(frame_dos, text="BUSQUEDA DE ", bg='#DCDAD5')
        busqueda_label.place(x=5, y=20)
        campo_busqueda = tk.Entry(frame_dos, bg='#DCDAD5', width=24)
        combo = ttk.Combobox(frame_dos, values=['NOMBRE XML', 'RFC', 'PROOVEDOR', 'REGIMEN FISCAL', 'FOLIO', 'SERIE',
                                                'FORMA DE PAGO', 'TIPO DE COMPRO.', 'METODO DE PAGO', 'FECHA',
                                                'SUBTOTAL', 'IMPUESTOS', 'TOTAL', 'TRANSFERENCIA', 'NO. DE CLIENTE',
                                                'RECIBIO_DIVAR', 'CATEGORIA', 'CONCEPTO', 'NO.COTIZACION', 'CARPETA'],
                             state="readonly")
        combo.set('PROOVEDOR')
        combo.place(x=261, y=20)
        campo_busqueda.bind("<KeyRelease>", lambda event: self.search_registers(campo_busqueda.get().upper(),
                                                                                combo.get()))
        campo_busqueda.place(x=90, y=21)
        seleccion_label = tk.Label(frame_dos, text="EN ", bg='#DCDAD5')
        seleccion_label.place(x=237, y=20)

        self.frame_botones_tabla = tk.Frame(frame_dos, bg='white', bd=0, relief=tk.RAISED, width=1100, height=46)

        no_pagadas_button = tk.Button(frame_dos, image=self._imagen_fp, command=self.act_mostrar_no_pagadas,
                                      height=17, width=17)
        no_pagadas_button.place(x=99, y=63)

        inc_button = tk.Button(frame_dos, image=self._imagen, command=self.act_mostrar_faltantes, height=17, width=17)
        inc_button.place(x=122, y=63)

        self.actualizar_bottones()
        self.frame_botones_tabla.place(x=5, y=40)

        # Frame y funciones para mostrar la tabla de datos

        self.frame_t = tk.Frame(frame_dos, width=1100, height=500, bg='red')
        self.tabular()
        self.frame_t.place(x=5, y=88)

    # METODOS PARA EL FRAME DE CAMBIO DE TABLA (INDICE-NUEVAS)
    def abrir_archivo(self):
        # Abre el cuadro de diálogo de selección de archivo
        archivo = filedialog.askopenfilename(title="Selecciona un archivo",
                                             filetypes=[("Zip", "*.zip"), ("Todos los archivos", "*.*")])

        self._worm.descomprimir_zip(archivo, "Facturas Descargadas")
        self._worm.agregar_facturas_bd("Facturas Descargadas", "FACTURAS_NUEVAS", "CONCEPTOS_NUEVOS")
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS_NUEVAS")
        self.tabular()
        self.actualizar_anios()

    def switch_to_index(self):
        self._tabla_actual = "FACTURAS"
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS")
        self._titulo_tabla.configure(text="FACTURAS   INDICE   GENERAL")
        self._anio_actual = "0000"
        self._anio_tabla.place_forget()
        self.tabular()
        self.actualiza_button_switch("FACTURAS NUEVAS")
        self.actualizar_anios()

    def switch_to_new(self):
        self._tabla_actual = "FACTURAS_NUEVAS"
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS_NUEVAS")
        self._titulo_tabla.configure(text="FACTURAS RECIEN AGREGADAS")
        self._anio_actual = "0000"
        self._anio_tabla.place_forget()
        self.tabular()
        self.actualiza_button_switch("INDICE GENERAL")
        self.actualizar_anios()

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

    def tabular(self):

        # Eliminar los widgets anteriores dentro del scroll de la tabla
        for widget in self.frame_t.winfo_children():
            widget.destroy()

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
            tree.insert("", "end", text=registro[0], values=(registro[8], registro[1], registro[7],
                                                             registro[2], registro[3], registro[11], registro[10],
                                                             registro[12], registro[4][8:10] + '-' + registro[4][5:7] +
                                                             '-' + registro[4][0:4], registro[6], registro[9],
                                                             registro[5], registro[13], registro[15], registro[16],
                                                             registro[14], registro[17], registro[18]))

        scrollbar_v.config(command=tree.yview)
        scrollbar_h.config(command=tree.xview)
        tree.configure(xscrollcommand=scrollbar_h.set)
        tree.configure(yscrollcommand=scrollbar_v.set)

        tree.bind("<Motion>", lambda event: self.on_mouse_hover(event, tree))
        tree.bind("<Button-3>", lambda event: self.show_context_menu(event, tree))

        tree.place(x=0, y=0, width=1085)
        table_frame.place(x=15, y=15)

    def search_registers(self, criterio, columna):

        claves = [('NOMBRE XML', 'NOMBRE_XML'), ('RFC', 'RFC'), ('PROOVEDOR', 'NOMBRE'),
                  ('REGIMEN FISCAL', 'REGIMEN_FISCAL'), ('FOLIO', 'FOLIO'), ('SERIE', 'SERIE'),
                  ('FORMA DE PAGO', 'FORMA_PAGO'), ('TIPO DE COMPRO.', 'TIPO_COMPROBANTE'),
                  ('METODO DE PAGO', 'METODO_PAGO '), ('FECHA', 'FECHA'), ('SUBTOTAL', 'SUBTOTAL'),
                  ('IMPUESTOS', 'IMPUESTOS'), ('TOTAL', 'TOTAL'), ('TRANSFERENCIA', 'TRANSFERENCIA '),
                  ('NO. DE CLIENTE', 'NUMERO_CLIENTE'), ('RECIBIO_DIVAR', 'RECIBIO_DIVAR'),
                  ('CATEGORIA', 'CATEGORIA_FACTURA'), ('NO_COTIZACION', 'NO.COTIZACION'), ('CARPETA', 'CARPETA')]

        columna_sql = 'NOMBRE'

        if columna == "FECHA":
            if self._tabla_actual == "FACTURAS":
                self._consulta_actual = self._worm.consulta_sql_rango_fecha("FACTURAS", criterio, self._anio_actual)
                # print(self._consulta_actual)
            else:
                self._consulta_actual = self._worm.consulta_sql_rango_fecha("FACTURAS_NUEVAS", criterio,
                                                                            self._anio_actual)
            self.tabular()
            return

        if columna == "CONCEPTO":
            if self._tabla_actual == "FACTURAS":
                self._consulta_actual = self._worm.consulta_sql_concepto("FACTURAS", "CONCEPTOS", criterio)
            else:
                self._consulta_actual = self._worm.consulta_sql_concepto("FACTURAS_NUEVAS", "CONCEPTOS_NUEVOS",
                                                                         criterio)

            # self._consulta_actual = self._worm._cursor.execute(consulta, ['%' + criterio + '%'])
            self.tabular()
            return

        for clave in claves:
            if columna == clave[0]:
                columna_sql = clave[1]
                break

        if self._anio_actual == "0000":
            self._consulta_actual = self._worm.consulta_sql(self._tabla_actual, columna_sql, criterio)
        else:
            self._consulta_actual = self._worm.consulta_sql(self._tabla_actual, columna_sql, criterio,
                                                            self._anio_actual)
        self.tabular()

    @ staticmethod
    def on_mouse_hover(event, tree):
        item = tree.identify_row(event.y)
        tree.selection_set(item)

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

    def ocultar_campo(self, campo):
        self._columnas_ocultas.append(campo)
        # print(self._columnas_ocultas)
        self.actualizar_bottones()
        self.tabular()

    def mostrar_campo(self, campo):
        self._columnas_ocultas.pop(self._columnas_ocultas.index(campo))
        self.actualizar_bottones()
        self.tabular()

    def act_anios(self, anio, tabla):
        self._consulta_actual = self._worm.consulta_sql_por_anio(anio, tabla)
        self._anio_actual = str(anio)
        self._anio_tabla.configure(text=str(anio))
        self._anio_tabla.place(relx=0.5, y=25, anchor='center')
        self.tabular()

    def act_gen(self, tabla):
        self._consulta_actual = self._worm.consulta_sql_completa(tabla)
        self._anio_actual = "0000"
        self._anio_tabla.place_forget()
        self.tabular()

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

        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

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

        # AQUÍ ME QUEDE....................................................

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

    def pop_editar_concepto(self, value, c_cliente):

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

    def pop_editar_registro(self, value, c_transferencia, c_cliente, c_divar, c_categoria, c_no_cotizacion, c_carpeta):

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

    def act_regresar(self):
        for pop_up in self._top_levels_activos:
            pop_up.destroy()

    def act_regresar_uno(self):
        self._top_levels_activos[-1].destroy()

    def act_mostrar_faltantes(self):
        # print(self._anio_actual)
        self._consulta_actual = self._worm.consulta_sql_faltantes(self._tabla_actual, self._anio_actual)
        self.tabular()

    def act_mostrar_no_pagadas(self):
        self._consulta_actual = self._worm.consulta_sql_no_pagadas(self._tabla_actual, self._anio_actual)
        self.tabular()

    def act_aceptar_er(self, trans, numero_cliente, r_divar, categoria, n_cot, carpeta, nombre_xml):

        self._worm.agregar_a_reg_tab_simple("CLIENTES", numero_cliente.upper())
        self._worm.agregar_a_reg_tab_simple("CATEGORIA", categoria.upper())
        self._worm.agregar_a_reg_tab_simple("PERSONAL", r_divar.upper())
        self._worm.consulta_sql_actualizar_registro(trans, numero_cliente, r_divar, categoria, n_cot, carpeta,
                                                    nombre_xml, self._tabla_actual)
        self.pop_modificaion_exitosa()
        self._consulta_actual = self._worm.consulta_sql_completa(self._tabla_actual)
        self.tabular()

    def act_aceptar_ec(self, cliente, identificador):

        self._worm.agregar_a_reg_tab_simple("CLIENTES", cliente.upper())
        self._worm.consulta_sql_actualizar_concepto(identificador, cliente, self._tabla_actual)
        self.pop_modificaion_exitosa_conceptos()
        # self._consulta_actual = self._worm.consulta_sql_completa(self._tabla_actual)
        # self.tabular()

    def mov_reg_to_index(self):
        self._worm.consulta_sql_mover_registros_index()
        self._worm.clasificar_facturas_archivos("Facturas Descargadas", "Facturas Clasificadas")
        self._consulta_actual = self._worm.consulta_sql_completa("FACTURAS_NUEVAS")
        self._columnas_actual = self._worm.get_column_desc()
        self.tabular()
        self.actualizar_anios()

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
            print("C_OCULTAS: ", columnas)
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
        print("CONCEPTOS: ", consulta_de_conceptos)
        self._worm.exportar_xlsx(consulta_de_conceptos, columnas_con, "conceptos_asociados")

    def abortar(self):
        for pop_up in self._top_levels_activos:
            pop_up.destroy()
        self._root.destroy()
        self._worm.cerrar_conexion()
