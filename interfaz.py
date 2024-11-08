# interfaz.py

import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from tkinter import ttk
from operaciones import calcular_totales, calcular_contenedores, mostrar_resultados

class Interfaz:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Gestor de Contenedores")
        self.root.geometry("800x620")  # Tamaño ajustado de la ventana
        self.root.configure(bg="#f0f0f0")  # Color de fondo suave

        # Estilos
        self.style = ttk.Style()
        self.style.theme_use("clam")  # Tema moderno

        # Configuración de estilos personalizados
        self.configurar_estilos()

        # Título de la Aplicación
        titulo = ttk.Label(
            self.root,
            text="Gestor de Contenedores",
            font=("Arial", 20, "bold"),
            background="#f0f0f0",
            foreground="#4CAF50"
        )
        titulo.pack(pady=20)

        # Instrucción para el usuario (Etiqueta dinámica)
        self.instruccion = ttk.Label(
            self.root,
            text="Por favor, sube tu archivo Excel para comenzar:",
            font=("Arial", 12),
            background="#f0f0f0",
            foreground="#333333"
        )
        self.instruccion.pack(pady=10)

        # Botón para subir archivo
        self.boton_subir = ttk.Button(
            self.root,
            text="Subir Archivo",
            command=self.subir_archivo,
            style="Primary.TButton"
        )
        self.boton_subir.pack(pady=20)

        # Campo de búsqueda con su botón
        buscar_frame = ttk.Frame(self.root, style="TFrame")
        buscar_frame.pack(pady=10, fill='x', padx=50)

        self.entry_buscar = ttk.Entry(buscar_frame, font=("Arial", 12), style="TEntry")
        self.entry_buscar.pack(side='left', fill='x', expand=True, padx=(0, 10))
        self.entry_buscar.bind('<KeyRelease>', self.filtrar_datos)

        # Listbox para mostrar 'Nombre' con Scrollbar
        listbox_frame = ttk.Frame(self.root, style="TFrame")
        listbox_frame.pack(pady=10, fill='x', padx=50)

        self.listbox_nombres = tk.Listbox(
            listbox_frame,
            selectmode=tk.SINGLE,
            width=100,
            height=10,  # Ajuste de altura
            exportselection=0,
            font=("Arial", 12),
            bg="#ffffff",
            fg="#000000",
            highlightthickness=1,
            relief="solid"
        )
        self.listbox_nombres.pack(side='left', fill='both', expand=True)

        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.listbox_nombres.yview)
        scrollbar.pack(side='right', fill='y')
        self.listbox_nombres.config(yscrollcommand=scrollbar.set)

        # Botones adicionales
        botones_frame = ttk.Frame(self.root, style="TFrame")
        botones_frame.pack(pady=20)

        self.boton_buscar = ttk.Button(
            botones_frame,
            text="Buscar Datos",
            command=self.buscar_datos,
            style="Secondary.TButton"
        )
        self.boton_buscar.grid(row=0, column=0, padx=10)

        self.boton_procesar = ttk.Button(
            botones_frame,
            text="Procesar Selección",
            command=self.procesar_seleccion,
            style="Primary.TButton"
        )
        self.boton_procesar.grid(row=0, column=1, padx=10)

        # Nuevo Botón: Mostrar Inventario
        self.boton_mostrar_inventario = ttk.Button(
            botones_frame,
            text="Mostrar Inventario",
            command=self.mostrar_inventario,
            style="Secondary.TButton"
        )
        self.boton_mostrar_inventario.grid(row=0, column=2, padx=10)

        # Variable para la capacidad del contenedor
        self.capacidad_contenedor = tk.DoubleVar(value=33.2)  # Valor por defecto para 20ft

        # Datos completos para el filtrado
        self.df_subido = pd.DataFrame()
        self.datos_filtrados = pd.DataFrame()

    def configurar_estilos(self):
        # Estilos de los Botones
        self.style.configure("Primary.TButton",
                             font=("Arial", 12),
                             padding=10,
                             background="#2196F3",
                             foreground="white")
        self.style.configure("Secondary.TButton",
                             font=("Arial", 12),
                             padding=10,
                             background="#FF9800",
                             foreground="white")

        # Eliminar efectos de hover configurando map para que no cambie nada
        self.style.map("Primary.TButton",
                       background=[("active", "#2196F3")],
                       foreground=[("active", "white")])
        self.style.map("Secondary.TButton",
                       background=[("active", "#FF9800")],
                       foreground=[("active", "white")])

        # Estilo para Frame
        self.style.configure("TFrame", background="#f0f0f0")

        # Estilo para Entry
        self.style.configure("TEntry",
                             fieldbackground="#ffffff",
                             foreground="#000000",
                             bordercolor="#4CAF50",
                             borderwidth=1)

        # Estilo para Checkbutton y Radiobutton
        self.style.configure("TCheckbutton",
                             background="#f0f0f0",
                             foreground="#333333",
                             font=("Arial", 12))
        self.style.configure("TRadiobutton",
                             background="#f0f0f0",
                             foreground="#333333",
                             font=("Arial", 12))

    def subir_archivo(self):
        archivo_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if archivo_path:
            self.cargar_datos(archivo_path)

    def cargar_datos(self, archivo_path):
        try:
            df_subido = pd.read_excel(archivo_path, header=0)
            df_subido.dropna(how='all', inplace=True)
            df_subido.columns = df_subido.columns.str.strip()

            # Verificar si las columnas necesarias existen
            columnas_necesarias = ['Lote', 'Nombre', 'Cliente']
            for columna in columnas_necesarias:
                if columna not in df_subido.columns:
                    messagebox.showerror(
                        "Error",
                        f"La columna '{columna}' no se encontró en el archivo."
                    )
                    return

            # Crear la columna 'Grupo' basada en la columna 'Lote'
            df_subido['Grupo'] = df_subido['Lote'].astype(str).apply(
                lambda x: 'GCFOODS' if 'GZ' in x else 'NOEL'
            )

            # Eliminar espacios en la columna 'Nombre' y 'Cliente'
            df_subido['Nombre'] = df_subido['Nombre'].astype(str).str.replace(' ', '').str.strip()
            df_subido['Cliente'] = df_subido['Cliente'].astype(str).str.strip()

            # Solicitar al usuario que seleccione los grupos
            self.seleccionar_grupo(df_subido)

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar los datos: {e}")

    def seleccionar_grupo(self, df_subido):
        # Crear una nueva ventana para seleccionar los grupos
        grupo_window = tk.Toplevel(self.root)
        grupo_window.title("Seleccionar Grupo")
        grupo_window.geometry("450x300")
        grupo_window.configure(bg="#f0f0f0")

        # Estilos
        grupo_style = ttk.Style(grupo_window)
        grupo_style.theme_use("clam")
        grupo_style.configure("TLabel", background="#f0f0f0", font=("Arial", 12), foreground="#333333")
        grupo_style.configure("TButton", font=("Arial", 12), padding=10, background="#2196F3", foreground="white")
        grupo_style.map("TButton",
                        background=[("active", "#2196F3")],
                        foreground=[("active", "white")])
        grupo_style.configure("TCheckbutton",
                              background="#f0f0f0",
                              foreground="#333333",
                              font=("Arial", 12))

        # Título de la ventana
        titulo_grupo = ttk.Label(
            grupo_window,
            text="Seleccione los grupos de datos a procesar:",
            font=("Arial", 16, "bold"),
            background="#f0f0f0",
            foreground="#4CAF50"
        )
        titulo_grupo.pack(pady=20)

        # Variables para almacenar la selección
        self.grupo_gcfoods = tk.BooleanVar(value=True)
        self.grupo_noel = tk.BooleanVar(value=True)

        # Checkbuttons para seleccionar GCFOODS y NOEL
        check_gcfoods = ttk.Checkbutton(
            grupo_window,
            text="GCFOODS",
            variable=self.grupo_gcfoods
        )
        check_gcfoods.pack(pady=10, anchor='w', padx=50)

        check_noel = ttk.Checkbutton(
            grupo_window,
            text="NOEL",
            variable=self.grupo_noel
        )
        check_noel.pack(pady=10, anchor='w', padx=50)

        # Botón de Confirmar
        boton_confirmar = ttk.Button(
            grupo_window,
            text="Confirmar Selección",
            command=lambda: self.confirmar_grupo_seleccion(grupo_window, df_subido),
            style="Primary.TButton"
        )
        boton_confirmar.pack(pady=30)

    def confirmar_grupo_seleccion(self, ventana, df_subido):
        if not self.grupo_gcfoods.get() and not self.grupo_noel.get():
            messagebox.showwarning(
                "Advertencia",
                "Debe seleccionar al menos un grupo."
            )
            return
        ventana.destroy()
        self.mostrar_datos(df_subido)

    def mostrar_datos(self, df_subido):
        # Filtrar df_subido según los grupos seleccionados
        grupos_seleccionados = []
        if self.grupo_gcfoods.get():
            grupos_seleccionados.append('GCFOODS')
        if self.grupo_noel.get():
            grupos_seleccionados.append('NOEL')

        df_subido = df_subido[df_subido['Grupo'].isin(grupos_seleccionados)]

        # Modificar 'Lote' para el grupo GCFOODS
        df_subido.loc[df_subido['Grupo'] == 'GCFOODS', 'Lote'] = (
            '20' + df_subido.loc[df_subido['Grupo'] == 'GCFOODS', 'Lote'].astype(str)
        )

        # Asegurar que las columnas clave sean de tipo string
        df_subido['Lote'] = df_subido['Lote'].astype(str)
        df_subido = df_subido.dropna(subset=['Nombre', 'Cliente'])
        df_subido['Nombre'] = df_subido['Nombre'].astype(str).str.replace(' ', '').str.strip()
        df_subido['Cliente'] = df_subido['Cliente'].astype(str).str.strip()

        # Almacenar los datos para el filtrado
        self.df_subido = df_subido.copy()
        self.datos_filtrados = pd.DataFrame(
            {'Nombre': self.df_subido['Nombre'].unique()}
        )

        # Mostrar todos los datos inicialmente
        self.actualizar_listbox(self.datos_filtrados['Nombre'].tolist())

        # Actualizar la instrucción
        self.instruccion.config(text="Archivo cargado exitosamente. Por favor, seleccione los nombres que desea procesar.")

    def actualizar_listbox(self, lista_nombres):
        """
        Actualiza la Listbox con una lista de nombres, eliminando cualquier valor NaN,
        el string 'nan', o valores vacíos o de solo espacios.
        """
        # Filtrar para eliminar valores NaN, 'nan' (en cualquier forma) y nombres vacíos o de solo espacios
        lista_nombres_limpia = [nombre for nombre in lista_nombres if pd.notna(nombre) and nombre.lower() != 'nan' and nombre.strip() != '']

        # Limpiar la listbox actual antes de insertar los nuevos valores
        self.listbox_nombres.delete(0, tk.END)

        # Insertar los nombres limpios en la listbox
        for nombre in lista_nombres_limpia:
            self.listbox_nombres.insert(tk.END, nombre)


    def filtrar_datos(self, event):
        texto_busqueda = self.entry_buscar.get().lower()
        if texto_busqueda == '':
            nombres_filtrados = self.df_subido['Nombre'].unique()
        else:
            # Filtrar donde el texto coincida en 'Nombre' o 'Cliente'
            df_filtrado = self.df_subido[
                self.df_subido['Nombre'].str.lower().str.contains(texto_busqueda) |
                self.df_subido['Cliente'].str.lower().str.contains(texto_busqueda)
            ]
            nombres_filtrados = df_filtrado['Nombre'].unique()
        self.actualizar_listbox(nombres_filtrados)

    def procesar_seleccion(self):
        # Obtener los nombres seleccionados en el Listbox
        seleccionados = [
            self.listbox_nombres.get(i) for i in self.listbox_nombres.curselection()
        ]

        if not seleccionados:
            messagebox.showwarning(
                "Advertencia",
                "Debe seleccionar al menos un elemento para continuar."
            )
            return

        # Filtrar df_subido según los nombres seleccionados
        df_filtrado = self.df_subido[self.df_subido['Nombre'].isin(seleccionados)]

        # Actualizar la instrucción
        self.instruccion.config(text="Seleccione el tipo de contenedor.")

        # Antes de procesar el archivo, seleccionar el tipo de contenedor
        self.seleccionar_contenedor(df_filtrado)

    def seleccionar_contenedor(self, df_filtrado):
        # Crear una nueva ventana para seleccionar el tipo de contenedor
        contenedor_window = tk.Toplevel(self.root)
        contenedor_window.title("Seleccionar Tipo de Contenedor")
        contenedor_window.geometry("400x300")
        contenedor_window.configure(bg="#f0f0f0")

        # Estilos
        contenedor_style = ttk.Style(contenedor_window)
        contenedor_style.theme_use("clam")
        contenedor_style.configure("TLabel", background="#f0f0f0", font=("Arial", 12), foreground="#333333")
        contenedor_style.configure("TRadiobutton", background="#f0f0f0", font=("Arial", 12), foreground="#333333")
        contenedor_style.configure("TButton", font=("Arial", 12), padding=10, background="#2196F3", foreground="white")
        contenedor_style.map("TButton",
                             background=[("active", "#2196F3")],
                             foreground=[("active", "white")])

        # Título de la ventana
        titulo_contenedor = ttk.Label(
            contenedor_window,
            text="Seleccione el tipo de contenedor:",
            font=("Arial", 16, "bold"),
            background="#f0f0f0",
            foreground="#4CAF50"
        )
        titulo_contenedor.pack(pady=20)

        # Variable para la opción seleccionada
        opcion_seleccionada = tk.StringVar(value='20ft')  # Valor por defecto

        # Función para actualizar la capacidad según la selección
        def actualizar_capacidad():
            seleccion = opcion_seleccionada.get()
            if seleccion == '20ft':
                self.capacidad_contenedor.set(30)
            elif seleccion == '40st':
                self.capacidad_contenedor.set(58)
            elif seleccion == '40HC':
                self.capacidad_contenedor.set(69)
            elif seleccion == 'Sencillo':
                self.capacidad_contenedor.set(44)
            elif seleccion == 'Van':
                self.capacidad_contenedor.set(56)
            elif seleccion == 'Tractomula':
                self.capacidad_contenedor.set(62)
            elif seleccion == 'Otro':
                # Pedir al usuario que ingrese la capacidad
                capacidad = simpledialog.askfloat(
                    "Capacidad del Contenedor",
                    "Ingrese la capacidad máxima del contenedor:",
                    minvalue=0.1
                )
                if capacidad:
                    self.capacidad_contenedor.set(capacidad)
                else:
                    self.capacidad_contenedor.set(0)  # Valor por defecto si cancela
            else:
                self.capacidad_contenedor.set(0)

        # Opciones de contenedores con Radiobuttons estilizados
        opciones = ['20ft', '40st', '40HC', 'Sencillo', 'Van', 'Tractomula', 'Otro']

        # Crear un frame para los Radiobuttons
        frame_opciones = ttk.Frame(contenedor_window, style="TFrame")
        frame_opciones.pack(pady=10)

        # Usar grid para colocar los Radiobuttons en dos columnas
        for idx, opcion in enumerate(opciones):
            row = idx % 3
            col = idx // 3
            ttk.Radiobutton(
                frame_opciones,
                text=opcion,
                variable=opcion_seleccionada,
                value=opcion,
                command=actualizar_capacidad,
                style="TRadiobutton"
            ).grid(row=row, column=col, sticky='w', padx=20, pady=5)

        # Botón de Confirmar
        boton_confirmar = ttk.Button(
            contenedor_window,
            text="Confirmar Selección",
            command=lambda: self.confirmar_contenedor_seleccion(contenedor_window, df_filtrado),
            style="Primary.TButton"
        )
        boton_confirmar.pack(pady=30)

    def confirmar_contenedor_seleccion(self, ventana, df_filtrado):
        if self.capacidad_contenedor.get() <= 0:
            messagebox.showwarning(
                "Advertencia",
                "Debe seleccionar un tipo de contenedor válido."
            )
            return
        ventana.destroy()

        # Actualizar la instrucción
        self.instruccion.config(text="Procesando los datos. Por favor, espere...")

        # Después de cerrar la ventana, procesar el archivo
        self.procesar_archivo(df_filtrado)

    def procesar_archivo(self, df_filtrado):
        try:
            columnas_necesarias = [
                'Material', 'Texto de mensaje', 'Bruto', 'Neto', 'Volumen',
                'Importe', 'Cliente', 'Nombre', 'Contador'  # Agregamos 'Contador'
            ]

            # Intentar leer el archivo Portafolio con diferentes filas de encabezado
            df_portafolio = None
            for header_row in [0, 1, 2, 3, 4]:
                try:
                    df_temp = pd.read_excel(
                        "Data_Base\Portafoliocompletointernacional.xlsx",
                        header=header_row
                    )
                    df_temp.dropna(how='all', inplace=True)
                    df_temp.columns = df_temp.columns.str.strip()
                    if any(col in df_temp.columns for col in columnas_necesarias):
                        df_portafolio = df_temp
                        break
                except:
                    continue  # Intentar con la siguiente fila de encabezado

            if df_portafolio is None:
                raise ValueError(
                    "No se encontraron las columnas necesarias en el portafolio."
                )

            # Mapear las columnas
            def find_closest_column(column_name, df_columns):
                for col in df_columns:
                    if column_name.lower() == col.lower():
                        return col
                for col in df_columns:
                    if column_name.lower() in col.lower():
                        return col
                return None

            columnas_mapeadas = {}
            for columna in columnas_necesarias:
                columna_encontrada = find_closest_column(columna, df_portafolio.columns)
                if columna_encontrada:
                    columnas_mapeadas[columna] = columna_encontrada
                else:
                    if columna != 'Contador':
                        raise ValueError(
                            f"No se encontró una columna similar a '{columna}' "
                            "en el portafolio."
                        )

            # Eliminar espacios en la columna 'Nombre' del portafolio
            df_portafolio[columnas_mapeadas['Nombre']] = df_portafolio[
                columnas_mapeadas['Nombre']
            ].astype(str).str.replace(' ', '').str.strip()

            # Verificar y eliminar valores nulos en las columnas clave
            df_filtrado.dropna(
                subset=['Material', 'Texto breve material', 'Cliente', 'Nombre'],
                inplace=True
            )
            df_portafolio.dropna(
                subset=[
                    columnas_mapeadas['Material'],
                    columnas_mapeadas['Texto de mensaje'],
                    columnas_mapeadas['Cliente'],
                    columnas_mapeadas['Nombre']
                ],
                inplace=True
            )

            # Convertir las columnas clave a string y eliminar espacios en blanco
            claves_filtrado = ['Material', 'Texto breve material', 'Cliente', 'Nombre']
            claves_portafolio = [
                columnas_mapeadas['Material'],
                columnas_mapeadas['Texto de mensaje'],
                columnas_mapeadas['Cliente'],
                columnas_mapeadas['Nombre']
            ]

            for col in claves_filtrado:
                df_filtrado[col] = df_filtrado[col].astype(str).str.replace(' ', '').str.strip()
            for col in claves_portafolio:
                df_portafolio[col] = df_portafolio[col].astype(str).str.replace(' ', '').str.strip()

            # Realizar el merge
            df_resultado = pd.merge(
                df_filtrado,
                df_portafolio[list(columnas_mapeadas.values())],
                left_on=claves_filtrado,
                right_on=claves_portafolio,
                how='inner',
                suffixes=('_filtrado', '_portafolio')
            )

            if df_resultado.empty:
                messagebox.showerror(
                    "Error",
                    "No se encontraron coincidencias entre los archivos."
                )
                return

            # Columnas adicionales que queremos agregar del archivo subido
            columnas_adicionales = ['Doc.comer.', 'LibrUtiliz', 'Grupo', 'Lote']

            # Determinar el nombre correcto de las columnas adicionales
            columnas_adicionales_encontradas = []
            for col_adicional in columnas_adicionales:
                if f"{col_adicional}_filtrado" in df_resultado.columns:
                    columnas_adicionales_encontradas.append(f"{col_adicional}_filtrado")
                elif col_adicional in df_resultado.columns:
                    columnas_adicionales_encontradas.append(col_adicional)
                elif col_adicional == 'LibrUtiliz':
                    # Buscar 'Libre utilización' en su lugar
                    if 'Libre utilización_filtrado' in df_resultado.columns:
                        columnas_adicionales_encontradas.append('Libre utilización_filtrado')
                    elif 'Libre utilización' in df_resultado.columns:
                        columnas_adicionales_encontradas.append('Libre utilización')
                    else:
                        messagebox.showwarning(
                            "Advertencia",
                            f"'{col_adicional}' o 'Libre utilización' no se encontró en los datos."
                        )
                else:
                    messagebox.showwarning(
                        "Advertencia",
                        f"'{col_adicional}' no se encontró en los datos."
                    )

            # Filtrar las columnas necesarias
            columnas_para_df_final = list(columnas_mapeadas.values()) + \
                columnas_adicionales_encontradas
            df_final = df_resultado[columnas_para_df_final]

            # Renombrar las columnas adicionales
            columnas_renombradas = {}
            for col in columnas_adicionales_encontradas:
                if '_filtrado' in col:
                    col_sin_filtrado = col.replace('_filtrado', '')
                else:
                    col_sin_filtrado = col
                if col_sin_filtrado == 'Libre utilización':
                    columnas_renombradas[col] = 'LibrUtiliz'
                else:
                    columnas_renombradas[col] = col_sin_filtrado
            df_final.rename(columns=columnas_renombradas, inplace=True)

            # Eliminar comas y convertir a numérico
            for col in ['Bruto', 'Neto', 'Volumen', 'Importe']:
                df_final[col] = df_final[col].replace({',': ''}, regex=True)
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

            # Convertir 'LibrUtiliz' a numérico si no lo es
            if 'LibrUtiliz' in df_final.columns:
                df_final['LibrUtiliz'] = pd.to_numeric(
                    df_final['LibrUtiliz'], errors='coerce'
                ).fillna(0)

            # Calcular totales y contenedores
            totales = calcular_totales(df_final)
            total_neto = totales.get("Total peso Neto", 0)

            if total_neto == 0:
                raise ValueError(
                    "No se pudo calcular el Total peso Neto. Verifica los valores."
                )

            contenedores_pesos, mensajes_contenedores = calcular_contenedores(
                df_final, columnas_mapeadas, self.capacidad_contenedor.get()
            )

            mostrar_resultados(
                totales, contenedores_pesos, mensajes_contenedores,
                df_final, columnas_mapeadas, self.capacidad_contenedor.get()
            )

            # Al finalizar el procesamiento, actualizar la instrucción
            self.instruccion.config(text="Procesamiento completado. Puede ver los resultados.")

        except Exception as e:
            messagebox.showerror("Error", f"Error procesando el archivo: {e}")
            # Actualizar la instrucción en caso de error
            self.instruccion.config(text="Error al procesar el archivo. Por favor, intente nuevamente.")

    def buscar_datos(self):
        # Obtener los nombres seleccionados en el Listbox
        seleccionados = [
            self.listbox_nombres.get(i) for i in self.listbox_nombres.curselection()
        ]

        if not seleccionados:
            messagebox.showwarning(
                "Advertencia",
                "Debe seleccionar al menos un nombre en la lista para buscar."
            )
            return

        # Si hay múltiples nombres seleccionados, tomaremos el primero
        nombre_seleccionado = seleccionados[0]

        # Ventana para ingresar los datos de búsqueda
        input_window = tk.Toplevel(self.root)
        input_window.title("Buscar Datos")
        input_window.geometry("450x300")
        input_window.configure(bg="#f0f0f0")

        # Estilos
        input_style = ttk.Style(input_window)
        input_style.theme_use("clam")
        input_style.configure("TLabel", background="#f0f0f0", font=("Arial", 12), foreground="#333333")
        input_style.configure("TButton", font=("Arial", 12), padding=10, background="#2196F3", foreground="white")
        input_style.map("TButton",
                        background=[("active", "#2196F3")],
                        foreground=[("active", "white")])

        # Título de la ventana
        titulo_buscar = ttk.Label(
            input_window,
            text="Buscar Datos por Material",
            font=("Arial", 16, "bold"),
            background="#f0f0f0",
            foreground="#4CAF50"
        )
        titulo_buscar.pack(pady=20)

        # Campo para ingresar 'Material'
        frame_material = ttk.Frame(input_window, style="TFrame")
        frame_material.pack(pady=10, padx=50, fill='x')

        label_material = ttk.Label(
            frame_material,
            text="Material:",
            font=("Arial", 12),
            background="#f0f0f0",
            foreground="#333333"
        )
        label_material.pack(side='left', padx=(0, 10))

        material_entry = ttk.Entry(frame_material, font=("Arial", 12), style="TEntry")
        material_entry.pack(side='left', fill='x', expand=True)

        # Mostrar el nombre seleccionado
        label_nombre = ttk.Label(
            input_window,
            text=f"Nombre seleccionado: {nombre_seleccionado}",
            font=("Arial", 12),
            background="#f0f0f0",
            foreground="#333333"
        )
        label_nombre.pack(pady=10)

        # Botón para realizar la búsqueda
        boton_buscar = ttk.Button(
            input_window,
            text="Buscar",
            command=lambda: self.realizar_busqueda(
                material_entry.get(),
                nombre_seleccionado,
                input_window
            ),
            style="Primary.TButton"
        )
        boton_buscar.pack(pady=20)

    def realizar_busqueda(self, material, nombre, input_window):
        # Validar entrada de material
        if not material.strip():
            messagebox.showwarning(
                "Advertencia",
                "Debe ingresar un valor para 'Material'."
            )
            return

        # Cargar el portafolio
        df_portafolio, columnas_mapeadas = self.cargar_portafolio()

        if df_portafolio is None:
            return  # Error ya mostrado en cargar_portafolio

        # Asegurarse de que las columnas son cadenas y eliminar espacios
        df_portafolio[columnas_mapeadas['Material']] = df_portafolio[
            columnas_mapeadas['Material']
        ].astype(str).str.strip()
        df_portafolio[columnas_mapeadas['Nombre']] = df_portafolio[
            columnas_mapeadas['Nombre']
        ].astype(str).str.replace(' ', '').str.strip()
        df_portafolio[columnas_mapeadas['Texto de mensaje']] = df_portafolio[
            columnas_mapeadas['Texto de mensaje']
        ].astype(str).str.strip()
        df_portafolio[columnas_mapeadas['Cliente']] = df_portafolio[
            columnas_mapeadas['Cliente']
        ].astype(str).str.strip()

        # Realizar la búsqueda
        mask = (
            (df_portafolio[columnas_mapeadas['Material']] == material.strip()) &
            (df_portafolio[columnas_mapeadas['Nombre']] == nombre.strip())
        )
        df_resultado = df_portafolio[mask]

        if df_resultado.empty:
            messagebox.showinfo(
                "Sin Resultados",
                "No se encontraron coincidencias con los datos ingresados."
            )
        else:
            # Copiar los datos al portapapeles en formato Excel
            columnas_a_mostrar = [
                'Material', 'Texto de mensaje', 'Bruto', 'Neto', 'Volumen',
                'Importe', 'Cliente', 'Nombre', 'Contador'
            ]
            self.copiar_al_portapapeles(df_resultado, columnas_mapeadas, columnas_a_mostrar)
            # Cerrar la ventana de búsqueda
            input_window.destroy()

    def copiar_al_portapapeles(self, df_resultado, columnas_mapeadas, columnas_a_copiar):
        # Obtener los datos en el orden de las columnas
        datos = df_resultado[[columnas_mapeadas[col] for col in columnas_a_copiar]].copy()

        # Insertar una columna vacía después de "Texto de mensaje"
        if 'Texto de mensaje' in columnas_a_copiar:
            idx = columnas_a_copiar.index('Texto de mensaje')  # Encontrar el índice de la columna
            datos.insert(idx + 1, 'Columna Vacía', '')  # Insertar la columna vacía

        # Formatear números con comas en lugar de puntos
        def format_number_with_comma(x):
            try:
                return ('{0:.3f}'.format(float(x))).replace('.', ',')
            except:
                return x

        for col in ['Bruto', 'Neto', 'Volumen', 'Importe']:
            mapped_col = columnas_mapeadas.get(col)
            if mapped_col in datos.columns:
                datos[mapped_col] = datos[mapped_col].apply(format_number_with_comma)

        # Convertir los datos a texto separado por tabulaciones
        datos_string = datos.to_csv(sep='\t', index=False, header=False)

        # Copiar al portapapeles
        self.root.clipboard_clear()
        self.root.clipboard_append(datos_string)

        messagebox.showinfo(
            "Datos Copiados",
            "Los datos han sido copiados al portapapeles en formato Excel."
        )

    def cargar_portafolio(self):
        try:
            columnas_necesarias = [
                'Material', 'Texto de mensaje', 'Bruto', 'Neto', 'Volumen',
                'Importe', 'Cliente', 'Nombre', 'Contador'
            ]

            # Intentar leer el archivo con diferentes filas de encabezado
            df_portafolio = None
            for header_row in [0, 1, 2, 3, 4]:
                try:
                    df_temp = pd.read_excel(
                        "Data_Base\Portafoliocompletointernacional.xlsx",
                        header=header_row
                    )
                    df_temp.dropna(how='all', inplace=True)
                    df_temp.columns = df_temp.columns.str.strip()
                    if any(col in df_temp.columns for col in columnas_necesarias):
                        df_portafolio = df_temp
                        break
                except:
                    continue  # Intentar con la siguiente fila de encabezado

            if df_portafolio is None:
                raise ValueError(
                    "No se encontraron las columnas necesarias en el portafolio."
                )

            # Mapear las columnas
            def find_closest_column(column_name, df_columns):
                for col in df_columns:
                    if column_name.lower() == col.lower():
                        return col
                for col in df_columns:
                    if column_name.lower() in col.lower():
                        return col
                return None

            columnas_mapeadas = {}
            for columna in columnas_necesarias:
                columna_encontrada = find_closest_column(columna, df_portafolio.columns)
                if columna_encontrada:
                    columnas_mapeadas[columna] = columna_encontrada
                else:
                    if columna != 'Contador':
                        raise ValueError(
                            f"No se encontró una columna similar a '{columna}' "
                            "en el portafolio."
                        )

            # Eliminar espacios en la columna 'Nombre' del portafolio
            df_portafolio[columnas_mapeadas['Nombre']] = df_portafolio[
                columnas_mapeadas['Nombre']
            ].astype(str).str.replace(' ', '').str.strip()

            return df_portafolio, columnas_mapeadas
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el portafolio: {e}")
            return None, None

    def run(self):
        self.root.mainloop()

    def mostrar_inventario(self):
        # Obtener los nombres seleccionados en el Listbox
        seleccionados = [
            self.listbox_nombres.get(i) for i in self.listbox_nombres.curselection()
        ]

        if not seleccionados:
            messagebox.showwarning(
                "Advertencia",
                "Debe seleccionar al menos un nombre para mostrar el inventario."
            )
            return

        # Filtrar los datos según los nombres seleccionados
        df_inventario = self.df_subido[self.df_subido['Nombre'].isin(seleccionados)]

        if df_inventario.empty:
            messagebox.showinfo(
                "Sin Datos",
                "No hay datos para los nombres seleccionados."
            )
            return

        # Crear una nueva ventana para mostrar el inventario
        inventario_window = tk.Toplevel(self.root)
        inventario_window.title("Inventario")
        inventario_window.geometry("1000x600")  # Ajustar el tamaño según sea necesario

        # Configurar estilo para la nueva ventana
        style_inventario = ttk.Style(inventario_window)
        style_inventario.theme_use("clam")
        style_inventario.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#2196F3", foreground="white")
        style_inventario.configure("Treeview", font=("Arial", 11), rowheight=25, fieldbackground="#f0f0f0")
        style_inventario.map("Treeview", background=[("selected", "#ADD8E6")], foreground=[("selected", "black")])

        # Frame para el Treeview
        frame_inventario = ttk.Frame(inventario_window)
        frame_inventario.pack(fill='both', expand=True)

        # Scrollbar vertical
        scrollbar_v = ttk.Scrollbar(frame_inventario, orient='vertical')
        scrollbar_v.pack(side='right', fill='y')

        # Scrollbar horizontal
        scrollbar_h = ttk.Scrollbar(inventario_window, orient='horizontal')
        scrollbar_h.pack(side='bottom', fill='x')

        # Crear el Treeview
        tree_inventario = ttk.Treeview(frame_inventario, yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        tree_inventario.pack(side='left', fill='both', expand=True)

        scrollbar_v.config(command=tree_inventario.yview)
        scrollbar_h.config(command=tree_inventario.xview)

        # Definir las columnas que deseas mostrar
        columnas_inventario = list(df_inventario.columns)
        tree_inventario['columns'] = columnas_inventario
        tree_inventario['show'] = 'headings'

        # Definir los encabezados y configurar las columnas
        for col in columnas_inventario:
            tree_inventario.heading(col, text=col)
            tree_inventario.column(col, width=100, minwidth=100, stretch=True)
            # Puedes ajustar 'width' y 'minwidth' según tus necesidades

        # Insertar los datos en el Treeview
        for _, row in df_inventario.iterrows():
            valores = [row[col] for col in columnas_inventario]
            tree_inventario.insert("", "end", values=valores)

        # Botón para exportar a Excel
        boton_exportar = ttk.Button(
            inventario_window,
            text="Exportar a Excel",
            command=lambda: self.exportar_inventario_a_excel(df_inventario)
        )
        boton_exportar.pack(pady=10)

    def exportar_inventario_a_excel(self, df_inventario):
        # Pedir al usuario que seleccione la ubicación y nombre del archivo
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
            title="Guardar Inventario como"
        )
        if file_path:
            try:
                df_inventario.to_excel(file_path, index=False)
                messagebox.showinfo("Exportación Exitosa", f"El inventario ha sido exportado a:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo exportar el inventario:\n{e}")

if __name__ == "__main__":
    app = Interfaz()
    app.run()
