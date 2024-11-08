# archivo_excel.py

import tkinter as tk
from tkinter import ttk, messagebox
from operaciones import exportar_a_excel

def mostrar_resultados(totales, contenedores, mensajes_contenedores, df_filtrado, columnas_mapeadas):
    root = tk.Tk()
    root.title("Resultados de Contenedores")
    root.geometry("800x600")  # Ajuste del tamaño de la ventana

    # Configurar estilo
    style = ttk.Style(root)
    style.theme_use("clam")  # Puedes cambiar el tema según preferencia

    # Definir estilos personalizados
    style.configure("Treeview.Heading",
                    font=("Calibri", 14, "bold"),
                    background="#0000FF",  # Azul
                    foreground="white",
                    borderwidth=0)
    style.configure("Treeview",
                    font=("Calibri", 14),
                    rowheight=30,
                    fieldbackground="#0000FF",  # Azul
                    background="#0000FF",  # Azul
                    foreground="white",
                    borderwidth=0,
                    highlightthickness=0)
    style.map("Treeview",
              # Azul claro al seleccionar
              background=[("selected", "#ADD8E6")],
              foreground=[("selected", "black")])

    # Mostrar la tabla de totales
    frame_totales = tk.Frame(root)
    frame_totales.pack(pady=10)

    tk.Label(frame_totales, text="Descripción", font=(
        "Arial", 12, "bold")).grid(row=0, column=0)
    tk.Label(frame_totales, text="Valor", font=(
        "Arial", 12, "bold")).grid(row=0, column=1)

    for i, (desc, valor) in enumerate(totales.items()):
        tk.Label(frame_totales, text=desc, font=(
            "Arial", 12)).grid(row=i+1, column=0)
        tk.Label(frame_totales, text=str(valor), font=(
            "Arial", 12)).grid(row=i+1, column=1)

    # Crear la tabla de contenedores
    frame_contenedores = tk.Frame(root)
    frame_contenedores.pack(pady=20)

    tk.Label(frame_contenedores, text="Contenedores",
             font=("Arial", 14, "bold")).pack()

    tree = ttk.Treeview(frame_contenedores, columns=(
        "Peso"), show="headings", height=10, style="Treeview")
    tree.heading("Peso", text="Peso Neto")
    tree.column("Peso", width=300)  # Ajuste del ancho de la columna
    tree.pack(expand=True, fill='both')  # Ajuste para expandir el Treeview

    for i, contenedor in enumerate(contenedores):
        tree.insert("", "end", iid=i, values=(
            f"Contenedor {i+1}: {contenedor:.2f} unidades de peso neto",))

    def on_select(event):
        if not tree.selection():
            return

        selected_item = tree.selection()[0]
        contenedor_index = int(selected_item)

        # Crear una nueva ventana para mostrar y mover los mensajes
        mensajes_ventana = tk.Toplevel(root)
        mensajes_ventana.title(f"Mensajes para Contenedor {
                               contenedor_index + 1}")
        mensajes_ventana.geometry("600x400")

        mensajes_listbox = tk.Listbox(
            mensajes_ventana, selectmode=tk.MULTIPLE, width=50, height=10)
        mensajes_listbox.pack(pady=10)

        for mensaje in mensajes_contenedores[contenedor_index]:
            mensajes_listbox.insert(tk.END, mensaje)

    tree.bind("<<TreeviewSelect>>", on_select)

    tk.Button(root, text="Exportar a Excel", command=lambda: exportar_a_excel(
        contenedores, mensajes_contenedores, df_filtrado, columnas_mapeadas)).pack(pady=10)

    root.mainloop()