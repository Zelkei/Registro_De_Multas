import os
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
from datetime import datetime
from ttkthemes import ThemedStyle
from tkcalendar import DateEntry
import pandas as pd

class GestorDocumentos:
    def __init__(self, root):
        self.root = root
        root.title("Gestor de Documentos Legales")

        self.style = ThemedStyle(root)
        self.style.set_theme("plastik")  # Puedes cambiar el tema aquí

        self.conn = sqlite3.connect('gestor_documentos.db')
        self.c = self.conn.cursor()

        self.c.execute('''CREATE TABLE IF NOT EXISTS documentos
                  (id INTEGER PRIMARY KEY AUTOINCREMENT,
                   nombre TEXT,
                   ruta TEXT,
                   tipo_archivo TEXT,
                   fecha_creacion TEXT,
                   etiquetas TEXT,
                   favorito BOOLEAN DEFAULT 0,
                   descripcion TEXT,
                   fecha_termino TEXT DEFAULT NULL)''')

        self.c.execute('''CREATE TABLE IF NOT EXISTS fechas_importantes
                  (id INTEGER PRIMARY KEY AUTOINCREMENT,
                   fecha TEXT,
                   descripcion TEXT)''')

        self.create_ui()

        self.registrando = True

    def create_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Frame para datos del documento
        frame_documento = ttk.LabelFrame(main_frame, text="Datos del Documento")
        frame_documento.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.label_nombre = tk.Label(frame_documento, text="Nombre:")
        self.entry_nombre = tk.Entry(frame_documento)
        self.label_ruta = tk.Label(frame_documento, text="Ruta:")
        self.entry_ruta = tk.Entry(frame_documento)
        self.button_seleccionar_archivo = tk.Button(frame_documento, text="Seleccionar Archivo", command=self.seleccionar_archivo)
        self.label_tipo = tk.Label(frame_documento, text="Tipo de Archivo:")
        self.entry_tipo = tk.Entry(frame_documento)
        self.label_etiquetas = tk.Label(frame_documento, text="Etiquetas:")
        self.entry_etiquetas = tk.Entry(frame_documento)
        self.label_descripcion = tk.Label(frame_documento, text="Descripción:")
        self.entry_descripcion = tk.Entry(frame_documento)
        self.label_fecha_termino = tk.Label(frame_documento, text="Fecha de Término:")
        self.entry_fecha_termino = DateEntry(frame_documento, date_pattern='yyyy-mm-dd')

        self.label_nombre.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_nombre.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.label_ruta.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_ruta.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.button_seleccionar_archivo.grid(row=1, column=2, padx=5, pady=5)
        self.label_tipo.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.entry_tipo.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.label_etiquetas.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.entry_etiquetas.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.label_descripcion.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.entry_descripcion.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.label_fecha_termino.grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.entry_fecha_termino.grid(row=5, column=1, padx=5, pady=5, sticky="w")

        # Frame para botones y lista de documentos
        frame_botones_documentos = ttk.LabelFrame(main_frame, text="Documentos")
        frame_botones_documentos.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        self.button_agregar = tk.Button(frame_botones_documentos, text="Agregar Documento", command=self.agregar_documento)
        self.button_eliminar_documento = tk.Button(frame_botones_documentos, text="Eliminar Documento", command=self.confirmar_eliminar_documento)
        self.button_exportar_documentos = tk.Button(frame_botones_documentos, text="Exportar Documentos a Excel", command=self.exportar_documentos_a_excel)

        self.button_agregar.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.button_eliminar_documento.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.button_exportar_documentos.grid(row=2, column=0, padx=5, pady=5, sticky="ew")



        self.treeview = ttk.Treeview(main_frame, columns=("Nombre",), show="headings", height=15)
        self.treeview.heading("Nombre", text="Documentos")
        self.treeview.column("Nombre", width=600)  # Ajusta el ancho como desees
        self.treeview.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        # Frame para fechas importantes
        frame_fechas_importantes = ttk.LabelFrame(main_frame, text="Fechas Importantes")
        frame_fechas_importantes.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")

        self.button_agregar_fecha = tk.Button(frame_fechas_importantes, text="Agregar Fecha Importante", command=self.agregar_fecha_importante)
        self.button_eliminar_fecha = tk.Button(frame_fechas_importantes, text="Eliminar Fecha Importante", command=self.confirmar_eliminar_fecha_importante)

        self.button_agregar_fecha.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.button_eliminar_fecha.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        self.treeview_fechas_importantes = ttk.Treeview(frame_fechas_importantes, columns=("Fecha", "Descripción"), show="headings", height=15)
        self.treeview_fechas_importantes.heading("Fecha", text="Fecha")
        self.treeview_fechas_importantes.heading("Descripción", text="Descripción")
        self.treeview_fechas_importantes.column("Fecha", width=150)
        self.treeview_fechas_importantes.column("Descripción", width=400)
        self.treeview_fechas_importantes.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        frame_fechas_importantes.rowconfigure(0, weight=1)
        frame_fechas_importantes.columnconfigure(0, weight=1)

        # Lista de documentos
        self.treeview.bind("<Double-Button-1>", self.vista_previa_doble_clic)
        self.treeview.bind("<Button-3>", self.mostrar_menu_contextual)

        self.listar_documentos()
        self.listar_fechas_importantes()

        self.crear_menu_contextual()

        self.root.after(86400000, self.verificar_alertas_vencimiento)  # 24 horas en milisegundos

    def verificar_alertas_vencimiento(self):
        hoy = datetime.now()
        self.c.execute("SELECT nombre, fecha_termino FROM documentos WHERE fecha_termino IS NOT NULL")
        documentos = self.c.fetchall()

        for nombre, fecha_termino in documentos:
            fecha_termino = datetime.strptime(fecha_termino, '%Y-%m-%d')  # Convertir fecha de término a objeto datetime
            diferencia = fecha_termino - hoy
            if 0 < diferencia.days <= 30:
                messagebox.showwarning("Alerta de Vencimiento", f"Faltan {diferencia.days} días para el vencimiento del documento '{nombre}'.")

    def seleccionar_archivo(self):
        archivo_seleccionado = filedialog.askopenfilename()
        self.entry_ruta.delete(0, tk.END)
        self.entry_ruta.insert(0, archivo_seleccionado)

    def agregar_documento(self):
        if not self.registrando:
            messagebox.showinfo("Registro Desactivado", "El registro de documentos está desactivado. Activa la casilla para registrar documentos.")
            return

        nombre = self.entry_nombre.get()

        if not nombre:  # Verificar si el campo de nombre está vacío
            messagebox.showerror("Nombre Vacío", "Por favor, ingresa un nombre para el documento.")
            return

        ruta = self.entry_ruta.get()
        tipo_archivo = self.entry_tipo.get()
        etiquetas = self.entry_etiquetas.get()
        descripcion = self.entry_descripcion.get()
        fecha_termino = self.entry_fecha_termino.get()  # Obtener la fecha de término

        fecha_creacion = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        self.c.execute("INSERT INTO documentos (nombre, ruta, tipo_archivo, fecha_creacion, etiquetas, descripcion, fecha_termino) VALUES (?, ?, ?, ?, ?, ?, ?)",
                       (nombre, ruta, tipo_archivo, fecha_creacion, etiquetas, descripcion, fecha_termino))
        self.conn.commit()
        self.listar_documentos()

        self.toggle_registro()  # Desactivar el registro después de agregar un documento

    def listar_documentos(self):
        self.treeview.delete(*self.treeview.get_children())  # Borra todos los elementos de la tabla
        self.c.execute("SELECT nombre, descripcion, ruta FROM documentos")
        documentos = self.c.fetchall()
        for documento in documentos:
            self.treeview.insert("", "end", values=documento)

    def marcar_favorito(self):
        seleccion = self.treeview.selection()
        if seleccion:
            indice = seleccion[0]
            documento = self.treeview.item(indice, "values")
            favorito = not documento[6]
            self.c.execute("UPDATE documentos SET favorito=? WHERE nombre=?", (favorito, documento[0]))
            self.conn.commit()
            self.listar_documentos()

    def vista_previa(self):
        seleccion = self.treeview.selection()
        if seleccion:
            indice = seleccion[0]
            documento = self.treeview.item(indice, "values")
            archivo_ruta = documento[2]
            try:
                archivo_ruta = os.path.normpath(archivo_ruta)  # Normalizar la ruta (opcional)
                os.startfile(archivo_ruta)  # Abrir el archivo con la aplicación predeterminada
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir la vista previa del documento: {str(e)}")

    def vista_previa_doble_clic(self, event):
        self.vista_previa()

    def mostrar_menu_contextual(self, event):
        self.menu_contextual.post(event.x_root, event.y_root)

    def crear_menu_contextual(self):
        self.menu_contextual = Menu(self.treeview, tearoff=0)
        self.menu_contextual.add_command(label="Vista Previa", command=self.vista_previa)
        self.menu_contextual.add_command(label="Agregar Datos Adicionales", command=self.agregar_datos_adicionales)
        self.menu_contextual.add_command(label="Consultar Información Adicional", command=self.consultar_informacion_adicional)
        self.menu_contextual.add_command(label="Editar Fecha de Término", command=self.editar_fecha_termino)  # Agregar opción para editar fecha de término

    def confirmar_eliminar_documento(self):
        seleccion = self.treeview.selection()
        if seleccion:
            respuesta = messagebox.askyesno("Confirmación", "¿Estás seguro de que deseas eliminar este documento?")
            if respuesta:
                self.eliminar_documento(seleccion)

    def eliminar_documento(self, seleccion):
        for indice in seleccion:
            documento = self.treeview.item(indice, "values")
            nombre = documento[0]
            self.c.execute("DELETE FROM documentos WHERE nombre=?", (nombre,))
            self.conn.commit()
        self.listar_documentos()
    def exportar_documentos_a_excel(self):
        self.c.execute("SELECT nombre, descripcion, ruta FROM documentos")
        documentos = self.c.fetchall()
        df = pd.DataFrame(documentos, columns=["Nombre del Documento", "Descripción", "Ruta del Archivo"])
        ruta_guardar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if ruta_guardar:
            df.to_excel(ruta_guardar, index=False)
            messagebox.showinfo("Exportación Exitosa", "La lista de documentos se ha exportado exitosamente a Excel.")
    def exportar_fechas_a_excel(self):
        self.c.execute("SELECT fecha, descripcion FROM fechas_importantes")
        fechas_importantes = self.c.fetchall()
        df = pd.DataFrame(fechas_importantes, columns=["Fecha", "Descripción"])
        ruta_guardar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if ruta_guardar:
            df.to_excel(ruta_guardar, index=False)
            messagebox.showinfo("Exportación Exitosa", "La lista de fechas importantes se ha exportado exitosamente a Excel.")


    def agregar_datos_adicionales(self):
        seleccion = self.treeview.selection()
        if not seleccion:
            messagebox.showerror("Error", "Por favor, selecciona un documento para agregar datos adicionales.")
            return

        indice = seleccion[0]
        documento = self.treeview.item(indice, "values")

        dialogo = tk.Toplevel(self.root)
        dialogo.title(f"Agregar Datos Adicionales para: {documento[0]}")

        campos = [
            "Fecha de Suscripción",
            "Fecha de Término",
            "Formalidades para Notificar Término",
            "Individualización de las Partes",
            "Servicios/Productos Ofrecidos",
            "Precio",
            "Periodos de Pago",
            "Obligaciones",
            "Prohibiciones",
            "Cláusulas de Término con Indemnización",
            "Sistemas de Resolución de Conflictos",
            "Anexos",
            "Extensiones",
            "Addendum"
        ]

        etiquetas = []
        entradas = []

        for campo in campos:
            etiqueta = tk.Label(dialogo, text=f"{campo}:")
            etiqueta.pack()
            entrada = tk.Entry(dialogo, width=40)
            entrada.pack()
            etiquetas.append(etiqueta)
            entradas.append(entrada)

        boton_guardar = tk.Button(dialogo, text="Guardar", command=lambda: self.guardar_datos_adicionales(documento[0], campos, entradas))
        boton_guardar.pack()

    def guardar_datos_adicionales(self, documento_nombre, campos, entradas):
        datos_adicionales = ""
        for campo, entrada in zip(campos, entradas):
            valor = entrada.get()
            if valor:
                datos_adicionales += f"{campo}: {valor}\n"

        self.c.execute("SELECT descripcion FROM documentos WHERE nombre=?", (documento_nombre,))
        descripcion_actual = self.c.fetchone()

        if descripcion_actual is not None:
            descripcion_actual = descripcion_actual[0]
        else:
            descripcion_actual = ""

        nueva_descripcion = "\n".join([descripcion_actual, datos_adicionales])

        self.c.execute("UPDATE documentos SET descripcion=? WHERE nombre=?", (nueva_descripcion, documento_nombre))
        self.conn.commit()
        self.listar_documentos()

        messagebox.showinfo("Confirmación", "Los datos han sido guardados correctamente.")
        self.root.focus_set()

    def consultar_informacion_adicional(self):
        seleccion = self.treeview.selection()
        if not seleccion:
            messagebox.showerror("Error", "Por favor, selecciona un documento para consultar la información adicional.")
            return

        indice = seleccion[0]
        documento = self.treeview.item(indice, "values")

        dialogo = tk.Toplevel(self.root)
        dialogo.title(f"Información Adicional para: {documento[0]}")

        treeview = ttk.Treeview(dialogo, columns=("Campo", "Valor"), show="headings", height=15)
        treeview.heading("Campo", text="Campo")
        treeview.heading("Valor", text="Valor")
        treeview.column("Campo", width=200)
        treeview.column("Valor", width=400)
        treeview.pack()

        descripcion = documento[1]
        if descripcion:
            campos_y_valores = [line.split(": ", 1) for line in descripcion.split("\n") if ": " in line]
        
            # Crear una fila especial para mostrar la descripción en primer lugar
            if descripcion.strip():  # Asegurarse de que la descripción no esté vacía
                treeview.insert("", "end", values=("Descripción", descripcion.strip()))

            # Agregar otros campos y valores
            for campo, valor in campos_y_valores:
                if campo != "Descripción":  # Omitir duplicados
                    treeview.insert("", "end", values=(campo, valor))

        boton_editar = tk.Button(dialogo, text="Editar Información Adicional", command=lambda: self.editar_informacion_adicional(documento[0], treeview))
        boton_editar.pack()

    def editar_informacion_adicional(self, documento_nombre, treeview):
        seleccion = treeview.selection()
        if not seleccion:
            messagebox.showerror("Error", "Por favor, selecciona una fila para editar.")
            return

        indice = seleccion[0]
        fila = treeview.item(indice, "values")
        campo = fila[0]
        valor = fila[1]

        dialogo_editar = tk.Toplevel(self.root)
        dialogo_editar.title(f"Editar Información Adicional para: {documento_nombre}")

        label_campo = tk.Label(dialogo_editar, text="Campo:")
        label_campo.grid(row=0, column=0)
        entry_campo = tk.Entry(dialogo_editar, width=40)
        entry_campo.insert(0, campo)
        entry_campo.grid(row=0, column=1)

        label_valor = tk.Label(dialogo_editar, text="Valor:")
        label_valor.grid(row=1, column=0)
        entry_valor = tk.Entry(dialogo_editar, width=40)
        entry_valor.insert(0, valor)
        entry_valor.grid(row=1, column=1)

        boton_guardar = tk.Button(dialogo_editar, text="Guardar Cambios", command=lambda: self.guardar_cambios_informacion_adicional(documento_nombre, treeview, indice, entry_campo, entry_valor))
        boton_guardar.grid(row=2, columnspan=2)

    def guardar_cambios_informacion_adicional(self, documento_nombre, treeview, indice, entry_campo, entry_valor):
        campo_nuevo = entry_campo.get()
        valor_nuevo = entry_valor.get()

        treeview.item(indice, values=(campo_nuevo, valor_nuevo))

        # Actualizar la descripción en la base de datos
        filas = []
        for item in treeview.get_children():
            fila = treeview.item(item, "values")
            if fila:
                filas.append(f"{fila[0]}: {fila[1]}")
        nueva_descripcion = "\n".join(filas)

        self.c.execute("UPDATE documentos SET descripcion=? WHERE nombre=?", (nueva_descripcion, documento_nombre))
        self.conn.commit()

        messagebox.showinfo("Confirmación", "Los cambios han sido guardados correctamente.")
        treeview.focus_set()

    def editar_fecha_termino(self):
        seleccion = self.treeview.selection()
        if not seleccion:
            messagebox.showerror("Error", "Por favor, selecciona un documento para editar la fecha de término.")
            return

        indice = seleccion[0]
        documento = self.treeview.item(indice, "values")

        dialogo = tk.Toplevel(self.root)
        dialogo.title(f"Editar Fecha de Término para: {documento[0]}")

        label_fecha_termino = tk.Label(dialogo, text="Fecha de Término:")
        label_fecha_termino.pack()
        entry_fecha_termino = DateEntry(dialogo, date_pattern='yyyy-mm-dd')
        entry_fecha_termino.pack()

        boton_guardar = tk.Button(dialogo, text="Guardar Cambios", command=lambda: self.guardar_fecha_termino(documento[0], entry_fecha_termino))
        boton_guardar.pack()

    def guardar_fecha_termino(self, documento_nombre, entry_fecha_termino):
        nueva_fecha_termino = entry_fecha_termino.get()
        self.c.execute("UPDATE documentos SET fecha_termino=? WHERE nombre=?", (nueva_fecha_termino, documento_nombre))
        self.conn.commit()
        self.listar_documentos()
        messagebox.showinfo("Confirmación", f"La fecha de término para '{documento_nombre}' ha sido actualizada correctamente.")
        self.root.focus_set()

    def toggle_registro(self):
        if self.registrando:
            self.registrando = False
            self.button_agregar.config(state="disabled")
        else:
            self.registrando = True
            self.button_agregar.config(state="normal")

    def agregar_fecha_importante(self):
        dialogo = tk.Toplevel(self.root)
        dialogo.title("Agregar Fecha Importante")

        label_fecha = tk.Label(dialogo, text="Fecha:")
        label_fecha.pack()
        entry_fecha = DateEntry(dialogo, date_pattern='yyyy-mm-dd')
        entry_fecha.pack()

        label_descripcion = tk.Label(dialogo, text="Descripción:")
        label_descripcion.pack()
        entry_descripcion = tk.Entry(dialogo, width=40)
        entry_descripcion.pack()

        boton_guardar = tk.Button(dialogo, text="Guardar", command=lambda: self.guardar_fecha_importante(entry_fecha, entry_descripcion))
        boton_guardar.pack()

    def guardar_fecha_importante(self, entry_fecha, entry_descripcion):
        fecha = entry_fecha.get()
        descripcion = entry_descripcion.get()

        if not fecha:
            messagebox.showerror("Error", "Por favor, ingresa una fecha.")
            return

        self.c.execute("INSERT INTO fechas_importantes (fecha, descripcion) VALUES (?, ?)", (fecha, descripcion))
        self.conn.commit()
        self.listar_fechas_importantes()
        messagebox.showinfo("Confirmación", "La fecha importante ha sido agregada correctamente.")
        self.root.focus_set()

    def confirmar_eliminar_fecha_importante(self):
        seleccion = self.treeview_fechas_importantes.selection()
        if seleccion:
            respuesta = messagebox.askyesno("Confirmación", "¿Estás seguro de que deseas eliminar esta fecha importante?")
            if respuesta:
                self.eliminar_fecha_importante(seleccion)

    def eliminar_fecha_importante(self, seleccion):
        for indice in seleccion:
            fecha_importante = self.treeview_fechas_importantes.item(indice, "values")
            fecha = fecha_importante[0]
            self.c.execute("DELETE FROM fechas_importantes WHERE fecha=?", (fecha,))
            self.conn.commit()
        self.listar_fechas_importantes()

    def listar_fechas_importantes(self):
        self.treeview_fechas_importantes.delete(*self.treeview_fechas_importantes.get_children())  # Borra todos los elementos de la tabla
        self.c.execute("SELECT fecha, descripcion FROM fechas_importantes")
        fechas_importantes = self.c.fetchall()
        for fecha_importante in fechas_importantes:
            self.treeview_fechas_importantes.insert("", "end", values=fecha_importante)

    def run(self):
        self.root.mainloop()

    def __del__(self):
        self.conn.close()

if __name__ == '__main__':
    root = tk.Tk()
    app = GestorDocumentos(root)
    app.run()

