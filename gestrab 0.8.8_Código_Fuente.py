import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import sqlite3
import pandas as pd
from ttkthemes import ThemedStyle
from ttkthemes import ThemedTk
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

class TrabajadoresApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Registro de Trabajadores")
        
        self.categorias = ["Administración", "Autónoma", "Demaria", "Engie Norte", "Engie Q", "Ecologico", "Clean Xpert", "Spot", "Taller", "Varios", "Conexxa", "Frutos de Lonquén"]

        self.tablas_trabajadores = {}
        self.tablas_actuales = None

        self.create_base_de_datos()

        self.listbook = ttk.Notebook(self.root)
        self.listbook.pack(fill='both', expand=True)

        for categoria in self.categorias:
            frame = ttk.Frame(self.listbook)
            self.listbook.add(frame, text=categoria)

            columnas = ["ID", "Trabajador", "Rut", "Cargo", "Fecha de Contrato", "Tipo Último Anexo", "Actualización en DT", "Fecha Actualización DT"]
            tabla = ttk.Treeview(frame, columns=columnas, show='headings')
            tabla.bind("<Button-3>", self.mostrar_menu_contextual)
            for col in columnas:
                tabla.heading(col, text=col)
                tabla.column(col, width=120)

            tabla.pack(fill='both', expand=True)
            self.tablas_trabajadores[categoria] = tabla

            tabla.bind("<Button-3>", self.mostrar_menu_contextual)
            tabla.bind("<Double-1>", self.ver_datos_adicionales)

        self.create_input_fields()
        self.create_buttons()
        self.bind_input_fields()
        self.disable_add_button()
        self.load_data_for_category(self.categorias[0])
        self.create_update_button()  # Agregar un botón de actualización
# Cargar datos al iniciar la aplicación
        for categoria in self.categorias:
            self.load_data_for_category(categoria)

    def create_update_button(self):
        update_button = ttk.Button(self.root, text="Actualizar", command=self.actualizar_datos)
        update_button.pack()

    def actualizar_datos(self):
        # Función para actualizar los datos en la base de datos y en la interfaz gráfica
        try:
            categoria_actual = self.listbook.tab(self.listbook.select(), "text")
            tabla = self.tablas_trabajadores[categoria_actual]
            seleccion = tabla.selection()

            if not seleccion:
                messagebox.showinfo("Mensaje", "Selecciona un trabajador para actualizar.")
                return

            id_trabajador = seleccion[0]

            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()

            # Aquí debes obtener los nuevos valores de los campos y actualizar la base de datos
            # Puedes usar simpledialog para obtener los nuevos valores, por ejemplo:
            nuevo_trabajador = simpledialog.askstring("Actualizar Trabajador", "Nuevo Trabajador:")
            if nuevo_trabajador is not None:
                cursor.execute('UPDATE trabajadores SET trabajador = ? WHERE id = ?', (nuevo_trabajador, id_trabajador))
                conexion.commit()
                conexion.close()

                # Actualiza la interfaz gráfica
                self.load_data_for_category(categoria_actual)
                messagebox.showinfo("Mensaje", "Datos actualizados correctamente.")
        except sqlite3.Error as e:
            self.show_error_message(f"Error al actualizar trabajador: {str(e)}")

    def create_base_de_datos(self):
        try:
            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS trabajadores (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    categoria TEXT,
                    trabajador TEXT,
                    rut TEXT,
                    cargo TEXT,
                    fecha_contrato TEXT,
                    tipo_anexo TEXT,
                    actualizacion_dt TEXT,
                    fecha_actualizacion_dt TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS datos_adicionales (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    id_trabajador INTEGER,
                    historial_anexos TEXT,
                    datos_extras TEXT
                )
            ''')

            conexion.commit()
            conexion.close()
        except sqlite3.Error as e:
            print(f"Error al crear la base de datos: {str(e)}")

    def create_input_fields(self):
        self.entry_trabajador = ttk.Entry(self.root)
        self.entry_rut = ttk.Entry(self.root)
        self.entry_cargo = ttk.Entry(self.root)
        self.entry_fecha_contrato = ttk.Entry(self.root)
        self.entry_tipo_anexo = ttk.Entry(self.root)
        self.entry_actualizacion_dt = ttk.Entry(self.root)
        self.entry_fecha_actualizacion_dt = ttk.Entry(self.root)

        etiqueta_trabajador = ttk.Label(self.root, text="Trabajador:")
        etiqueta_rut = ttk.Label(self.root, text="Rut:")
        etiqueta_cargo = ttk.Label(self.root, text="Cargo:")
        etiqueta_fecha_contrato = ttk.Label(self.root, text="Fecha de Contrato:")
        etiqueta_tipo_anexo = ttk.Label(self.root, text="Tipo Último Anexo:")
        etiqueta_actualizacion_dt = ttk.Label(self.root, text="Actualización en DT:")
        etiqueta_fecha_actualizacion_dt = ttk.Label(self.root, text="Fecha Actualización DT:")

        etiqueta_trabajador.pack()
        self.entry_trabajador.pack()
        etiqueta_rut.pack()
        self.entry_rut.pack()
        etiqueta_cargo.pack()
        self.entry_cargo.pack()
        etiqueta_fecha_contrato.pack()
        self.entry_fecha_contrato.pack()
        etiqueta_tipo_anexo.pack()
        self.entry_tipo_anexo.pack()
        etiqueta_actualizacion_dt.pack()
        self.entry_actualizacion_dt.pack()
        etiqueta_fecha_actualizacion_dt.pack()
        self.entry_fecha_actualizacion_dt.pack()

    def create_buttons(self):
        self.boton_agregar = ttk.Button(self.root, text="Agregar Trabajador", command=self.agregar_trabajador)
        self.boton_eliminar = ttk.Button(self.root, text="Eliminar Trabajador", command=self.eliminar_trabajador)

        self.boton_agregar.pack()
        self.boton_eliminar.pack()

    def bind_input_fields(self):
        self.entry_trabajador.bind("<KeyRelease>", lambda event: self.habilitar_botones())
        self.entry_rut.bind("<KeyRelease>", lambda event: self.habilitar_botones())
        self.entry_cargo.bind("<KeyRelease>", lambda event: self.habilitar_botones())
        self.entry_fecha_contrato.bind("<KeyRelease>", lambda event: self.habilitar_botones())
        self.entry_tipo_anexo.bind("<KeyRelease>", lambda event: self.habilitar_botones())
        self.entry_actualizacion_dt.bind("<KeyRelease>", lambda event: self.habilitar_botones())
        self.entry_fecha_actualizacion_dt.bind("<KeyRelease>", lambda event: self.habilitar_botones())

    def disable_add_button(self):
        self.boton_agregar.config(state=tk.DISABLED)

    def enable_add_button(self):
        self.boton_agregar.config(state=tk.NORMAL)

    def agregar_trabajador(self):
        try:
            trabajador = self.entry_trabajador.get()
            rut = self.entry_rut.get()
            cargo = self.entry_cargo.get()
            fecha_contrato = self.entry_fecha_contrato.get()
            tipo_anexo = self.entry_tipo_anexo.get()
            actualizacion_dt = self.entry_actualizacion_dt.get()
            fecha_actualizacion_dt = self.entry_fecha_actualizacion_dt.get()
            
            categoria_actual = self.listbook.tab(self.listbook.select(), "text")
            
            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()
            
            cursor.execute('''
                INSERT INTO trabajadores (categoria, trabajador, rut, cargo, fecha_contrato, tipo_anexo, actualizacion_dt, fecha_actualizacion_dt)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (categoria_actual, trabajador, rut, cargo, fecha_contrato, tipo_anexo, actualizacion_dt, fecha_actualizacion_dt))
            
            conexion.commit()

            cursor.execute('SELECT last_insert_rowid()')
            id_trabajador = cursor.fetchone()[0]

            conexion.close()
            
            self.clear_input_fields()
            
            self.load_data_for_category(categoria_actual)
        except sqlite3.Error as e:
            self.show_error_message(f"Error al agregar trabajador: {str(e)}")

    def eliminar_trabajador(self):
        try:
            categoria_actual = self.listbook.tab(self.listbook.select(), "text")
            seleccion = self.tablas_trabajadores[categoria_actual].selection()
        
            if seleccion:
                for item in seleccion:
                    id_trabajador = item[0]
                    
                    conexion = sqlite3.connect("trabajadores.db")
                    cursor = conexion.cursor()
                
                    cursor.execute('DELETE FROM trabajadores WHERE categoria = ? AND id = ?', (categoria_actual, id_trabajador))
                    cursor.execute('DELETE FROM datos_adicionales WHERE id_trabajador = ?', (id_trabajador,))
                    conexion.commit()
                    conexion.close()
                
                    self.tablas_trabajadores[categoria_actual].delete(item)
        except sqlite3.Error as e:
            self.show_error_message(f"Error al eliminar trabajador: {str(e)}")

    def load_data_for_category(self, categoria):
        try:
            conexion = sqlite3.connect("trabajadores.db")
            query = f'SELECT id, trabajador, rut, cargo, fecha_contrato, tipo_anexo, actualizacion_dt, fecha_actualizacion_dt FROM trabajadores WHERE categoria = ?'
            df = pd.read_sql_query(query, conexion, params=(categoria,))
            
            tabla = self.tablas_trabajadores[categoria]
            
            tabla.delete(*tabla.get_children())
            
            for _, row in df.iterrows():
                datos = row.tolist()
                tabla.insert('', 'end', values=datos)
            
            conexion.close()
        except sqlite3.Error as e:
            self.show_error_message(f"Error al cargar datos desde la base de datos: {str(e)}")

    def habilitar_botones(self):
        trabajador = self.entry_trabajador.get()
        rut = self.entry_rut.get()
        cargo = self.entry_cargo.get()
        fecha_contrato = self.entry_fecha_contrato.get()
        tipo_anexo = self.entry_tipo_anexo.get()
        actualizacion_dt = self.entry_actualizacion_dt.get()
        fecha_actualizacion_dt = self.entry_fecha_actualizacion_dt.get()
        
        if (
            trabajador.strip() != "" and
            rut.strip() != "" and
            cargo.strip() != "" and
            fecha_contrato.strip() != "" and
            tipo_anexo.strip() != "" and
            actualizacion_dt.strip() != "" and
            fecha_actualizacion_dt.strip() != ""
        ):
            self.enable_add_button()
        else:
            self.disable_add_button()

    def clear_input_fields(self):
        self.entry_trabajador.delete(0, 'end')
        self.entry_rut.delete(0, 'end')
        self.entry_cargo.delete(0, 'end')
        self.entry_fecha_contrato.delete(0, 'end')
        self.entry_tipo_anexo.delete(0, 'end')
        self.entry_actualizacion_dt.delete(0, 'end')
        self.entry_fecha_actualizacion_dt.delete(0, 'end')

    def show_error_message(self, message):
        messagebox.showerror("Error", message)

    def mostrar_menu_contextual(self, event):
        categoria_actual = self.listbook.tab(self.listbook.select(), "text")
        tabla = self.tablas_trabajadores[categoria_actual]

        item_seleccionado = tabla.item(tabla.selection())
    
        if item_seleccionado and 'values' in item_seleccionado:
            id_trabajador = item_seleccionado['values'][0]
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Editar Trabajador", command=lambda: self.editar_trabajador(id_trabajador))
            menu.add_command(label="Agregar/Editar Datos Adicionales", command=lambda: self.agregar_editar_datos_adicionales(id_trabajador))
            menu.add_command(label="Ver Historial de Anexos", command=lambda: self.ver_historial_anexos(id_trabajador))
            menu.add_command(label="Ver Datos Extras", command=lambda: self.ver_datos_extras(id_trabajador))
            menu.post(event.x_root, event.y_root)

    def editar_trabajador(self, id_trabajador):
        try:
            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()
            
            cursor.execute('SELECT trabajador, rut, cargo, fecha_contrato, tipo_anexo, actualizacion_dt, fecha_actualizacion_dt FROM trabajadores WHERE id = ?', (id_trabajador,))
            datos_trabajador = cursor.fetchone()
            
            conexion.close()

            if datos_trabajador:
                ventana_editar = tk.Toplevel(self.root)
                ventana_editar.title("Editar Trabajador")

                frame_editar = ttk.Frame(ventana_editar)
                frame_editar.pack(padx=20, pady=10)

                trabajador_label = ttk.Label(frame_editar, text="Trabajador:")
                rut_label = ttk.Label(frame_editar, text="Rut:")
                cargo_label = ttk.Label(frame_editar, text="Cargo:")
                fecha_contrato_label = ttk.Label(frame_editar, text="Fecha de Contrato:")
                tipo_anexo_label = ttk.Label(frame_editar, text="Tipo Último Anexo:")
                actualizacion_dt_label = ttk.Label(frame_editar, text="Actualización en DT:")
                fecha_actualizacion_dt_label = ttk.Label(frame_editar, text="Fecha Actualización DT:")

                trabajador_entry = ttk.Entry(frame_editar, width=30)
                rut_entry = ttk.Entry(frame_editar, width=30)
                cargo_entry = ttk.Entry(frame_editar, width=30)
                fecha_contrato_entry = ttk.Entry(frame_editar, width=30)
                tipo_anexo_entry = ttk.Entry(frame_editar, width=30)
                actualizacion_dt_entry = ttk.Entry(frame_editar, width=30)
                fecha_actualizacion_dt_entry = ttk.Entry(frame_editar, width=30)

                guardar_button = ttk.Button(frame_editar, text="Guardar Cambios", command=lambda: self.guardar_cambios_trabajador(id_trabajador, trabajador_entry.get(), rut_entry.get(), cargo_entry.get(), fecha_contrato_entry.get(), tipo_anexo_entry.get(), actualizacion_dt_entry.get(), fecha_actualizacion_dt_entry.get(), ventana_editar))

                trabajador_label.grid(row=0, column=0, sticky='w')
                trabajador_entry.grid(row=0, column=1)
                rut_label.grid(row=1, column=0, sticky='w')
                rut_entry.grid(row=1, column=1)
                cargo_label.grid(row=2, column=0, sticky='w')
                cargo_entry.grid(row=2, column=1)
                fecha_contrato_label.grid(row=3, column=0, sticky='w')
                fecha_contrato_entry.grid(row=3, column=1)
                tipo_anexo_label.grid(row=4, column=0, sticky='w')
                tipo_anexo_entry.grid(row=4, column=1)
                actualizacion_dt_label.grid(row=5, column=0, sticky='w')
                actualizacion_dt_entry.grid(row=5, column=1)
                fecha_actualizacion_dt_label.grid(row=6, column=0, sticky='w')
                fecha_actualizacion_dt_entry.grid(row=6, column=1)
                guardar_button.grid(row=7, columnspan=2)

                trabajador_entry.insert(0, datos_trabajador[0])
                rut_entry.insert(0, datos_trabajador[1])
                cargo_entry.insert(0, datos_trabajador[2])
                fecha_contrato_entry.insert(0, datos_trabajador[3])
                tipo_anexo_entry.insert(0, datos_trabajador[4])
                actualizacion_dt_entry.insert(0, datos_trabajador[5])
                fecha_actualizacion_dt_entry.insert(0, datos_trabajador[6])
        except sqlite3.Error as e:
            self.show_error_message(f"Error al editar trabajador: {str(e)}")

    def guardar_cambios_trabajador(self, id_trabajador, trabajador, rut, cargo, fecha_contrato, tipo_anexo, actualizacion_dt, fecha_actualizacion_dt, ventana_editar):
        try:
            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()

            cursor.execute('UPDATE trabajadores SET trabajador = ?, rut = ?, cargo = ?, fecha_contrato = ?, tipo_anexo = ?, actualizacion_dt = ?, fecha_actualizacion_dt = ? WHERE id = ?', (trabajador, rut, cargo, fecha_contrato, tipo_anexo, actualizacion_dt, fecha_actualizacion_dt, id_trabajador))

            conexion.commit()
            conexion.close()

            ventana_editar.destroy()

            categoria_actual = self.listbook.tab(self.listbook.select(), "text")
            self.load_data_for_category(categoria_actual)
        except sqlite3.Error as e:
            self.show_error_message(f"Error al guardar cambios del trabajador: {str(e)}")

    def agregar_editar_datos_adicionales(self, id_trabajador):
        try:
            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()
            
            cursor.execute('SELECT historial_anexos, datos_extras FROM datos_adicionales WHERE id_trabajador = ?', (id_trabajador,))
            datos_adicionales = cursor.fetchone()
            conexion.close()

            historial_anexos = datos_adicionales[0] if datos_adicionales else ""
            datos_extras = datos_adicionales[1] if datos_adicionales else ""

            ventana_datos_adicionales = tk.Toplevel(self.root)
            ventana_datos_adicionales.title("Agregar/Editar Datos Adicionales")

            etiqueta_historial_anexos = ttk.Label(ventana_datos_adicionales, text="Historial de Anexos:")
            etiqueta_historial_anexos.pack()

            historial_anexos_texto = tk.Text(ventana_datos_adicionales, height=10, width=50)
            historial_anexos_texto.insert(tk.END, historial_anexos)
            historial_anexos_texto.pack(fill='both', expand=True)

            etiqueta_datos_extras = ttk.Label(ventana_datos_adicionales, text="Datos Extras:")
            etiqueta_datos_extras.pack()

            datos_extras_texto = tk.Text(ventana_datos_adicionales, height=10, width=50)
            datos_extras_texto.insert(tk.END, datos_extras)
            datos_extras_texto.pack(fill='both', expand=True)

            boton_guardar = ttk.Button(ventana_datos_adicionales, text="Guardar Cambios", command=lambda: self.guardar_datos_adicionales(id_trabajador, historial_anexos_texto.get('1.0', tk.END), datos_extras_texto.get('1.0', tk.END), ventana_datos_adicionales))
            boton_guardar.pack()

        except sqlite3.Error as e:
            self.show_error_message(f"Error al agregar/editar datos adicionales: {str(e)}")

    def guardar_datos_adicionales(self, id_trabajador, historial_anexos, datos_extras, ventana_datos_adicionales):
        try:
            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()

            cursor.execute('INSERT OR REPLACE INTO datos_adicionales (id_trabajador, historial_anexos, datos_extras) VALUES (?, ?, ?)', (id_trabajador, historial_anexos, datos_extras))

            conexion.commit()
            conexion.close()

            ventana_datos_adicionales.destroy()

        except sqlite3.Error as e:
            self.show_error_message(f"Error al guardar datos adicionales: {str(e)}")

    def ver_historial_anexos(self, id_trabajador):
        try:
            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()
            
            cursor.execute('SELECT historial_anexos FROM datos_adicionales WHERE id_trabajador = ?', (id_trabajador,))
            historial_anexos = cursor.fetchone()[0]
            conexion.close()

            ventana_historial_anexos = tk.Toplevel(self.root)
            ventana_historial_anexos.title("Historial de Anexos")

            etiqueta_historial_anexos = ttk.Label(ventana_historial_anexos, text="Historial de Anexos:")
            etiqueta_historial_anexos.pack()

            historial_anexos_texto = tk.Text(ventana_historial_anexos, height=10, width=50)
            historial_anexos_texto.insert(tk.END, historial_anexos)
            historial_anexos_texto.pack(fill='both', expand=True)

            boton_cerrar = ttk.Button(ventana_historial_anexos, text="Cerrar", command=ventana_historial_anexos.destroy)
            boton_cerrar.pack()
        except sqlite3.Error as e:
            self.show_error_message(f"Error al ver historial de anexos: {str(e)}")

    def ver_datos_extras(self, id_trabajador):
        try:
            conexion = sqlite3.connect("trabajadores.db")
            cursor = conexion.cursor()
            
            cursor.execute('SELECT datos_extras FROM datos_adicionales WHERE id_trabajador = ?', (id_trabajador,))
            datos_extras = cursor.fetchone()[0]
            conexion.close()

            ventana_datos_extras = tk.Toplevel(self.root)
            ventana_datos_extras.title("Datos Extras")

            etiqueta_datos_extras = ttk.Label(ventana_datos_extras, text="Datos Extras:")
            etiqueta_datos_extras.pack()

            datos_extras_texto = tk.Text(ventana_datos_extras, height=10, width=50)
            datos_extras_texto.insert(tk.END, datos_extras)
            datos_extras_texto.pack(fill='both', expand=True)

            boton_cerrar = ttk.Button(ventana_datos_extras, text="Cerrar", command=ventana_datos_extras.destroy)
            boton_cerrar.pack()
        except sqlite3.Error as e:
            self.show_error_message(f"Error al ver datos extras: {str(e)}")

    def ver_datos_adicionales(self, event):
        categoria_actual = self.listbook.tab(self.listbook.select(), "text")
        tabla = self.tablas_trabajadores[categoria_actual]
        
        item_seleccionado = tabla.item(tabla.selection())
        id_trabajador = item_seleccionado['values'][0] if item_seleccionado else None

        if id_trabajador:
            try:
                conexion = sqlite3.connect("trabajadores.db")
                cursor = conexion.cursor()

                cursor.execute('SELECT historial_anexos, datos_extras FROM datos_adicionales WHERE id_trabajador = ?', (id_trabajador,))
                datos_adicionales = cursor.fetchone()
                conexion.close()

                historial_anexos = datos_adicionales[0] if datos_adicionales else ""
                datos_extras = datos_adicionales[1] if datos_adicionales else ""

                ventana_datos_adicionales = tk.Toplevel(self.root)
                ventana_datos_adicionales.title("Datos Adicionales")

                etiqueta_historial_anexos = ttk.Label(ventana_datos_adicionales, text="Historial de Anexos:")
                etiqueta_historial_anexos.pack()

                historial_anexos_texto = tk.Text(ventana_datos_adicionales, height=10, width=50)
                historial_anexos_texto.insert(tk.END, historial_anexos)
                historial_anexos_texto.pack(fill='both', expand=True)

                etiqueta_datos_extras = ttk.Label(ventana_datos_adicionales, text="Datos Extras:")
                etiqueta_datos_extras.pack()

                datos_extras_texto = tk.Text(ventana_datos_adicionales, height=10, width=50)
                datos_extras_texto.insert(tk.END, datos_extras)
                datos_extras_texto.pack(fill='both', expand=True)

                boton_cerrar = ttk.Button(ventana_datos_adicionales, text="Cerrar", command=ventana_datos_adicionales.destroy)
                boton_cerrar.pack()
            except sqlite3.Error as e:
                self.show_error_message(f"Error al ver datos adicionales: {str(e)}")

                

if __name__ == "__main__":
    # Crea una ventana tematizada con Clearlooks
    root = ThemedTk(theme="clearlooks")
    app = TrabajadoresApp(root)
    root.mainloop()