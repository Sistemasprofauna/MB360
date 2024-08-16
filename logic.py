import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from zk import ZK
import pandas as pd
import socket

# Función para obtener registros de asistencia
def obtener_asistencia(ip, port, tree, fetch_button):
    zk = ZK(ip, port=port, timeout=5, password=0, force_udp=False, ommit_ping=False)
    conn = None  # Inicializar la variable de conexión

    try:
        # Deshabilitar el botón para evitar múltiples clics
        fetch_button.config(state=tk.DISABLED)
        
        # Validar la dirección IP
        socket.inet_aton(ip)
        
        # Intentar conectar con el dispositivo
        conn = zk.connect()
        messagebox.showinfo("Conexión", "Conexión exitosa")
        
        # Obtener registros de asistencia
        asistencia = conn.get_attendance()
        
        # Limpiar la tabla
        for i in tree.get_children():
            tree.delete(i)
        
        # Crear una lista para almacenar los datos que se exportarán a Excel
        datos = []
        
        # Agregar registros a la tabla y a la lista
        for registro in asistencia:
            tree.insert("", "end", values=(registro.user_id, registro.timestamp, registro.status, registro.punch))
            datos.append([registro.user_id, registro.timestamp, registro.status, registro.punch])
        
        # Convertir la lista a un DataFrame de pandas
        df = pd.DataFrame(datos, columns=["UserID", "Timestamp", "Status","Punch"])
        
        # Guardar los datos en un archivo Excel
        ruta_archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if ruta_archivo:
            df.to_excel(ruta_archivo, index=False)
            messagebox.showinfo("Éxito", f"Registros exportados a {ruta_archivo} con éxito")
    
    except socket.error:
        messagebox.showerror("Error", "Dirección IP inválida.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al conectar o interactuar con el dispositivo: {e}")
    finally:
        # Rehabilitar el botón después de obtener los datos
        fetch_button.config(state=tk.NORMAL)
        
        # Desconectar del dispositivo si la conexión fue exitosa
        if conn:
            conn.disconnect()

# Función para obtener usuarios
def obtener_usuarios(ip, port, tree, fetch_button):
    zk = ZK(ip, port=port, timeout=5, password=0, force_udp=False, ommit_ping=False)
    conn = None  # Inicializar la variable de conexión

    try:
        # Deshabilitar el botón para evitar múltiples clics
        fetch_button.config(state=tk.DISABLED)
        
        # Validar la dirección IP
        socket.inet_aton(ip)
        
        # Intentar conectar con el dispositivo
        conn = zk.connect()
        messagebox.showinfo("Conexión", "Conexión exitosa")
        
        # Obtener la lista de usuarios
        usuarios = conn.get_users()
        
        if not usuarios:
            messagebox.showinfo("Información", "No se encontraron usuarios.")
            return
        
        # Limpiar la tabla
        for i in tree.get_children():
            tree.delete(i)
        
        # Crear una lista para almacenar los datos que se exportarán a Excel
        datos = []
        
        # Agregar usuarios a la tabla y a la lista
        for usuario in usuarios:
            tree.insert("", "end", values=(usuario.user_id, usuario.name, usuario.privilege))
            datos.append([usuario.user_id, usuario.name, usuario.privilege])
        
        # Convertir la lista a un DataFrame de pandas
        df = pd.DataFrame(datos, columns=["UserID", "Name", "Privilege"])
        
        # Guardar los datos en un archivo Excel
        ruta_archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if ruta_archivo:
            df.to_excel(ruta_archivo, index=False)
            messagebox.showinfo("Éxito", f"Usuarios exportados a {ruta_archivo} con éxito")
    
    except socket.error:
        messagebox.showerror("Error", "Dirección IP inválida.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al conectar o interactuar con el dispositivo: {e}")
    finally:
        # Rehabilitar el botón después de obtener los datos
        fetch_button.config(state=tk.NORMAL)
        
        # Desconectar del dispositivo si la conexión fue exitosa
        if conn:
            conn.disconnect()

# Función para obtener el estado del dispositivo
def obtener_estado_dispositivo(ip, port, text_widget, fetch_button):
    zk = ZK(ip, port=port, timeout=5, password=0, force_udp=False, ommit_ping=False)
    conn = None  # Inicializar la variable de conexión

    try:
        # Deshabilitar el botón para evitar múltiples clics
        fetch_button.config(state=tk.DISABLED)
        
        # Validar la dirección IP
        socket.inet_aton(ip)
        
        # Intentar conectar con el dispositivo
        conn = zk.connect()
        messagebox.showinfo("Conexión", "Conexión exitosa")
        
        # Obtener el estado del dispositivo
        estado = conn.get_device_status()
        
        # Mostrar el estado en el widget de texto
        estado_texto = (
            f"Usuarios Registrados: {estado.user_count}\n"
            f"Huellas Dactilares: {estado.finger_count}\n"
            f"Registros de Asistencia: {estado.attendance_count}\n"
        )
        text_widget.config(state=tk.NORMAL)  # Habilitar el widget para editar
        text_widget.delete(1.0, tk.END)  # Limpiar el contenido actual
        text_widget.insert(tk.END, estado_texto)  # Insertar el nuevo estado
        text_widget.config(state=tk.DISABLED)  # Deshabilitar el widget para solo lectura
        
    except socket.error:
        messagebox.showerror("Error", "Dirección IP inválida.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al conectar o interactuar con el dispositivo: {e}")
    finally:
        # Rehabilitar el botón después de obtener los datos
        fetch_button.config(state=tk.NORMAL)
        
        # Desconectar del dispositivo si la conexión fue exitosa
        if conn:
            conn.disconnect()

# Crear la ventana principal
root = tk.Tk()
root.title("Registros de Asistencia ZKTeco MB360")

# Crear y ubicar los widgets
ip_label = tk.Label(root, text="IP del dispositivo:")
ip_label.grid(row=0, column=0, padx=10, pady=10)

ip_entry = tk.Entry(root)
ip_entry.grid(row=0, column=1, padx=10, pady=10)
ip_entry.insert(0, '192.168.19.118')  # Dirección IP por defecto

port_label = tk.Label(root, text="Puerto:")
port_label.grid(row=1, column=0, padx=10, pady=10)

port_entry = tk.Entry(root)
port_entry.grid(row=1, column=1, padx=10, pady=10)
port_entry.insert(0, '4370')  # Puerto por defecto

# Botón para obtener registros de asistencia
fetch_asistencia_button = tk.Button(root, text="Obtener Registros", command=lambda: obtener_asistencia(ip_entry.get(), int(port_entry.get()), tree_asistencia, fetch_asistencia_button))
fetch_asistencia_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# Crear la tabla para mostrar los registros de asistencia
columnas_asistencia = ("UserID", "Timestamp", "Status", "Punch")
tree_asistencia = ttk.Treeview(root, columns=columnas_asistencia, show="headings")
tree_asistencia.heading("UserID", text="UserID")
tree_asistencia.heading("Timestamp", text="Timestamp")
tree_asistencia.heading("Status", text="Status")
tree_asistencia.heading("Punch", text="Punch")
tree_asistencia.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# Botón para obtener usuarios
fetch_usuarios_button = tk.Button(root, text="Obtener Usuarios", command=lambda: obtener_usuarios(ip_entry.get(), int(port_entry.get()), tree_usuarios, fetch_usuarios_button))
fetch_usuarios_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# Crear la tabla para mostrar los usuarios
columnas_usuarios = ("UserID", "Name", "Privilege")
tree_usuarios = ttk.Treeview(root, columns=columnas_usuarios, show="headings")
tree_usuarios.heading("UserID", text="UserID")
tree_usuarios.heading("Name", text="Name")
tree_usuarios.heading("Privilege", text="Privilege")
tree_usuarios.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

# Botón para obtener estado del dispositivo
fetch_estado_button = tk.Button(root, text="Obtener Estado del Dispositivo", command=lambda: obtener_estado_dispositivo(ip_entry.get(), int(port_entry.get()), estado_text, fetch_estado_button))
fetch_estado_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

# Crear el widget de texto para mostrar el estado del dispositivo
estado_text = tk.Text(root, height=6, width=50)
estado_text.grid(row=7, column=0, columnspan=2, padx=10, pady=10)
estado_text.config(state=tk.DISABLED)  # Inicialmente solo lectura

# Iniciar el bucle principal de Tkinter
root.mainloop()
