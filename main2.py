import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from zk import ZK, const
import pandas as pd
import socket

def obtener_asistencia():
    ip = ip_entry.get()
    port = int(port_entry.get())
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
        
        # Convertir la lista a un DataFrame de pandas o numpy
        df = pd.DataFrame(datos, columns=["UserID", "Timestamp", "Status","Punch"])
        
        # Guardar los datos en un archivo Excel
        ruta_archivo = "registros_asistencia.xlsx"
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

def obtener_usuarios():
    ip = ip_entry.get()
    port = int(port_entry.get())
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
        
        # Limpiar la tabla
        for i in tree.get_children():
            tree.delete(i)
        
        # Crear una lista para almacenar los datos que se exportarán a Excel
        datos = []
        
        # Agregar usuarios a la tabla y a la lista
        for usuario in usuarios:
           
            tree.insert("", "end", values=(usuario.user_id, usuario.name, usuario.privilege))
            datos.append([usuario.user_id, usuario.name, usuario.privilege])
        
        # Convertir la lista a un DataFrame de pandas, se prodria hacer en numpy  pero es al gusto
        df = pd.DataFrame(datos, columns=["UserID", "Name", "Privilege"])
        
        # Guardar los datos en un archivo Excel, falta agregar directorio
        ruta_archivo = "usuarios_registrados.xlsx"
        df.to_excel(ruta_archivo, index=False)
        
        messagebox.showinfo("Éxito", f"Usuarios exportados a {ruta_archivo} con éxito")
    
    except socket.error:
        messagebox.showerror("Error", "Dirección IP inválida.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al conectar o interactuar con el dispositivo: {e}")
    finally:
        # Rehabilitar el botón después de obtener los datos, aun falta doble validacion
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

fetch_button= tk.Button(root, text="Obtener Registros", command=obtener_asistencia)
fetch_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

fetch_button_usr = tk.Button(root, text="Obtener usuarios", command=obtener_usuarios)
fetch_button_usr.grid(row=2, column=1, columnspan=2, padx=10, pady=10)

# Crear la tabla para mostrar los registros
columnas = ("UserID", "Timestamp", "Status","Punch")
tree = ttk.Treeview(root, columns=columnas, show="headings")
tree.heading("UserID", text="UserID")
tree.heading("Timestamp", text="Timestamp")
tree.heading("Status", text="Status")
tree.heading("Punch", text="Punch")




tree.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# Iniciar el bucle principal de Tkinter
root.mainloop()
