import tkinter as tk
import customtkinter as ctk
from tkinter import ttk
from tkinter import messagebox, filedialog
import winrm
import pandas as pd
import datetime
import os
import json
import threading
import requests
import ipaddress
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from email.mime.base import MIMEBase
from email import encoders

# Configuración de CustomTkinter
ctk.set_appearance_mode("light")  # Modos: "System" (por defecto), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Temas: "blue" (por defecto), "green", "dark-blue"

# Configuración común
SERVIDORES_FILE = "servidores.json"
EMAIL_SENDER = "tucorreo@dominio.com"
EMAIL_PASSWORD = "tu_clave"
EMAIL_DESTINATARIO = "destinatario@dominio.com"
SMTP_SERVER = "smtp.dominio.com"
SMTP_PORT = 587

# Configuración uniforme de botones
BUTTON_WIDTH = 200
BUTTON_HEIGHT = 50

# Funciones comunes
def cargar_servidores():
    if os.path.exists(SERVIDORES_FILE):
        with open(SERVIDORES_FILE, "r") as f:
            return json.load(f)
    return []

def guardar_servidores(servidores):
    with open(SERVIDORES_FILE, "w") as f:
        json.dump(servidores, f, indent=4)

# Funciones para el módulo de conexiones TCP
def es_ip_privada(ip):
    try:
        return ipaddress.ip_address(ip).is_private
    except:
        return False

def consulta_arin(ip):
    url = f"https://rdap.arin.net/registry/ip/{ip}"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            entities = data.get("entities", [])
            for entity in entities:
                vcard = entity.get("vcardArray", [])
                if len(vcard) > 1:
                    for item in vcard[1]:
                        if item[0] == "fn":
                            return item[3]  # Nombre completo de la organización
            if "name" in data:
                return data["name"]
            return "Nombre no encontrado"
        else:
            return "IP no encontrada en ARIN"
    except Exception as e:
        return f"Error: {str(e)}"

def obtener_conexiones_red_windows(servidor, filtro_loopback=True, filtro_ipv6=True, filtro_puertos_cero=True, prefijos_excluir=None):
    if prefijos_excluir is None:
        prefijos_excluir = []
    try:
        url = f'http://{servidor["host"]}:5985/wsman'
        session = winrm.Session(url, auth=(servidor["usuario"], servidor["clave"]), transport='ntlm')

        ps_script = """
        Get-NetTCPConnection | ForEach-Object {
            $proc = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                LocalAddress = $_.LocalAddress
                LocalPort = $_.LocalPort
                RemoteAddress = $_.RemoteAddress
                RemotePort = $_.RemotePort
                State = $_.State
                ProcessName = if ($proc) { $proc.ProcessName } else { "N/A" }
            }
        } | ConvertTo-Json -Depth 3
        """

        result = session.run_ps(ps_script)
        output = result.std_out.decode("utf-8").strip()

        if not output:
            return [{
                "Servidor": servidor["nombre"],
                "Host": servidor["host"],
                "LocalAddress": "",
                "LocalPort": "",
                "RemoteAddress": "",
                "RemotePort": "",
                "State": "No se encontraron conexiones o error",
                "ProcessName": ""
            }]

        conexiones = json.loads(output)
        if isinstance(conexiones, dict):
            conexiones = [conexiones]

        resultados = []
        for c in conexiones:
            local = c.get("LocalAddress", "")
            remote = c.get("RemoteAddress", "")
            local_port = c.get("LocalPort", 0)
            remote_port = c.get("RemotePort", 0)

            # Filtros generales
            if filtro_loopback and (local in ["127.0.0.1", "0.0.0.0", "::"] or remote in ["127.0.0.1", "0.0.0.0", "::"]):
                continue
            if filtro_ipv6 and ("fe80::" in local or "fe80::" in remote or "%" in local or "%" in remote):
                continue
            if filtro_puertos_cero and (local_port == 0 or remote_port == 0):
                continue

            # Filtro de prefijos excluidos para RemoteAddress
            if any(remote.startswith(prefijo) for prefijo in prefijos_excluir):
                continue

            resultados.append({
                "Servidor": servidor["nombre"],
                "Host": servidor["host"],
                "LocalAddress": local,
                "LocalPort": local_port,
                "RemoteAddress": remote,
                "RemotePort": remote_port,
                "State": c.get("State", ""),
                "ProcessName": c.get("ProcessName", "")
            })

        return resultados

    except Exception as e:
        return [{
            "Servidor": servidor["nombre"],
            "Host": servidor["host"],
            "LocalAddress": "Error",
            "LocalPort": "",
            "RemoteAddress": "",
            "RemotePort": "",
            "State": str(e),
            "ProcessName": ""
        }]

# Funciones para el módulo de espacio en disco
def obtener_disco_windows(servidor):
    try:
        url = f'http://{servidor["host"]}:5985/wsman'
        session = winrm.Session(url, auth=(servidor["usuario"], servidor["clave"]), transport='ntlm')
        cmd = 'wmic logicaldisk get Caption,FreeSpace,Size,VolumeName'
        result = session.run_cmd(cmd)

        output = result.std_out.decode("utf-8").strip().splitlines()
        headers = output[0].split()
        lines = output[1:]

        resultados = []

        for linea in lines:
            if not linea.strip():
                continue
            partes = linea.strip().split()
            if len(partes) < 3:
                continue

            caption = partes[0]
            free = partes[1]
            size = partes[2]
            volume_name = " ".join(partes[3:]) if len(partes) > 3 else ""

            libre_gb = round(int(free) / (1024**3), 2)
            total_gb = round(int(size) / (1024**3), 2)
            usado_gb = round(total_gb - libre_gb, 2)
            porcentaje = round((usado_gb / total_gb) * 100, 2) if total_gb else 0

            resultados.append({
                "Servidor": servidor["nombre"],
                "Host": servidor["host"],
                "Disco": caption,
                "Nombre Volumen": volume_name,
                "Total (GB)": total_gb,
                "Usado (GB)": usado_gb,
                "Libre (GB)": libre_gb,
                "Uso (%)": f"{porcentaje} %"
            })

        return resultados
    except Exception as e:
        return [{
            "Servidor": servidor["nombre"],
            "Host": servidor["host"],
            "Disco": "Error",
            "Nombre Volumen": "",
            "Total (GB)": "",
            "Usado (GB)": "",
            "Libre (GB)": "",
            "Uso (%)": str(e)
        }]

# Funciones para el módulo de memoria RAM
def obtener_memoria(servidor):
    try:
        url = f'http://{servidor["host"]}:5985/wsman'
        session = winrm.Session(url, auth=(servidor["usuario"], servidor["clave"]), transport='ntlm')
        ps_cmd = "Get-CimInstance Win32_OperatingSystem | Select-Object TotalVisibleMemorySize,FreePhysicalMemory"
        result = session.run_ps(ps_cmd)

        output = result.std_out.decode("utf-8").strip().splitlines()
        if len(output) < 3:
            raise Exception("Formato inesperado")

        valores = output[2].split()
        total_kb = int(valores[0])
        libre_kb = int(valores[1])
        total_gb = round(total_kb / 1024 / 1024, 1)
        libre_gb = round(libre_kb / 1024 / 1024, 1)
        usado_gb = round(total_gb - libre_gb, 1)
        porcentaje = round((usado_gb / total_gb) * 100)

        fecha_hora = datetime.datetime.now().strftime("%#d/%#m/%Y %#H:%M")
        resumen_formateado = f"Servidor: {servidor['nombre']}\n{fecha_hora}\t{usado_gb}/{total_gb} GB ({porcentaje}%)"

        return {
            "Servidor": servidor["nombre"],
            "Host": servidor["host"],
            "Usado (GB)": usado_gb,
            "Total (GB)": total_gb,
            "Porcentaje (%)": f"{porcentaje}%",
            "Resumen": resumen_formateado
        }

    except Exception as e:
        fecha_hora = datetime.datetime.now().strftime("%#d/%#m/%Y %#H:%M")
        return {
            "Servidor": servidor["nombre"],
            "Host": servidor["host"],
            "Usado (GB)": "",
            "Total (GB)": "",
            "Porcentaje (%)": "",
            "Resumen": f"Servidor: {servidor['nombre']}\n{fecha_hora}\tError: {str(e)}"
        }

# Funciones para generar reportes y enviar emails
def generar_excel(data, path, tipo):
    df = pd.DataFrame(data)
    if tipo == "conexiones":
        columnas = [
            "Servidor", "Host", "LocalAddress", "LocalPort",
            "RemoteAddress", "RemotePort", "State", "ProcessName", "Organizacion"
        ]
        df = pd.DataFrame(data, columns=columnas)
    elif tipo == "disco":
        columnas_orden = ["Servidor", "Host", "Disco", "Nombre Volumen", "Total (GB)", "Usado (GB)", "Libre (GB)", "Uso (%)"]
        df = df[columnas_orden]
    elif tipo == "memoria":
        df.drop(columns=["Resumen"], inplace=True, errors='ignore')
    
    df.to_excel(path, index=False)

def enviar_email_con_adjunto(path, tipo):
    msg = EmailMessage()
    
    if tipo == "conexiones":
        msg['From'] = formataddr(("Reporte Conexiones", EMAIL_SENDER))
        msg['Subject'] = "📊 Reporte de Conexiones TCP"
        msg.set_content("Adjunto encontrarás el reporte de conexiones TCP generado automáticamente.")
    elif tipo == "disco":
        msg['From'] = formataddr(("Reporte Disco", EMAIL_SENDER))
        msg['Subject'] = "📊 Reporte de Espacio en Disco"
        msg.set_content("Adjunto encontrarás el reporte de espacio en disco generado automáticamente.")
    elif tipo == "memoria":
        msg['From'] = formataddr(("Reporte Memoria", EMAIL_SENDER))
        msg['Subject'] = "📊 Reporte de Uso de Memoria RAM"
        msg.set_content("Adjunto encontrarás el reporte de uso de memoria RAM generado automáticamente.")
    
    msg['To'] = EMAIL_DESTINATARIO

    with open(path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(path)

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(file_data)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
    msg.add_attachment(part.get_payload(decode=True), maintype='application', subtype='octet-stream', filename=file_name)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

def enviar_email_memoria(resumenes, excel_path):
    msg = EmailMessage()
    msg['From'] = formataddr(("Reporte Memoria", EMAIL_SENDER))
    msg['To'] = EMAIL_DESTINATARIO
    msg['Subject'] = "📊 Reporte de Uso de Memoria RAM"

    cuerpo = "\n\n".join(resumenes)
    msg.set_content(f"Resumen de uso de memoria RAM:\n\n{cuerpo}\n\nAutomatizado.")

    with open(excel_path, "rb") as f:
        archivo = f.read()
        file_name = os.path.basename(excel_path)

    part = MIMEBase("application", "octet-stream")
    part.set_payload(archivo)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{file_name}"')
    msg.add_attachment(part.get_payload(decode=True), maintype="application", subtype="octet-stream", filename=file_name)

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

# Clase para crear tablas personalizadas
class ScrollableTreeView(ctk.CTkFrame):
    def __init__(self, master, columns, height=400, **kwargs):
        super().__init__(master, **kwargs)
        
        # Crear un frame para contener el Treeview y las barras de desplazamiento
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True)
        
        # Crear el Treeview con estilo personalizado
        style = ttk.Style()
        for candidate in ("clam", "default", "alt"):
            try:
                style.theme_use(candidate)
                break
            except Exception:
                continue
        style_name = "Monit.Treeview"
        style.configure(style_name, background="#E0F0FF", foreground="#333333", rowheight=25, fieldbackground="#E0F0FF")
        style.configure(f"{style_name}.Heading", background="#D9D9D9", foreground="#000000", relief="flat", font=("Segoe UI", 10, "bold"), anchor="center")
        style.configure("Treeview.Heading", background="#D9D9D9", foreground="#000000", relief="flat", font=("Segoe UI", 10, "bold"), anchor="center")
        style.map(f"{style_name}.Heading",
                  background=[("pressed", "#D9D9D9"), ("active", "#E5E5E5"), ("!active", "#D9D9D9")],
                  foreground=[("pressed", "#000000"), ("active", "#000000"), ("!disabled", "#000000")])
        style.map("Treeview.Heading",
                  background=[("pressed", "#D9D9D9"), ("active", "#E5E5E5"), ("!active", "#D9D9D9")],
                  foreground=[("pressed", "#000000"), ("active", "#000000"), ("!disabled", "#000000")])
        style.map(style_name, background=[["selected", "#4B8BBE"]])
        style.layout("Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.text", {"sticky": "we"})
                ]})
            ]})
        ])
        
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings", height=height, style=style_name)
        
        # Configurar encabezados y columnas
        for col in columns:
            self.tree.heading(col, text=str(col), anchor="center")
            self.tree.column(col, width=100, anchor="center")
        
        # Crear barras de desplazamiento
        self.vsb = ctk.CTkScrollbar(self.tree_frame, orientation="vertical", command=self.tree.yview)
        self.hsb = ctk.CTkScrollbar(self, orientation="horizontal", command=self.tree.xview)
        
        # Configurar el Treeview para usar las barras de desplazamiento
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        
        # Colocar los widgets en el frame
        self.tree.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.hsb.pack(fill="x")
        
        # Crear menú contextual
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Copiar celda", command=self.copiar_celda)
        self.tree.bind("<Button-3>", self.mostrar_menu_contextual)
    
    def insert(self, parent, index, values):
        return self.tree.insert(parent, index, values=values)
    
    def delete(self, *items):
        self.tree.delete(*items)
    
    def get_children(self):
        return self.tree.get_children()
    
    def copiar_celda(self):
        item = self.tree.selection()
        if not item:
            return
        col = self.tree.identify_column(self.tree.winfo_pointerx() - self.tree.winfo_rootx())
        if not col:
            return
        col_index = int(col.replace("#", "")) - 1
        valores = self.tree.item(item, "values")
        if col_index < len(valores):
            self.clipboard_clear()
            self.clipboard_append(valores[col_index])
    
    def mostrar_menu_contextual(self, event):
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

# Clase principal de la aplicación
class AppIntegrada(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🖥️ Herramienta de Monitoreo de Servidores")
        self.geometry("1200x700")  # Aumenté un poco el tamaño
        self.configure(fg_color="#E0F0FF")  # Color de fondo celeste claro
        self.protocol("WM_DELETE_WINDOW", self.on_cerrar)

        self.servidores = cargar_servidores()
        
        # Crear pestañas con estilo personalizado
        self.tabview = ctk.CTkTabview(self, fg_color="#E0F0FF")
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Crear pestañas
        self.tab_servidores = self.tabview.add("Servidores")
        self.tab_conexiones = self.tabview.add("Conexiones TCP")
        self.tab_disco = self.tabview.add("Espacio en Disco")
        self.tab_memoria = self.tabview.add("Memoria RAM")
        
        # Inicializar componentes de cada pestaña
        self.setup_tab_servidores()
        self.setup_tab_conexiones()
        self.setup_tab_disco()
        self.setup_tab_memoria()
    
    def setup_tab_servidores(self):
        # Frame para agregar servidores
        frame_form = ctk.CTkFrame(self.tab_servidores)
        frame_form.pack(padx=10, pady=10, fill="x")
        
        ctk.CTkLabel(frame_form, text="Gestión de Servidores", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        # Crear grid para los campos de entrada
        form_grid = ctk.CTkFrame(frame_form, fg_color="transparent")
        form_grid.pack(padx=20, pady=10, fill="x")
        
        # Etiquetas y campos de entrada
        ctk.CTkLabel(form_grid, text="Nombre:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_nombre = ctk.CTkEntry(form_grid, width=200)
        self.entry_nombre.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(form_grid, text="Host/IP:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.entry_host = ctk.CTkEntry(form_grid, width=200)
        self.entry_host.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(form_grid, text="Usuario:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_usuario = ctk.CTkEntry(form_grid, width=200)
        self.entry_usuario.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(form_grid, text="Clave:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.entry_clave = ctk.CTkEntry(form_grid, width=200, show="*")
        self.entry_clave.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        
        # Botón para agregar servidor
        btn_frame = ctk.CTkFrame(frame_form, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        ctk.CTkButton(btn_frame, text="➕ Agregar Servidor", command=self.agregar_servidor,
                      width=BUTTON_WIDTH, height=BUTTON_HEIGHT).pack(side="left", padx=10)
        
        # Lista de servidores
        list_frame = ctk.CTkFrame(self.tab_servidores)
        list_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        ctk.CTkLabel(list_frame, text="Servidores Configurados", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        # Crear un frame para la lista con borde
        self.lista_frame = ctk.CTkFrame(list_frame, fg_color="#FFFFFF", border_width=1, border_color="#CCCCCC")
        self.lista_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        # Usar un Listbox de tkinter con estilo personalizado
        self.lista_servidores = tk.Listbox(self.lista_frame, bg="#FFFFFF", fg="#333333", font=("Segoe UI", 11), 
                                          selectbackground="#4B8BBE", height=10, bd=0, highlightthickness=0)
        self.lista_servidores.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        
        # Barra de desplazamiento para la lista
        scrollbar = ctk.CTkScrollbar(self.lista_frame, command=self.lista_servidores.yview)
        scrollbar.pack(side="right", fill="y")
        self.lista_servidores.config(yscrollcommand=scrollbar.set)
        
        # Botón para eliminar servidor
        ctk.CTkButton(list_frame, text="🗑️ Eliminar Seleccionado", 
                     fg_color="#FF5252", hover_color="#FF0000",
                     command=self.eliminar_servidor,
                     width=BUTTON_WIDTH, height=BUTTON_HEIGHT).pack(pady=10)
        
        self.actualizar_lista_servidores()
    
    def setup_tab_conexiones(self):
        # Crear un frame principal con scroll para toda la pestaña
        main_frame = ctk.CTkScrollableFrame(self.tab_conexiones)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Variables para filtros
        self.filtro_loopback = tk.BooleanVar(value=True)
        self.filtro_ipv6_linklocal = tk.BooleanVar(value=True)
        self.filtro_puertos_cero = tk.BooleanVar(value=True)
        self.prefijo_10_70 = tk.BooleanVar(value=True)
        self.prefijo_192_168 = tk.BooleanVar(value=False)
        self.prefijo_172_16 = tk.BooleanVar(value=False)
        
        # Frame para filtros
        filtros_frame = ctk.CTkFrame(main_frame)
        filtros_frame.pack(pady=(0, 10), fill="x")
        
        ctk.CTkLabel(filtros_frame, text="Filtros de Conexión", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        # Checkboxes para filtros
        check_frame = ctk.CTkFrame(filtros_frame, fg_color="transparent")
        check_frame.pack(padx=20, pady=5, fill="x")
        
        ctk.CTkCheckBox(check_frame, text="Omitir loopback (127.0.0.1, 0.0.0.0, ::)", 
                       variable=self.filtro_loopback, onvalue=True, offvalue=False).pack(anchor="w", pady=2)
        ctk.CTkCheckBox(check_frame, text="Omitir IPv6 Link-Local (fe80::, %)", 
                       variable=self.filtro_ipv6_linklocal, onvalue=True, offvalue=False).pack(anchor="w", pady=2)
        ctk.CTkCheckBox(check_frame, text="Omitir puertos con valor 0", 
                       variable=self.filtro_puertos_cero, onvalue=True, offvalue=False).pack(anchor="w", pady=2)
        
        # Prefijos IP remotos a excluir
        prefijos_label = ctk.CTkLabel(check_frame, text="Excluir conexiones con IP remota que empiecen por:")
        prefijos_label.pack(anchor="w", pady=(10, 5))
        
        prefijos_frame = ctk.CTkFrame(check_frame, fg_color="transparent")
        prefijos_frame.pack(anchor="w", fill="x")
        
        ctk.CTkCheckBox(prefijos_frame, text="10.70.0.", 
                       variable=self.prefijo_10_70, onvalue=True, offvalue=False).pack(side="left", padx=10)
        ctk.CTkCheckBox(prefijos_frame, text="192.168.", 
                       variable=self.prefijo_192_168, onvalue=True, offvalue=False).pack(side="left", padx=10)
        ctk.CTkCheckBox(prefijos_frame, text="172.16.", 
                       variable=self.prefijo_172_16, onvalue=True, offvalue=False).pack(side="left", padx=10)
        
        # Botón de consulta
        consulta_frame = ctk.CTkFrame(main_frame)
        consulta_frame.pack(pady=10, fill="x")
        
        consultar_btn = ctk.CTkButton(consulta_frame, text="🔍 Consultar Conexiones", 
                     command=self.consultar_conexiones,
                     font=ctk.CTkFont(size=14, weight="bold"),
                     fg_color="#2E74B5",
                     hover_color="#1A5085",
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT)
        consultar_btn.pack(pady=10)
        
        # Tabla de conexiones - con altura fija para evitar problemas
        tabla_frame = ctk.CTkFrame(main_frame)
        tabla_frame.pack(pady=10, fill="x")
        
        columns = (
            "Servidor", "Host", "LocalAddress", "LocalPort",
            "RemoteAddress", "RemotePort", "State", "ProcessName", "Organizacion"
        )
        
        self.tree_conexiones = ScrollableTreeView(tabla_frame, columns=columns, height=15)
        self.tree_conexiones.pack(fill="x", padx=5, pady=5)
        
        # Botones de exportación y envío
        btn_frame = ctk.CTkFrame(main_frame)
        btn_frame.pack(pady=10, fill="x")
        
        btn_container = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_container.pack(pady=10, fill="x")
        
        ctk.CTkButton(btn_container, text="📊 Exportar Excel", 
                     command=lambda: self.exportar_excel("conexiones"),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT,
                     fg_color="#4CAF50",
                     hover_color="#388E3C").pack(side="left", padx=15, pady=5)
        
        ctk.CTkButton(btn_container, text="📧 Enviar por Email", 
                     command=lambda: self.enviar_email("conexiones"),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT,
                     fg_color="#2196F3",
                     hover_color="#1976D2").pack(side="left", padx=15, pady=5)
        
        # Inicializar datos
        self.datos_conexiones = []
    
    def setup_tab_disco(self):
        # Crear un frame principal con scroll para toda la pestaña
        main_frame = ctk.CTkScrollableFrame(self.tab_disco)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Título
        ctk.CTkLabel(main_frame, text="Espacio en Disco", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(5, 10))
        
        # Barra de progreso
        self.progress_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        self.progress_frame.pack(fill="x", pady=(0, 10))
        
        self.progress_label = ctk.CTkLabel(self.progress_frame, text="")
        self.progress_label.pack(anchor="w", pady=(0, 5))
        
        self.progress_disco = ctk.CTkProgressBar(self.progress_frame)
        self.progress_disco.pack(fill="x")
        self.progress_disco.set(0)  # Inicialmente en 0
        
        # Botón principal de consulta
        consulta_frame = ctk.CTkFrame(main_frame)
        consulta_frame.pack(pady=10, fill="x")
        
        consultar_btn = ctk.CTkButton(consulta_frame, text="🔍 Consultar Servidores", 
                     command=self.consultar_discos,
                     font=ctk.CTkFont(size=14, weight="bold"),
                     fg_color="#2E74B5",
                     hover_color="#1A5085",
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT)
        consultar_btn.pack(pady=10)
        
        # Tabla de discos - con altura fija
        tabla_frame = ctk.CTkFrame(main_frame)
        tabla_frame.pack(pady=10, fill="x")
        
        columns = ("Servidor", "Host", "Disco", "Nombre Volumen", "Total (GB)", "Usado (GB)", "Libre (GB)", "Uso (%)")
        self.tabla_disco = ScrollableTreeView(tabla_frame, columns=columns, height=15)
        self.tabla_disco.pack(fill="x", padx=5, pady=5)
        
        # Botones de acciones
        btn_frame = ctk.CTkFrame(main_frame)
        btn_frame.pack(pady=10, fill="x")
        
        # Primera fila de botones
        btn_container1 = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_container1.pack(pady=10, fill="x")
        
        ctk.CTkButton(btn_container1, text="📊 Generar Excel", 
                     command=lambda: self.exportar_excel("disco"),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT,
                     fg_color="#4CAF50",
                     hover_color="#388E3C").pack(side="left", padx=15, pady=5)
        
        ctk.CTkButton(btn_container1, text="📧 Generar y Enviar Email", 
                     command=lambda: self.enviar_email("disco"),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT,
                     fg_color="#2196F3",
                     hover_color="#1976D2").pack(side="left", padx=15, pady=5)
        
        # Segunda fila de botones
        btn_container2 = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_container2.pack(pady=(0, 15), fill="x")
        
        ctk.CTkButton(btn_container2, text="🧹 Limpiar Tabla", 
                     command=lambda: self.limpiar_tabla("disco"),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT,
                     fg_color="#FF9800",
                     hover_color="#F57C00").pack(side="left", padx=15, pady=5)
        
        ctk.CTkButton(btn_container2, text="📝 Mostrar como texto", 
                     command=self.mostrar_texto_formateado_disco,
                     font=ctk.CTkFont(size=14, weight="bold"),
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT,
                     fg_color="#9C27B0",
                     hover_color="#7B1FA2").pack(side="left", padx=15, pady=5)
        
        # Inicializar datos
        self.datos_disco = []
    
    def setup_tab_memoria(self):
        # Crear un frame principal con scroll para toda la pestaña
        main_frame = ctk.CTkScrollableFrame(self.tab_memoria)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Título
        ctk.CTkLabel(main_frame, text="Reporte de Memoria RAM", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(5, 10))
        
        # Botón principal de consulta
        consulta_frame = ctk.CTkFrame(main_frame)
        consulta_frame.pack(pady=10, fill="x")
        
        consultar_btn = ctk.CTkButton(consulta_frame, text="🔍 Consultar Memoria", 
                     command=self.consultar_memoria,
                     font=ctk.CTkFont(size=14, weight="bold"),
                     fg_color="#2E74B5",
                     hover_color="#1A5085",
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT)
        consultar_btn.pack(pady=10)
        
        # Frame para el área de texto - altura fija
        text_frame = ctk.CTkFrame(main_frame, fg_color="#FFFFFF", border_width=1, border_color="#CCCCCC")
        text_frame.pack(fill="x", pady=10)
        
        # Área de texto con estilo y altura específica
        self.texto_memoria = ctk.CTkTextbox(text_frame, font=ctk.CTkFont(family="Consolas", size=12), 
                                          fg_color="#FFFFFF", text_color="#333333", corner_radius=0,
                                          height=300)  # Altura fija
        self.texto_memoria.pack(fill="x", padx=5, pady=5)
        
        # Botones de acciones
        btn_frame = ctk.CTkFrame(main_frame)
        btn_frame.pack(pady=10, fill="x")
        
        btn_container = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_container.pack(pady=15, fill="x")
        
        ctk.CTkButton(btn_container, text="💾 Guardar Excel", 
                     command=lambda: self.exportar_excel("memoria"),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT,
                     fg_color="#4CAF50",
                     hover_color="#388E3C").pack(side="left", padx=15, pady=5)
        
        ctk.CTkButton(btn_container, text="📧 Enviar Email", 
                     command=lambda: self.enviar_email("memoria"),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     width=BUTTON_WIDTH,
                     height=BUTTON_HEIGHT,
                     fg_color="#2196F3",
                     hover_color="#1976D2").pack(side="left", padx=15, pady=5)
        
        # Inicializar datos
        self.datos_memoria = []
        self.resumenes_memoria = []
    
    # Funciones comunes para gestión de servidores
    def agregar_servidor(self):
        nombre = self.entry_nombre.get().strip()
        host = self.entry_host.get().strip()
        usuario = self.entry_usuario.get().strip()
        clave = self.entry_clave.get().strip()
        if not nombre or not host or not usuario or not clave:
            messagebox.showwarning("Campos vacíos", "Por favor, complete todos los campos.")
            return
        self.servidores.append({
            "nombre": nombre,
            "host": host,
            "usuario": usuario,
            "clave": clave
        })
        guardar_servidores(self.servidores)
        self.actualizar_lista_servidores()
        self.entry_nombre.delete(0, tk.END)
        self.entry_host.delete(0, tk.END)
        self.entry_usuario.delete(0, tk.END)
        self.entry_clave.delete(0, tk.END)
        messagebox.showinfo("Servidor agregado", f"El servidor '{nombre}' ha sido agregado correctamente.")

    def eliminar_servidor(self):
        seleccionado = self.lista_servidores.curselection()
        if not seleccionado:
            messagebox.showwarning("Sin selección", "Por favor, seleccione un servidor para eliminar.")
            return
        index = seleccionado[0]
        nombre_servidor = self.servidores[index]["nombre"]
        if messagebox.askyesno("Confirmar eliminación", f"¿Está seguro de eliminar el servidor '{nombre_servidor}'?"):
            del self.servidores[index]
            guardar_servidores(self.servidores)
            self.actualizar_lista_servidores()
            messagebox.showinfo("Servidor eliminado", f"El servidor '{nombre_servidor}' ha sido eliminado.")

    def actualizar_lista_servidores(self):
        self.lista_servidores.delete(0, tk.END)
        for s in self.servidores:
            self.lista_servidores.insert(tk.END, f"{s['nombre']} ({s['host']})")
    
    # Funciones para la pestaña de conexiones TCP
    def consultar_conexiones(self):
        if not self.servidores:
            messagebox.showwarning("Sin servidores", "Agregue al menos un servidor antes de consultar.")
            return
        threading.Thread(target=self.thread_consultar_conexiones).start()

    def thread_consultar_conexiones(self):
        self.datos_conexiones = []
        self.tree_conexiones.delete(*self.tree_conexiones.get_children())
        prefijos_excluir = []
        if self.prefijo_10_70.get():
            prefijos_excluir.append("10.70.0.")
        if self.prefijo_192_168.get():
            prefijos_excluir.append("192.168.")
        if self.prefijo_172_16.get():
            prefijos_excluir.append("172.16.")

        for servidor in self.servidores:
            conexiones = obtener_conexiones_red_windows(
                servidor,
                filtro_loopback=self.filtro_loopback.get(),
                filtro_ipv6=self.filtro_ipv6_linklocal.get(),
                filtro_puertos_cero=self.filtro_puertos_cero.get(),
                prefijos_excluir=prefijos_excluir
            )
            for c in conexiones:
                org = ""
                remote_ip = c.get("RemoteAddress", "")
                if remote_ip and not es_ip_privada(remote_ip):
                    org = consulta_arin(remote_ip)
                c["Organizacion"] = org
                self.datos_conexiones.append(c)

        self.llenar_tabla_conexiones()

    def llenar_tabla_conexiones(self):
        self.tree_conexiones.delete(*self.tree_conexiones.get_children())
        for fila in self.datos_conexiones:
            valores = (
                fila.get("Servidor", ""),
                fila.get("Host", ""),
                fila.get("LocalAddress", ""),
                fila.get("LocalPort", ""),
                fila.get("RemoteAddress", ""),
                fila.get("RemotePort", ""),
                fila.get("State", ""),
                fila.get("ProcessName", ""),
                fila.get("Organizacion", "")
            )
            self.tree_conexiones.insert("", tk.END, values=valores)
    
    # Funciones para la pestaña de espacio en disco
    def consultar_discos(self):
        if not self.servidores:
            messagebox.showwarning("Sin servidores", "Agregue al menos un servidor antes de consultar.")
            return
        self.progress_disco.set(0)
        self.progress_disco.start()
        self.progress_label.configure(text="Consultando servidores...")
        threading.Thread(target=self.thread_consultar_discos).start()

    def thread_consultar_discos(self):
        self.tabla_disco.delete(*self.tabla_disco.get_children())
        self.datos_disco = []
        total_servidores = len(self.servidores)
        
        for i, srv in enumerate(self.servidores):
            # Actualizar progreso
            progress = (i / total_servidores)
            self.progress_disco.set(progress)
            self.progress_label.configure(text=f"Consultando {srv['nombre']} ({i+1}/{total_servidores})...")
            
            resultados = obtener_disco_windows(srv)
            self.datos_disco.extend(resultados)
            for fila in resultados:
                self.tabla_disco.insert("", tk.END, values=(
                    fila["Servidor"],
                    fila["Host"],
                    fila["Disco"],
                    fila["Nombre Volumen"],
                    fila["Total (GB)"],
                    fila["Usado (GB)"],
                    fila["Libre (GB)"],
                    fila["Uso (%)"]
                ))
        
        self.progress_disco.set(1)  # Completado
        self.progress_label.configure(text="Consulta completada")
    
    def mostrar_texto_formateado_disco(self):
        if not self.datos_disco:
            messagebox.showwarning("Sin datos", "Primero consulta los servidores.")
            return

        texto = ""
        servidores = {}
        fecha_formateada = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

        for fila in self.datos_disco:
            srv = fila["Servidor"]
            if srv not in servidores:
                servidores[srv] = []
            servidores[srv].append(fila)

        for servidor, discos in servidores.items():
            texto += f"Servidor {discos[0]['Host']}\n\n"
            for d in discos:
                nombre_volumen = d['Nombre Volumen'] if d['Nombre Volumen'] else d['Disco']
                linea = f"{fecha_formateada}  {nombre_volumen:<25}\t{d['Libre (GB)']} GB disponible de {d['Total (GB)']} GB"
                texto += linea + "\n"
            texto += "\n"

        # Crear ventana emergente con estilo CustomTkinter
        ventana_texto = ctk.CTkToplevel(self)
        ventana_texto.title("Reporte en texto formateado")
        ventana_texto.geometry("800x600")
        ventana_texto.grab_set()  # Hacer modal
        
        # Área de texto con estilo
        text_widget = ctk.CTkTextbox(ventana_texto, font=ctk.CTkFont(family="Consolas", size=12))
        text_widget.pack(padx=10, pady=10, fill="both", expand=True)
        text_widget.insert("1.0", texto)
        text_widget.configure(state="disabled")  # Hacer de solo lectura
        
        # Botón para cerrar
        ctk.CTkButton(ventana_texto, text="Cerrar", command=ventana_texto.destroy,
                      width=BUTTON_WIDTH, height=BUTTON_HEIGHT).pack(pady=10)
    
    # Funciones para la pestaña de memoria RAM
    def consultar_memoria(self):
        if not self.servidores:
            messagebox.showwarning("Sin servidores", "Agregue al menos un servidor antes de consultar.")
            return
        threading.Thread(target=self.thread_consultar_memoria).start()

    def thread_consultar_memoria(self):
        self.datos_memoria = []
        self.resumenes_memoria = []
        self.texto_memoria.delete("1.0", tk.END)

        for srv in self.servidores:
            resultado = obtener_memoria(srv)
            self.datos_memoria.append(resultado)
            self.resumenes_memoria.append(resultado["Resumen"])
            self.texto_memoria.insert(tk.END, resultado["Resumen"] + "\n\n")
    
    # Funciones comunes para todas las pestañas
    def exportar_excel(self, tipo):
        if tipo == "conexiones":
            datos = self.datos_conexiones
            nombre_archivo = f"Reporte_Conexiones_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        elif tipo == "disco":
            datos = self.datos_disco
            nombre_archivo = f"Reporte_Espacio_Disco_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        elif tipo == "memoria":
            datos = self.datos_memoria
            nombre_archivo = f"Reporte_Memoria_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        if not datos:
            messagebox.showinfo("Sin datos", "No hay datos para exportar.")
            return
        
        # Permitir al usuario elegir dónde guardar el archivo
        ruta_archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=nombre_archivo
        )
        
        if not ruta_archivo:  # Si el usuario cancela
            return
        
        try:
            generar_excel(datos, ruta_archivo, tipo)
            messagebox.showinfo("Exportado", f"Archivo guardado como {ruta_archivo}")
        except Exception as e:
            messagebox.showerror("Error exportando", str(e))

    def enviar_email(self, tipo):
        if tipo == "conexiones":
            datos = self.datos_conexiones
            nombre_archivo = f"Reporte_Conexiones_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        elif tipo == "disco":
            datos = self.datos_disco
            nombre_archivo = f"Reporte_Espacio_Disco_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        elif tipo == "memoria":
            datos = self.datos_memoria
            nombre_archivo = f"Reporte_Memoria_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        if not datos:
            messagebox.showinfo("Sin datos", "No hay datos para enviar.")
            return
        
        try:
            generar_excel(datos, nombre_archivo, tipo)
            if tipo == "memoria":
                enviar_email_memoria(self.resumenes_memoria, nombre_archivo)
            else:
                enviar_email_con_adjunto(nombre_archivo, tipo)
            messagebox.showinfo("Enviado", "Email enviado con éxito.")
        except Exception as e:
            messagebox.showerror("Error enviando email", str(e))
    
    def limpiar_tabla(self, tipo):
        if tipo == "conexiones":
            self.tree_conexiones.delete(*self.tree_conexiones.get_children())
            self.datos_conexiones = []
        elif tipo == "disco":
            self.tabla_disco.delete(*self.tabla_disco.get_children())
            self.datos_disco = []
        elif tipo == "memoria":
            self.texto_memoria.delete("1.0", tk.END)
            self.datos_memoria = []
            self.resumenes_memoria = []

    def on_cerrar(self):
        if messagebox.askokcancel("Salir", "¿Desea salir?"):
            self.destroy()

if __name__ == "__main__":
    app = AppIntegrada()
    app.mainloop()