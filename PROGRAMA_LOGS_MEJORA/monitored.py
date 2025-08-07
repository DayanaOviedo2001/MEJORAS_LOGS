import tkinter as tk
from tkinter import ttk, messagebox
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

SERVIDORES_FILE = "servidores.json"
EMAIL_SENDER = "tucorreo@dominio.com"
EMAIL_PASSWORD = "tu_clave"
EMAIL_DESTINATARIO = "destinatario@dominio.com"
SMTP_SERVER = "smtp.dominio.com"
SMTP_PORT = 587

def cargar_servidores():
    if os.path.exists(SERVIDORES_FILE):
        with open(SERVIDORES_FILE, "r") as f:
            return json.load(f)
    return []

def guardar_servidores(servidores):
    with open(SERVIDORES_FILE, "w") as f:
        json.dump(servidores, f, indent=4)

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

def generar_excel(data, path):
    columnas = [
        "Servidor", "Host", "LocalAddress", "LocalPort",
        "RemoteAddress", "RemotePort", "State", "ProcessName", "Organizacion"
    ]
    df = pd.DataFrame(data, columns=columnas)
    df.to_excel(path, index=False)

def enviar_email_con_adjunto(path):
    msg = EmailMessage()
    msg['From'] = formataddr(("Reporte Conexiones", EMAIL_SENDER))
    msg['To'] = EMAIL_DESTINATARIO
    msg['Subject'] = "📊 Reporte de Conexiones TCP"
    msg.set_content("Adjunto encontrarás el reporte de conexiones TCP generado automáticamente.")

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

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📊 Reporte de Conexiones TCP (Windows)")
        self.geometry("1100x600")
        self.configure(bg="#e6f0ff")
        self.protocol("WM_DELETE_WINDOW", self.on_cerrar)

        self.servidores = cargar_servidores()
        self.datos = []

        self.filtro_loopback = tk.BooleanVar(value=True)
        self.filtro_ipv6_linklocal = tk.BooleanVar(value=True)
        self.filtro_puertos_cero = tk.BooleanVar(value=True)

        # Variables para prefijos a excluir
        self.prefijo_10_70 = tk.BooleanVar(value=True)
        self.prefijo_192_168 = tk.BooleanVar(value=False)
        self.prefijo_172_16 = tk.BooleanVar(value=False)

        self.crear_widgets()

    def crear_widgets(self):
        frame_form = tk.LabelFrame(self, text="Agregar Servidor", bg="#e6f0ff")
        frame_form.pack(padx=10, pady=10, fill="x")

        self.entry_nombre = tk.Entry(frame_form)
        self.entry_host = tk.Entry(frame_form)
        self.entry_usuario = tk.Entry(frame_form)
        self.entry_clave = tk.Entry(frame_form, show="*")

        for i, (label, entry) in enumerate([
            ("Nombre", self.entry_nombre),
            ("Host/IP", self.entry_host),
            ("Usuario", self.entry_usuario),
            ("Clave", self.entry_clave)
        ]):
            tk.Label(frame_form, text=label, bg="#e6f0ff").grid(row=0, column=i*2)
            entry.grid(row=0, column=i*2+1, padx=5, pady=5)

        tk.Button(frame_form, text="➕ Agregar", command=self.agregar_servidor).grid(row=0, column=8, padx=10)

        self.lista_servidores = tk.Listbox(self, height=5)
        self.lista_servidores.pack(padx=10, fill="x")

        tk.Button(self, text="🗑️ Eliminar Seleccionado", command=self.eliminar_servidor).pack(pady=5)

        filtros_frame = tk.LabelFrame(self, text="Filtros de Conexión", bg="#e6f0ff")
        filtros_frame.pack(padx=10, pady=5, fill="x")

        tk.Checkbutton(filtros_frame, text="Omitir loopback (127.0.0.1, 0.0.0.0, ::)", variable=self.filtro_loopback, bg="#e6f0ff").pack(anchor="w")
        tk.Checkbutton(filtros_frame, text="Omitir IPv6 Link-Local (fe80::, %)", variable=self.filtro_ipv6_linklocal, bg="#e6f0ff").pack(anchor="w")
        tk.Checkbutton(filtros_frame, text="Omitir puertos con valor 0", variable=self.filtro_puertos_cero, bg="#e6f0ff").pack(anchor="w")

        # Prefijos IP remotos a excluir
        prefijos_frame = tk.LabelFrame(filtros_frame, text="Excluir conexiones con IP remota que empiecen por:", bg="#e6f0ff")
        prefijos_frame.pack(anchor="w", pady=5, fill="x")

        tk.Checkbutton(prefijos_frame, text="10.70.0.", variable=self.prefijo_10_70, bg="#e6f0ff").pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(prefijos_frame, text="192.168.", variable=self.prefijo_192_168, bg="#e6f0ff").pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(prefijos_frame, text="172.16.", variable=self.prefijo_172_16, bg="#e6f0ff").pack(side=tk.LEFT, padx=5)

        frame_tabla = tk.Frame(self)
        frame_tabla.pack(padx=10, pady=10, fill="both", expand=True)

        columns = (
            "Servidor", "Host", "LocalAddress", "LocalPort",
            "RemoteAddress", "RemotePort", "State", "ProcessName", "Organizacion"
        )
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview.Heading", foreground="black")
        style.configure("Treeview", rowheight=24)
        self.tree = ttk.Treeview(frame_tabla, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")
        self.tree.pack(side=tk.LEFT, fill="both", expand=True)

        scrollbar_v = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree.yview)
        scrollbar_v.pack(side=tk.RIGHT, fill="y")
        self.tree.configure(yscrollcommand=scrollbar_v.set)

        scrollbar_h = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        scrollbar_h.pack(fill="x")
        self.tree.configure(xscrollcommand=scrollbar_h.set)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Copiar celda", command=self.copiar_celda)
        self.tree.bind("<Button-3>", self.mostrar_menu_contextual)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Consultar Conexiones", command=self.iniciar_consulta).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Exportar Excel", command=self.exportar_excel).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Enviar por Email", command=self.enviar_email).pack(side=tk.LEFT, padx=10)

        self.actualizar_lista_servidores()

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

    def eliminar_servidor(self):
        seleccionado = self.lista_servidores.curselection()
        if not seleccionado:
            return
        index = seleccionado[0]
        del self.servidores[index]
        guardar_servidores(self.servidores)
        self.actualizar_lista_servidores()

    def actualizar_lista_servidores(self):
        self.lista_servidores.delete(0, tk.END)
        for s in self.servidores:
            self.lista_servidores.insert(tk.END, f"{s['nombre']} ({s['host']})")

    def iniciar_consulta(self):
        if not self.servidores:
            messagebox.showwarning("Sin servidores", "Agregue al menos un servidor antes de consultar.")
            return
        threading.Thread(target=self.consultar_conexiones).start()

    def consultar_conexiones(self):
        self.datos.clear()
        self.tree.delete(*self.tree.get_children())
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
                self.datos.append(c)

        self.llenar_tabla()

    def llenar_tabla(self):
        self.tree.delete(*self.tree.get_children())
        for fila in self.datos:
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
            self.tree.insert("", tk.END, values=valores)

    def exportar_excel(self):
        if not self.datos:
            messagebox.showinfo("Sin datos", "No hay datos para exportar.")
            return
        fecha_hora = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        archivo = f"Reporte_Conexiones_{fecha_hora}.xlsx"
        try:
            generar_excel(self.datos, archivo)
            messagebox.showinfo("Exportado", f"Archivo guardado como {archivo}")
        except Exception as e:
            messagebox.showerror("Error exportando", str(e))

    def enviar_email(self):
        if not self.datos:
            messagebox.showinfo("Sin datos", "No hay datos para enviar.")
            return
        fecha_hora = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        archivo = f"Reporte_Conexiones_{fecha_hora}.xlsx"
        try:
            generar_excel(self.datos, archivo)
            enviar_email_con_adjunto(archivo)
            messagebox.showinfo("Enviado", "Email enviado con éxito.")
        except Exception as e:
            messagebox.showerror("Error enviando email", str(e))

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

    def on_cerrar(self):
        if messagebox.askokcancel("Salir", "¿Desea salir?"):
            self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()
