import tkinter as tk
from tkinter import ttk, messagebox
import winrm
import pandas as pd
import datetime
import os
import json
import threading
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

def generar_excel(data, path):
    df = pd.DataFrame(data)
    columnas_orden = ["Servidor", "Host", "Disco", "Nombre Volumen", "Total (GB)", "Usado (GB)", "Libre (GB)", "Uso (%)"]
    df = df[columnas_orden]
    df.to_excel(path, index=False)

def enviar_email_con_adjunto(path):
    msg = EmailMessage()
    msg['From'] = formataddr(("Reporte Disco", EMAIL_SENDER))
    msg['To'] = EMAIL_DESTINATARIO
    msg['Subject'] = "📊 Reporte de Espacio en Disco"
    msg.set_content("Adjunto encontrarás el reporte de espacio en disco generado automáticamente.")

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
        self.title("\ud83d\udcca Reporte de Espacio en Disco (Windows)")
        self.geometry("1000x600")
        self.configure(bg="#e6f0ff")
        self.protocol("WM_DELETE_WINDOW", self.on_cerrar)

        self.servidores = cargar_servidores()
        self.datos = []

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

        frame_tabla = tk.Frame(self)
        frame_tabla.pack(padx=10, pady=10, fill="both", expand=True)

        columns = ("Servidor", "Host", "Disco", "Nombre Volumen", "Total (GB)", "Usado (GB)", "Libre (GB)", "Uso (%)")
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style_name = "Monit.Treeview"
        style.configure(style_name, rowheight=24)
        style.configure(f"{style_name}.Heading", background="#D9D9D9", foreground="#000000", relief="flat", font=("Segoe UI", 10, "bold"), anchor="center")
        style.configure("Treeview.Heading", background="#D9D9D9", foreground="#000000", relief="flat", font=("Segoe UI", 10, "bold"), anchor="center")
        style.map(f"{style_name}.Heading",
                  background=[("pressed", "#D9D9D9"), ("active", "#E5E5E5"), ("!active", "#D9D9D9")],
                  foreground=[("pressed", "#000000"), ("active", "#000000"), ("!disabled", "#000000")])
        style.map("Treeview.Heading",
                  background=[("pressed", "#D9D9D9"), ("active", "#E5E5E5"), ("!active", "#D9D9D9")],
                  foreground=[("pressed", "#000000"), ("active", "#000000"), ("!disabled", "#000000")])
        style.layout("Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.text", {"sticky": "we"})
                ]})
            ]})
        ])
        self.tabla = ttk.Treeview(frame_tabla, columns=columns, show="headings", style=style_name)
        self.tabla.tag_configure("row", background="#FFFFFF", foreground="#000000")

        for col in columns:
            self.tabla.heading(col, text=str(col), anchor="center")
            self.tabla.column(col, width=110, anchor="center")

        vsb = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=vsb.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.progress = ttk.Progressbar(self, orient="horizontal", mode="indeterminate")
        self.progress.pack(fill="x", padx=20, pady=(0,10))

        btn_frame = tk.Frame(self, bg="#e6f0ff")
        btn_frame.pack(pady=5)

        ttk.Button(btn_frame, text="🔍 Consultar Servidores", command=self.thread_consultar).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="📂 Generar Excel", command=self.generar_excel_gui).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="📧 Generar y Enviar Email", command=self.generar_y_enviar).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="🧹 Limpiar Tabla", command=self.limpiar_tabla).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="📝 Mostrar como texto", command=self.mostrar_texto_formateado).pack(side="left", padx=10)

        self.actualizar_lista_servidores()

    def agregar_servidor(self):
        servidor = {
            "nombre": self.entry_nombre.get(),
            "host": self.entry_host.get(),
            "usuario": self.entry_usuario.get(),
            "clave": self.entry_clave.get()
        }
        if not all(servidor.values()):
            messagebox.showwarning("Campos vacíos", "Completa todos los campos")
            return
        self.servidores.append(servidor)
        guardar_servidores(self.servidores)
        self.actualizar_lista_servidores()
        self.entry_nombre.delete(0, tk.END)
        self.entry_host.delete(0, tk.END)
        self.entry_usuario.delete(0, tk.END)
        self.entry_clave.delete(0, tk.END)

    def eliminar_servidor(self):
        seleccion = self.lista_servidores.curselection()
        if not seleccion:
            return
        index = seleccion[0]
        del self.servidores[index]
        guardar_servidores(self.servidores)
        self.actualizar_lista_servidores()

    def actualizar_lista_servidores(self):
        self.lista_servidores.delete(0, tk.END)
        for srv in self.servidores:
            self.lista_servidores.insert(tk.END, f"{srv['nombre']} - {srv['host']}")

    def thread_consultar(self):
        self.progress.start()
        threading.Thread(target=self.consultar_servidores).start()

    def consultar_servidores(self):
        self.tabla.delete(*self.tabla.get_children())
        self.datos = []
        for srv in self.servidores:
            resultados = obtener_disco_windows(srv)
            self.datos.extend(resultados)
            for d in resultados:
                self.tabla.insert("", tk.END, values=(
                    srv["nombre"], srv["host"], d["Disco"], d["Nombre Volumen"], d["Total (GB)"], d["Usado (GB)"], d["Libre (GB)"], d["Uso (%)"]
                ), tags=("row",))
        self.progress.stop()

    def generar_excel_gui(self):
        if not self.datos:
            messagebox.showwarning("Sin datos", "Primero consulta los servidores.")
            return
        archivo = "reporte_espacio_disco.xlsx"
        generar_excel(self.datos, archivo)
        messagebox.showinfo("Éxito", f"Reporte generado: {archivo}")

    def generar_y_enviar(self):
        if not self.datos:
            messagebox.showwarning("Sin datos", "Primero consulta los servidores.")
            return
        archivo = "reporte_espacio_disco.xlsx"
        generar_excel(self.datos, archivo)
        try:
            enviar_email_con_adjunto(archivo)
            messagebox.showinfo("Éxito", "Correo enviado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo enviar el correo: {e}")

    def limpiar_tabla(self):
        self.tabla.delete(*self.tabla.get_children())
        self.datos = []

    def mostrar_texto_formateado(self):
        if not self.datos:
            messagebox.showwarning("Sin datos", "Primero consulta los servidores.")
            return

        texto = ""
        servidores = {}
        fecha_formateada = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

        for fila in self.datos:
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

        ventana_texto = tk.Toplevel(self)
        ventana_texto.title("Reporte en texto formateado")
        text_widget = tk.Text(ventana_texto, width=100, height=35)
        text_widget.pack(padx=10, pady=10)
        text_widget.insert("1.0", texto)
        text_widget.config(state="disabled")

    def on_cerrar(self):
        self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()
