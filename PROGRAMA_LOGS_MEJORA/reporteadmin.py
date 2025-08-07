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

def generar_excel(data, path):
    df = pd.DataFrame(data)
    df.drop(columns=["Resumen"], inplace=True)
    df.to_excel(path, index=False)

def generar_txt(resumenes, path):
    with open(path, "w", encoding="utf-8") as f:
        for linea in resumenes:
            f.write(linea + "\n")

def enviar_email(resumenes, excel_path):
    msg = EmailMessage()
    msg['From'] = formataddr(("Reporte Memoria", EMAIL_SENDER))
    msg['To'] = EMAIL_DESTINATARIO
    msg['Subject'] = "\U0001F4C8 Reporte de Uso de Memoria RAM"

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

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("\U0001F4C8 Uso de Memoria RAM")
        self.geometry("750x600")
        self.configure(bg="#f5faff")
        self.resumenes = []
        self.servidores = cargar_servidores()
        self.datos = []

        self.crear_widgets()

    def crear_widgets(self):
        titulo = tk.Label(self, text="Reporte de Memoria RAM", font=("Segoe UI", 16, "bold"), bg="#f5faff")
        titulo.pack(pady=(15, 5))

        frame = tk.Frame(self, bg="#ffffff", bd=1, relief="solid")
        frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.texto = tk.Text(frame, height=20, font=("Consolas", 11), bd=0, wrap="none")
        self.texto.pack(fill="both", expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(frame, command=self.texto.yview)
        scrollbar.pack(side="right", fill="y")
        self.texto.config(yscrollcommand=scrollbar.set)

        btns = tk.Frame(self, bg="#f5faff")
        btns.pack(pady=10)

        estilo = ttk.Style()
        estilo.configure("TButton", font=("Segoe UI", 10))

        ttk.Button(btns, text="\U0001F50D Consultar", command=self.consultar).pack(side="left", padx=10)
        ttk.Button(btns, text="\U0001F4BE Guardar Excel", command=self.guardar_excel).pack(side="left", padx=10)
        ttk.Button(btns, text="\U0001F4E7 Enviar Email", command=self.enviar).pack(side="left", padx=10)

    def consultar(self):
        self.datos = []
        self.resumenes = []
        self.texto.delete("1.0", tk.END)

        for srv in self.servidores:
            resultado = obtener_memoria(srv)
            self.datos.append(resultado)
            self.resumenes.append(resultado["Resumen"])
            self.texto.insert(tk.END, resultado["Resumen"] + "\n\n")

    def guardar_excel(self):
        if not self.datos:
            messagebox.showwarning("Sin datos", "Realiza una consulta primero.")
            return
        nombre = f"reporte_memoria_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        generar_excel(self.datos, nombre)
        messagebox.showinfo("Excel guardado", f"Archivo: {nombre}")

    def enviar(self):
        if not self.datos:
            messagebox.showwarning("Sin datos", "Realiza una consulta primero.")
            return
        nombre_excel = f"reporte_memoria_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        generar_excel(self.datos, nombre_excel)
        enviar_email(self.resumenes, nombre_excel)
        messagebox.showinfo("Correo enviado", f"Enviado a {EMAIL_DESTINATARIO}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
