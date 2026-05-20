import pandas as pd
import requests
import io
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# ================= CONFIGURACION =================
EXCEL_URL = "https://valserindustriales-my.sharepoint.com/:x:/p/sst/IQBrBvFaNIKdS4PM-S-5DxgMAS6dHMmzLp9DMHRM2fTcwJ0?e=9XHSzo&download=1"

CORREOS_DESTINO = [

    "tecnicodeservicios@valserindustriales.com"
]

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")

# ================= DESCARGAR EXCEL =================
def descargar_excel():
    r = requests.get(EXCEL_URL, allow_redirects=True)
    if r.status_code != 200:
        raise Exception(f"No fue posible descargar el Excel. Codigo HTTP: {r.status_code}")

    content_type = r.headers.get("Content-Type", "").lower()
    if "html" in content_type:
        raise Exception("OneDrive no entregó archivo Excel directo. Revisar permisos públicos del enlace.")

    return r.content

# ================= NORMALIZAR COLUMNAS =================
def normalizar_columnas(cols):
    nuevas = []
    for c in cols:
        c = str(c).strip().replace("\n", " ").replace("  ", " ")
        nuevas.append(c)
    return nuevas

def hacer_columnas_unicas(cols):
    contador = {}
    nuevas = []

    for c in cols:
        if c in contador:
            contador[c] += 1
            nuevas.append(f"{c}_{contador[c]}")
        else:
            contador[c] = 0
            nuevas.append(c)

    return nuevas

# ================= LEER MATRIZ =================
def cargar_matriz(bytes_excel):
    df = pd.read_excel(io.BytesIO(bytes_excel), sheet_name="Matriz de EMO", header=None)

    encabezados = df.iloc[3].fillna("").astype(str).tolist()
    encabezados = hacer_columnas_unicas(encabezados)

    data = df.iloc[4:].copy()
    data.columns = encabezados
    data = data.dropna(how="all")
    data = data.reset_index(drop=True)

    print("Columnas detectadas:", data.columns.tolist())
    return data

# ================= PREPARAR DATOS =================
def preparar_datos(df):
    # solo filas con documento
    df = df[df.iloc[:, 0].notna()]
    df = df[df.iloc[:, 0].astype(str).str.strip() != ""]

    hoy = pd.Timestamp.now().normalize()

        # ===== EXAMENES OCUPACIONALES
        df["FECHA PROXIMO EXAMEN"] = pd.to_datetime(
            df["FECHA EXAMEN PERIODICO"],
            errors="coerce",
            dayfirst=True
        )
        
        df["DIAS_EXAMEN"] = (
            df["FECHA PROXIMO EXAMEN"] - hoy
        ).dt.days
        
        # ===== CERTIFICADO ALTURAS
        df["FECHA ACTUALIZACION ALTURAS"] = pd.to_datetime(
            df["FECHA PARA ACTUALIZACION"],
            errors="coerce",
            dayfirst=True
        )

        df["DIAS_ALTURAS"] = (
            df["FECHA ACTUALIZACION ALTURAS"] - hoy
        ).dt.days
        
        # limpiar formato visual
        df["DIAS_EXAMEN"] = df["DIAS_EXAMEN"].astype("Int64")
        df["DIAS_ALTURAS"] = df["DIAS_ALTURAS"].astype("Int64")
            
            df["DIAS_EXAMEN"] = df["DIAS_EXAMEN"].astype("Int64")
            df["DIAS_ALTURAS"] = df["DIAS_ALTURAS"].astype("Int64")

    print("Total empleados válidos:", len(df))
    print("Fechas examen válidas:", df["FECHA PROXIMO EXAMEN"].notna().sum())
    print("Fechas alturas válidas:", df["FECHA ACTUALIZACION ALTURAS"].notna().sum())

    return df

# ================= CLASIFICAR ALERTAS GENERICO =================
def clasificar_alertas(df, columna_dias):
    preventivo = df[(df[columna_dias] >= 16) & (df[columna_dias] <= 45)]
    prioritario = df[(df[columna_dias] >= 0) & (df[columna_dias] <= 15)]
    critico = df[(df[columna_dias] < 0)]
    return preventivo, prioritario, critico

# ================= TABLA HTML GENERICA =================
def tabla_html(df, tipo="examen"):
    if df.empty:
        return "<p>No aplica.</p>"

    if tipo == "examen":
        mostrar = df[[
            "DOCUMENTO",
            "NOMBRES Y APELLIDOS",
            "CARGO",
            "FECHA PROXIMO EXAMEN",
            "DIAS_EXAMEN"
        ]].copy()

        mostrar["FECHA PROXIMO EXAMEN"] = mostrar["FECHA PROXIMO EXAMEN"].dt.strftime("%Y-%m-%d")
        mostrar.rename(columns={"DIAS_EXAMEN": "DIAS_RESTANTES"}, inplace=True)

    else:
        mostrar = df[[
            "DOCUMENTO",
            "NOMBRES Y APELLIDOS",
            "CARGO",
            "FECHA ACTUALIZACION ALTURAS",
            "DIAS_ALTURAS"
        ]].copy()

        mostrar["FECHA ACTUALIZACION ALTURAS"] = mostrar["FECHA ACTUALIZACION ALTURAS"].dt.strftime("%Y-%m-%d")
        mostrar.rename(columns={"DIAS_ALTURAS": "DIAS_RESTANTES"}, inplace=True)

    return mostrar.to_html(index=False, border=1)

# ================= ENVIAR CORREO =================
def enviar_correo(pe, pre, ce, pa, pra, ca):
    total = len(pe)+len(pre)+len(ce)+len(pa)+len(pra)+len(ca)

    if total == 0:
        print("Sin novedades, no se envia correo.")
        return

    html = f"""
    <html>
    <body style='font-family:Arial,sans-serif;'>

        <div style='background:#0b3d91;color:white;padding:15px;border-radius:8px;'>
            <h2>🔔 SISTEMA AUTOMÁTICO DE ALERTAS SST</h2>
            <h3>VALVULAS Y SERVICIOS INDUSTRIALES S.A.S</h3>
            <p>Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
        </div>

        <br>

        <h2 style='color:black;'>🩺 RESUMEN GENERAL EXÁMENES OCUPACIONALES</h2>
        <table style='width:100%;text-align:center;border-collapse:collapse;'>
            <tr>
                <td style='background:#d4edda;padding:12px;border:1px solid #ccc;'><b>🟢 Preventivos</b><br>{len(pe)}</td>
                <td style='background:#fff3cd;padding:12px;border:1px solid #ccc;'><b>🟡 Prioritarios</b><br>{len(pre)}</td>
                <td style='background:#f8d7da;padding:12px;border:1px solid #ccc;'><b>🔴 Críticos/Vencidos</b><br>{len(ce)}</td>
            </tr>
        </table>

        <br>

        <h2 style='color:black;'>⛑️ RESUMEN GENERAL CERTIFICADO DE ALTURAS</h2>
        <table style='width:100%;text-align:center;border-collapse:collapse;'>
            <tr>
                <td style='background:#d4edda;padding:12px;border:1px solid #ccc;'><b>🟢 Preventivos</b><br>{len(pa)}</td>
                <td style='background:#fff3cd;padding:12px;border:1px solid #ccc;'><b>🟡 Prioritarios</b><br>{len(pra)}</td>
                <td style='background:#f8d7da;padding:12px;border:1px solid #ccc;'><b>🔴 Críticos/Vencidos</b><br>{len(ca)}</td>
            </tr>
        </table>

        <br><hr>

        <h2 style='color:black;'>🩺 TABLAS DETALLADAS EXÁMENES OCUPACIONALES</h2>

        <h3 style='color:green;'>🟢 PERSONAL EN ALERTA PREVENTIVA EXÁMENES</h3>
        {tabla_html(pe, "examen")}

        <h3 style='color:#b8860b;'>🟡 PERSONAL EN ALERTA PRIORITARIA EXÁMENES</h3>
        {tabla_html(pre, "examen")}

        <h3 style='color:red;'>🔴 PERSONAL CRÍTICO / EXÁMENES VENCIDOS</h3>
        {tabla_html(ce, "examen")}

        <br><hr>

        <h2 style='color:black;'>⛑️ TABLAS DETALLADAS CERTIFICADO DE ALTURAS</h2>

        <h3 style='color:green;'>🟢 PERSONAL EN ALERTA PREVENTIVA ALTURAS</h3>
        {tabla_html(pa, "alturas")}

        <h3 style='color:#b8860b;'>🟡 PERSONAL EN ALERTA PRIORITARIA ALTURAS</h3>
        {tabla_html(pra, "alturas")}

        <h3 style='color:red;'>🔴 PERSONAL CRÍTICO / CERTIFICADO VENCIDO ALTURAS</h3>
        {tabla_html(ca, "alturas")}

        <br><hr>
        <p style='font-size:12px;color:gray;'>Este correo fue generado automáticamente por el Sistema de Vigilancia Documental SST - VALSER.</p>

    </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "🔔 Alerta SST - Exámenes Ocupacionales y Certificados de Alturas"
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(CORREOS_DESTINO)
    msg.attach(MIMEText(html, "html"))

    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(SMTP_USER, SMTP_PASS)
    server.sendmail(SMTP_USER, CORREOS_DESTINO, msg.as_string())
    server.quit()

    print("Correo enviado correctamente")

# ================= MAIN =================
def main():
    try:
        bytes_excel = descargar_excel()
        df = cargar_matriz(bytes_excel)
        df = preparar_datos(df)

        pe, pre, ce = clasificar_alertas(df, "DIAS_EXAMEN")
        pa, pra, ca = clasificar_alertas(df, "DIAS_ALTURAS")

        enviar_correo(pe, pre, ce, pa, pra, ca)

    except Exception as e:
        print("ERROR GENERAL:", str(e))
        raise

if __name__ == "__main__":
    main()
