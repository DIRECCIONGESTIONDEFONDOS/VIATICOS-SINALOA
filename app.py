#!/usr/bin/env python3
import json, os, io, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse
from openpyxl import load_workbook

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
PTPL_SOL   = os.path.join(BASE_DIR, 'plantillas', 'Solicitud_plantilla.xlsx')
PTPL_REP   = os.path.join(BASE_DIR, 'plantillas', 'Reposicion_plantilla.xlsx')

GMAIL_USER = os.environ.get('GMAIL_USER', '')   # tu@gmail.com
GMAIL_PASS = os.environ.get('GMAIL_PASS', '')   # contraseña de aplicación

# ── NÚMERO A LETRAS ────────────────────────────────────────────────────────────
def num_letras(n):
    u = ['','UN','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE','DIEZ',
         'ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS','DIECISIETE','DIECIOCHO','DIECINUEVE']
    d = ['','','VEINTE','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA']
    c = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
         'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS']
    if n == 0: return 'CERO'
    r = ''
    if n >= 1000:
        m = int(n // 1000); r += ('MIL ' if m == 1 else num_letras(m) + ' MIL '); n = int(n % 1000)
    if n >= 100:
        r += ('CIEN ' if n == 100 else c[int(n // 100)] + ' '); n = int(n % 100)
    if n >= 20: r += d[int(n // 10)] + (' Y ' + u[int(n % 10)] if n % 10 > 0 else '') + ' '
    elif n > 0: r += u[int(n)] + ' '
    return r.strip()

def monto_letras(m):
    e = int(m); cts = round((m - e) * 100)
    return num_letras(e) + (f' CON {cts:02d}/100' if cts > 0 else '') + ' M.N.'

# ── RFC EN CELDAS ──────────────────────────────────────────────────────────────
def rfc_sol(rfc):
    ch = list(rfc.upper().replace(' ', ''))
    m = {}
    for i, col in enumerate(['C','D','E','F','G','H','I','J','K']):
        m[col + '21'] = ch[i] if i < len(ch) else ''
    m['P21'] = ch[9]  if len(ch) > 9  else ''
    m['Q21'] = ''.join(ch[10:]) if len(ch) > 10 else 'NOMBRE'
    return m

def rfc_rep(rfc):
    ch = list(rfc.upper().replace(' ', ''))
    m = {}
    for i, col in enumerate(['B','C','D','E','F','G','H','I','J']):
        m[col + '21'] = ch[i] if i < len(ch) else ''
    m['O21'] = ch[9]  if len(ch) > 9  else ''
    m['P21'] = ''.join(ch[10:]) if len(ch) > 10 else 'NOMBRE'
    return m

# ── GENERAR EXCEL ──────────────────────────────────────────────────────────────
def generar_solicitud(d):
    wb  = load_workbook(PTPL_SOL)
    ws  = wb['MAZATLAN']
    total  = float(d.get('total', 0))
    nombre = d.get('nombre', '').upper()
    ws['B11'] = d.get('titulo', 'SOLICITUD DE OFICIO DE COMISIÓN')
    ws['Q12'] = d.get('folio', '')
    ws['Q14'] = d.get('fecha_exp', '')
    ws['E18'] = d.get('programa', '')
    for cel, val in rfc_sol(d.get('rfc', '')).items(): ws[cel] = val
    ws['R21'] = nombre
    ws['C22'] = d.get('cargo', '').upper()
    ws['B25'] = d.get('destino', '').upper()
    ws['R25'] = d.get('fecha_ini', '')
    ws['B27'] = d.get('motivo', '').upper()
    ws['R27'] = d.get('fecha_fin', '')
    letras = monto_letras(total)
    ws['B29'] = (f"RECIBÍ DE LA SECRETARÍA DE ECONOMÍA LA CANTIDAD DE ${total:,.2f} "
                 f"SON: ({letras}) POR CONCEPTO DE  VIATICOS  PARA CUMPLIR LAS COMISIONES ARRIBA INDICADAS.")
    ws['R29'] = d.get('alimentos')   or ''
    ws['R30'] = d.get('gasolina')    or ''
    ws['R31'] = d.get('peaje')       or ''
    ws['R32'] = d.get('hospedaje')   or ''
    ws['B33'] = f"__________________________________________________                    {nombre}"
    ws['B41'] = d.get('auth1_nombre', '').upper()
    ws['P41'] = d.get('auth2_nombre', '').upper()
    ws['B42'] = d.get('auth1_cargo',  '').upper()
    ws['P42'] = d.get('auth2_cargo',  '').upper()
    for sn in list(wb.sheetnames):
        if sn != 'MAZATLAN': del wb[sn]
    ws.title = 'Solicitud'
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

def generar_reposicion(d):
    wb  = load_workbook(PTPL_REP)
    ws  = wb['reposición']
    total  = float(d.get('total', 0))
    nombre = d.get('nombre', '').upper()
    ws['P12'] = d.get('folio', '')
    ws['P14'] = d.get('fecha_exp', '')
    ws['D18'] = d.get('programa', '')
    for cel, val in rfc_rep(d.get('rfc', '')).items(): ws[cel] = val
    ws['Q21'] = nombre
    ws['B22'] = d.get('cargo', '').upper()
    ws['A25'] = d.get('destino', '').upper()
    ws['Q25'] = d.get('fecha_ini', '')
    ws['A27'] = d.get('motivo', '').upper()
    ws['Q27'] = d.get('fecha_fin', '')
    letras = monto_letras(total)
    ws['A29'] = (f"RECIBÍ DE LA SECRETARIA DE DESARROLLO ECONOMICO LA CANTIDAD DE ${total:,.2f} "
                 f"(SON: {letras}), POR CONCEPTO DE  VIÁTICOS,  PARA CUMPLIR LAS COMISIONES ARRIBA INDICADAS.")
    ws['Q29'] = d.get('alimentos')   or ''
    ws['Q30'] = d.get('hospedaje')   or ''
    ws['Q31'] = d.get('combustible') or ''
    ws['Q32'] = d.get('peaje')       or ''
    ws['Q33'] = d.get('transporte')  or ''
    ws['A33'] = nombre
    ws['A43'] = d.get('auth1_nombre', '').upper()
    ws['O43'] = d.get('auth2_nombre', '').upper()
    ws['A44'] = d.get('auth1_cargo',  '').upper()
    ws['O44'] = d.get('auth2_cargo',  '').upper()
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# ── CUERPO DEL CORREO ──────────────────────────────────────────────────────────
def html_correo(d, total, tipo_doc):
    fmt = lambda n: f"${float(n):,.2f}" if n else "—"
    es_rep = 'REPOSICIÓN' in tipo_doc.upper()
    conceptos_html = ""
    pares = [
        ("Alimentos",   d.get('alimentos')),
        ("Gasolina/Combustible", d.get('gasolina') or d.get('combustible')),
        ("Peaje",       d.get('peaje')),
        ("Hospedaje",   d.get('hospedaje')),
    ]
    if es_rep:
        pares.append(("Transporte", d.get('transporte')))
    for lbl, val in pares:
        if val:
            conceptos_html += f"""
            <tr>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;color:#555;">{lbl}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:600;">{fmt(val)}</td>
            </tr>"""

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f6fa;font-family:'DM Sans',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fa;padding:32px 0;">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08);">

  <!-- HEADER -->
  <tr><td style="background:#0f2543;padding:28px 36px;">
    <div style="font-size:11px;letter-spacing:.15em;color:#c8a84b;text-transform:uppercase;margin-bottom:6px;">Secretaría de Economía · Sinaloa</div>
    <div style="font-size:22px;font-weight:700;color:#c8a84b;">{tipo_doc}</div>
    <div style="font-size:13px;color:rgba(255,255,255,.5);margin-top:4px;">Folio: {d.get('folio','—')}</div>
  </td></tr>

  <!-- BODY -->
  <tr><td style="padding:32px 36px;">
    <p style="color:#555;font-size:14px;margin:0 0 24px;">Estimado(a) <strong>{d.get('nombre','').title()}</strong>,</p>
    <p style="color:#555;font-size:14px;margin:0 0 24px;">Se adjunta al presente correo tu solicitud de viáticos con el siguiente detalle:</p>

    <!-- DATOS COMISIÓN -->
    <div style="background:#f0f5ff;border-radius:8px;padding:16px 20px;margin-bottom:20px;">
      <div style="font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#c8a84b;margin-bottom:12px;">Datos de la comisión</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;color:#333;">
        <tr><td style="padding:4px 0;color:#888;width:140px;">Destino</td><td style="font-weight:600;">{d.get('destino','—').title()}</td></tr>
        <tr><td style="padding:4px 0;color:#888;">Fecha inicial</td><td style="font-weight:600;">{d.get('fecha_ini','—')}</td></tr>
        <tr><td style="padding:4px 0;color:#888;">Fecha final</td><td style="font-weight:600;">{d.get('fecha_fin','—')}</td></tr>
        <tr><td style="padding:4px 0;color:#888;">Programa</td><td style="font-weight:600;">{d.get('programa','—')}</td></tr>
        <tr><td style="padding:4px 0;color:#888;">Motivo</td><td style="font-weight:600;">{d.get('motivo','—').capitalize()}</td></tr>
      </table>
    </div>

    <!-- DESGLOSE -->
    <div style="font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#c8a84b;margin-bottom:10px;">Desglose de viáticos</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="font-size:14px;border:1px solid #eee;border-radius:8px;overflow:hidden;">
      {conceptos_html}
      <tr style="background:#0f2543;">
        <td style="padding:12px;color:#c8a84b;font-weight:700;font-size:13px;letter-spacing:.05em;text-transform:uppercase;">TOTAL</td>
        <td style="padding:12px;color:#c8a84b;font-weight:700;font-size:18px;text-align:right;">${total:,.2f}</td>
      </tr>
    </table>

    <p style="color:#aaa;font-size:12px;margin:24px 0 0;line-height:1.6;">
      Este correo es generado automáticamente por el Sistema de Viáticos de la Secretaría de Economía del Estado de Sinaloa.<br>
      El documento oficial se encuentra adjunto en formato .xlsx.
    </p>
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="background:#f9f9f9;padding:16px 36px;border-top:1px solid #eee;">
    <div style="font-size:11px;color:#aaa;text-align:center;">Secretaría de Economía · Gobierno del Estado de Sinaloa · 2025</div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""

# ── ENVIAR CORREO ──────────────────────────────────────────────────────────────
def enviar_correo(destinatario, asunto, html, xlsx_bytes, nombre_archivo):
    msg = MIMEMultipart('alternative')
    msg['From']    = f"Viáticos SE Sinaloa <{GMAIL_USER}>"
    msg['To']      = destinatario
    msg['Subject'] = asunto

    msg.attach(MIMEText(html, 'html', 'utf-8'))

    adj = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    adj.set_payload(xlsx_bytes)
    encoders.encode_base64(adj)
    adj.add_header('Content-Disposition', 'attachment', filename=nombre_archivo)
    msg.attach(adj)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.sendmail(GMAIL_USER, destinatario, msg.as_string())

# ── SERVIDOR HTTP ──────────────────────────────────────────────────────────────
HTML_PATH = os.path.join(BASE_DIR, 'index.html')

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args): pass

    def cors(self):
        self.send_header('Access-Control-Allow-Origin',  '*')
        self.send_header('Access-Control-Allow-Methods', 'POST,GET,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_OPTIONS(self):
        self.send_response(200); self.cors(); self.end_headers()

    def do_GET(self):
        if urlparse(self.path).path in ('/', '/index.html'):
            with open(HTML_PATH, 'rb') as f: content = f.read()
            self.send_response(200)
            self.send_header('Content-Type',   'text/html; charset=utf-8')
            self.send_header('Content-Length', len(content))
            self.cors(); self.end_headers(); self.wfile.write(content)
        else:
            self.send_response(404); self.end_headers()

    def do_POST(self):
        path   = urlparse(self.path).path
        length = int(self.headers.get('Content-Length', 0))
        d      = json.loads(self.rfile.read(length))

        # Determinar tipo
        es_rep = path == '/generar/reposicion'
        if path not in ('/generar/solicitud', '/generar/reposicion'):
            self.send_response(404); self.end_headers(); return

        # Generar Excel
        xlsx  = generar_reposicion(d) if es_rep else generar_solicitud(d)
        tipo  = d.get('titulo', 'SOLICITUD DE OFICIO DE COMISIÓN')
        total = float(d.get('total', 0))
        nom   = d.get('nombre', '').replace(' ', '_')[:20]
        dest_str = d.get('destino', '').replace(' ', '_')[:12]
        prefix = 'Reposicion' if es_rep else 'Solicitud'
        filename = f"{prefix}_{nom}_{dest_str}.xlsx"

        # Enviar correo si hay email
        email_dest = d.get('email_destinatario', '').strip()
        email_ok   = False
        email_err  = ''
        if email_dest and GMAIL_USER and GMAIL_PASS:
            try:
                asunto = f"Viáticos — {tipo} | {d.get('destino','').title()} | {d.get('fecha_ini','')}"
                html   = html_correo(d, total, tipo)
                enviar_correo(email_dest, asunto, html, xlsx, filename)
                email_ok = True
            except Exception as e:
                email_err = str(e)

        # Responder con el Excel + estado del correo
        resp = {
            'filename': filename,
            'email_ok':  email_ok,
            'email_err': email_err,
            'xlsx_b64':  __import__('base64').b64encode(xlsx).decode()
        }
        body = json.dumps(resp).encode()
        self.send_response(200)
        self.send_header('Content-Type',   'application/json')
        self.send_header('Content-Length', len(body))
        self.cors(); self.end_headers(); self.wfile.write(body)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    print(f"Servidor iniciado en puerto {port}")
    HTTPServer(('0.0.0.0', port), Handler).serve_forever()
