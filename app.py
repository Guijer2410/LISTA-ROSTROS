import base64, calendar, datetime as dt, hashlib, html, io, os, secrets, sqlite3, unicodedata, urllib.request
from zoneinfo import ZoneInfo
import altair as alt
import cv2, numpy as np, pandas as pd, streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageFilter

APP = "ESAVE Checklist"
FONDO = next((p for p in ("fondo.jpg", "fondo.jpeg", "fondo.png", "assets/fondo.jpg") if os.path.exists(p)), None)
FONDO_MOVIL = next((p for p in ("fondo_movil.jpg", "assets/fondo_movil.jpg") if os.path.exists(p)), None)
ESCUDO = next((p for p in ("escudo.png", "escudo.jpg", "escudo.jpeg", "assets/escudo.png") if os.path.exists(p)), None)
try:
    _icono = Image.open(ESCUDO) if ESCUDO else "✈️"
except Exception:
    _icono = "✈️"
st.set_page_config(page_title=APP, page_icon=_icono)
DB = "asistencia.db"
MAX_FALTAS = 3
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
URLS = {
    "yunet.onnx": "https://github.com/opencv/opencv_zoo/raw/main/models/face_detection_yunet/face_detection_yunet_2023mar.onnx",
    "sface.onnx": "https://github.com/opencv/opencv_zoo/raw/main/models/face_recognition_sface/face_recognition_sface_2021dec.onnx",
}

# ---------------- Datos del plan de estudios (Ingeniería Aeronáutica) ----------------
PENSUM = {
    1: [("Pre-Cálculo", 3), ("Programación I", 3), ("Física General", 3), ("Química", 3), ("Ética Profesional y Medio Ambiente", 2),
        ("Introducción a la Ing. Aeronáutica", 2), ("Cátedra Institucional y Constitución Política", 2)],
    2: [("Cálculo Diferencial", 3), ("Álgebra Lineal", 3), ("Programación II", 3), ("Física Mecánica", 3),
        ("Dibujo Técnico e Introducción al CAD", 3), ("Derecho Aéreo y Regulaciones Aeronáuticas", 3)],
    3: [("Cálculo Integral", 3), ("Estática y Dinámica", 3), ("Probabilidad y Estadística", 3), ("Física Calor y Ondas", 3),
        ("Electricidad y Magnetismo", 3), ("Fundamentos del Mantenimiento Aeronáutico", 3)],
    4: [("Cálculo Multivariado", 3), ("Mecánica de Fluidos", 2), ("Electrónica", 3), ("Combustibles y Lubricantes", 2),
        ("Gerencia Aeronáutica", 3), ("Confiabilidad", 3), ("Inglés Técnico I", 2)],
    5: [("Matemáticas Especiales", 3), ("Ciencia de Materiales", 2), ("Métodos Numéricos", 3), ("Termodinámica", 3),
        ("Aerodinámica", 3), ("Gestión del Mantenimiento", 2), ("Inglés Técnico II", 2)],
    6: [("Introducción a los Mecanismos", 2), ("Resistencia de Materiales", 2), ("Aviónica", 3), ("Dinámica de Vuelo", 2),
        ("Sistemas de las Aeronaves", 2), ("Adm. Recursos de Mantenimiento I", 3), ("Investigación de Accidentes de Aviación", 2),
        ("Metodología de la Investigación", 2)],
    7: [("Estructuras de Aeronaves", 3), ("Procesos de Manufactura", 2), ("Motores de Aviación", 4), ("Electiva Profesional I", 3),
        ("Inspección y Supervisión del Mantenimiento Aeronáutico", 2), ("Aeronaves no Tripuladas", 2), ("Proyecto de Grado I", 2)],
    8: [("Hélices Rotores y Tren de Potencia", 2), ("Mantenimiento y Reparación de Motores de Aviación", 3),
        ("Adm. de Recursos de Mantenimiento II", 3), ("Control y Aseguramiento de la Calidad", 3),
        ("Aeródromos, Aeropuertos y Helipuertos", 2), ("Electiva Profesional II", 3), ("Proyecto de Grado II", 2)],
    9: [("Control Automático de Vuelo", 3), ("Seguridad Operacional", 3), ("Operación de Aeronaves y Documentación Técnica", 3),
        ("Certificación de Organizaciones Aeronáuticas", 4), ("Electiva Complementaria", 3)],
}
DOCENTES_PDF = [
    ("Hernán D. Torres Valencia", "htorres"), ("Didier Aldana Rodríguez", "daldana"), ("José M. Arguello Pinzón", "jarguello"),
    ("Diego L. Ávila Granados", "davila"), ("Arjuna Castrillón Camacho", "acastrillon"), ("Rafael R. Cubillos Ramírez", "rcubillos"),
    ("Juan Carlos Gamba Torres", "jgamba"), ("Paula A. Garcés Constaín", "pgarces"), ("John E. González Arias", "jgonzalez"),
    ("Daniel S. Gutiérrez Pacheco", "dgutierrez"), ("Juan C. Martínez Escobar", "jmartinez"), ("Pedro Fernando Melo Daza", "pmelo"),
    ("Fabio A. Merchán Rincón", "fmerchan"), ("Iván F. Rodríguez Barón", "irodriguez"), ("Oscar A. Rojas Martínez", "orojas"),
    ("David E. Sánchez Varela", "dsanchez"), ("Sebastián Valencia", "svalencia"), ("Mauricio Vinasco Téllez", "mvinasco"),
]


# ---------------- Base de datos ----------------
def db():
    con = sqlite3.connect(DB, timeout=30)
    con.row_factory = sqlite3.Row
    return con


def q(sql, args=()):
    con = db()
    try:
        filas = con.execute(sql, args).fetchall()
        con.commit()
        return filas
    finally:
        con.close()


def muchos(sql, lista):
    con = db()
    try:
        con.executemany(sql, lista)
        con.commit()
    finally:
        con.close()


def df(sql, args=()):
    return pd.DataFrame([dict(r) for r in q(sql, args)])


def marcas(lista):
    return ",".join("?" * len(lista))


def hashear(clave, sal):
    return hashlib.pbkdf2_hmac("sha256", clave.encode(), bytes.fromhex(sal), 100_000).hex()


def clave_aleatoria():
    return "".join(secrets.choice("abcdefghjkmnpqrstuvwxyz23456789") for _ in range(8))


def crear_usuario(usuario, nombre, rol, clave):
    sal = secrets.token_hex(8)
    q("INSERT INTO usuarios(usuario,nombre,rol,sal,clave) VALUES(?,?,?,?,?)",
      (usuario.strip().lower(), nombre.strip(), rol, sal, hashear(clave, sal)))


def cambiar_clave(uid, clave):
    sal = secrets.token_hex(8)
    q("UPDATE usuarios SET sal=?, clave=? WHERE id=?", (sal, hashear(clave, sal), int(uid)))


def verificar(usuario, clave):
    f = q("SELECT * FROM usuarios WHERE usuario=?", (usuario.strip().lower(),))
    if f and secrets.compare_digest(f[0]["clave"], hashear(clave, f[0]["sal"])):
        return dict(f[0])
    return None


DIAS_SESION = 7


def crear_sesion(uid):
    token = secrets.token_urlsafe(24)
    q("INSERT INTO sesiones(token,usuario_id,creado) VALUES(?,?,?)", (token, int(uid), ahora().isoformat()))
    return token


def usuario_de_sesion(token):
    if not token:
        return None
    limite = (ahora() - dt.timedelta(days=DIAS_SESION)).isoformat()
    f = q("SELECT u.* FROM sesiones s JOIN usuarios u ON u.id=s.usuario_id WHERE s.token=? AND s.creado>=?", (str(token), limite))
    return dict(f[0]) if f else None


def cerrar_sesion(token):
    if token:
        q("DELETE FROM sesiones WHERE token=?", (str(token),))


def avisar(msg, icono="✅"):
    """Deja un aviso que se muestra (como notificación) justo después de refrescar la pantalla."""
    st.session_state.toast = (msg, icono)


def mostrar_aviso():
    t = st.session_state.pop("toast", None)
    if t:
        st.toast(t[0], icon=t[1])


def clave_admin_inicial():
    """Clave inicial del admin: viene de los Secrets de Streamlit (ADMIN_PASSWORD) o, si no hay, de un valor por defecto."""
    try:
        return str(st.secrets["ADMIN_PASSWORD"])
    except Exception:
        return "admin123"


def init_db():
    con = db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS usuarios(id INTEGER PRIMARY KEY, usuario TEXT UNIQUE, nombre TEXT, rol TEXT, sal TEXT, clave TEXT);
    CREATE TABLE IF NOT EXISTS materias(id INTEGER PRIMARY KEY, semestre INTEGER, nombre TEXT, creditos INTEGER, UNIQUE(semestre,nombre));
    CREATE TABLE IF NOT EXISTS asignaciones(docente_id INTEGER, materia_id INTEGER, PRIMARY KEY(docente_id,materia_id));
    CREATE TABLE IF NOT EXISTS estudiantes(id INTEGER PRIMARY KEY, codigo TEXT UNIQUE, nombre TEXT, semestre INTEGER, rostros BLOB, foto BLOB);
    CREATE TABLE IF NOT EXISTS clases(id INTEGER PRIMARY KEY, materia_id INTEGER, docente_id INTEGER, fecha TEXT, hora TEXT, UNIQUE(materia_id,fecha,hora));
    CREATE TABLE IF NOT EXISTS asistencia(clase_id INTEGER, estudiante_id INTEGER, presente INTEGER, metodo TEXT, PRIMARY KEY(clase_id,estudiante_id));
    """)
    if con.execute("SELECT COUNT(*) FROM materias").fetchone()[0] == 0:
        for sem, lista in PENSUM.items():
            con.executemany("INSERT OR IGNORE INTO materias(semestre,nombre,creditos) VALUES(?,?,?)", [(sem, n, c) for n, c in lista])
    try:
        if "foto" not in [r[1] for r in con.execute("PRAGMA table_info(estudiantes)")]:
            con.execute("ALTER TABLE estudiantes ADD COLUMN foto BLOB")
    except sqlite3.OperationalError:
        pass
    con.execute("CREATE TABLE IF NOT EXISTS sesiones(token TEXT PRIMARY KEY, usuario_id INTEGER, creado TEXT)")
    con.commit()
    con.close()
    if not q("SELECT 1 FROM usuarios WHERE rol='admin'"):
        try:
            crear_usuario("admin", "Administrador", "admin", clave_admin_inicial())
        except sqlite3.IntegrityError:
            pass  # otra sesión ya lo creó


def ahora():
    return dt.datetime.now(ZoneInfo("America/Bogota"))


# ---------------- Reconocimiento facial (OpenCV: YuNet + SFace) ----------------
@st.cache_resource(show_spinner="Descargando modelos de rostros (solo la primera vez)...")
def modelo_reconocedor():
    os.makedirs("modelos", exist_ok=True)
    for nombre, url in URLS.items():
        ruta = f"modelos/{nombre}"
        if not os.path.exists(ruta) or os.path.getsize(ruta) < 100_000:
            urllib.request.urlretrieve(url, ruta)
    return cv2.FaceRecognizerSF.create("modelos/sface.onnx", "")


def leer_imagen(archivo):
    img = cv2.imdecode(np.frombuffer(archivo.getvalue(), np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        return None
    h, w = img.shape[:2]
    m = max(h, w)
    if m > 1800:
        img = cv2.resize(img, (int(w * 1800 / m), int(h * 1800 / m)))
    return img


def rostros(img):
    rec = modelo_reconocedor()
    h, w = img.shape[:2]
    det = cv2.FaceDetectorYN.create("modelos/yunet.onnx", "", (w, h), 0.7, 0.3, 5000)
    _, caras = det.detect(img)
    if caras is None or len(caras) == 0:
        return np.zeros((0, 15), np.float32), np.zeros((0, 128), np.float32)
    caras = caras[np.argsort(caras[:, 0])]
    vect = np.array([rec.feature(rec.alignCrop(img, c)).ravel() for c in caras], dtype=np.float32)
    return caras, vect


def normalizar(v):
    return v / (np.linalg.norm(v, axis=1, keepdims=True) + 1e-9)


def reconocer(vect, umbral, semestre):
    """Devuelve, por cada rostro, (id_estudiante, similitud) o None. Cada estudiante se asigna a un solo rostro."""
    resultado = [None] * len(vect)
    ids, embs = [], []
    for e in q("SELECT id, rostros FROM estudiantes WHERE rostros IS NOT NULL AND semestre=?", (int(semestre),)):
        for fila in np.frombuffer(e["rostros"], np.float32).reshape(-1, 128):
            ids.append(e["id"])
            embs.append(fila)
    if len(vect) == 0 or not embs:
        return resultado
    S = normalizar(vect) @ normalizar(np.array(embs)).T
    est_ids = sorted(set(ids))
    M = np.zeros((len(vect), len(est_ids)))
    for j, eid in enumerate(est_ids):
        M[:, j] = S[:, [k for k, x in enumerate(ids) if x == eid]].max(axis=1)
    usados_c, usados_e = set(), set()
    for flat in np.argsort(-M, axis=None):
        i, j = divmod(int(flat), M.shape[1])
        if M[i, j] < umbral:
            break
        if i in usados_c or j in usados_e:
            continue
        resultado[i] = (est_ids[j], float(M[i, j]))
        usados_c.add(i)
        usados_e.add(j)
    return resultado


def sin_tildes(s):
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode()


def anotar(img, caras, res):
    im = img.copy()
    nombres = {r["id"]: r["nombre"] for r in q("SELECT id, nombre FROM estudiantes")}
    esc = max(0.5, im.shape[1] / 1500)
    for c, r in zip(caras, res):
        x, y, w, h = map(int, c[:4])
        color = (0, 190, 0) if r else (0, 0, 230)
        texto = " ".join(sin_tildes(nombres[r[0]]).split()[:2]) if r else "?"
        cv2.rectangle(im, (x, y), (x + w, y + h), color, max(2, int(esc * 2)))
        cv2.putText(im, texto, (x, max(18, y - 6)), cv2.FONT_HERSHEY_SIMPLEX, esc * 0.7, color, max(2, int(esc * 2)))
    return cv2.cvtColor(im, cv2.COLOR_BGR2RGB)


def miniatura(img, cara):
    """Recorte cuadrado del rostro (96x96, JPEG) para mostrarlo en las listas."""
    x, y, w, h = [float(v) for v in cara[:4]]
    cx, cy, r = x + w / 2, y + h / 2, max(w, h) * 0.85
    H, W = img.shape[:2]
    rec = img[int(max(0, cy - r)):int(min(H, cy + r)), int(max(0, cx - r)):int(min(W, cx + r))]
    if rec.size == 0:
        return None
    rec = cv2.resize(rec, (96, 96), interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", rec, [cv2.IMWRITE_JPEG_QUALITY, 80])
    return buf.tobytes() if ok else None


def uri_foto(b):
    return "data:image/jpeg;base64," + base64.b64encode(bytes(b)).decode() if isinstance(b, (bytes, bytearray)) else None


def guardar_estudiante(codigo, nombre, sem, vecs, foto=None):
    nuevo = np.array(vecs, np.float32)
    ex = q("SELECT id, rostros FROM estudiantes WHERE codigo=?", (codigo,))
    if ex:
        previos = np.frombuffer(ex[0]["rostros"], np.float32).reshape(-1, 128) if ex[0]["rostros"] else np.zeros((0, 128), np.float32)
        todo = np.vstack([previos, nuevo])[-8:]
        q("UPDATE estudiantes SET nombre=?, semestre=?, rostros=?, foto=COALESCE(?, foto) WHERE id=?", (nombre, int(sem), todo.tobytes(), foto, ex[0]["id"]))
    else:
        q("INSERT INTO estudiantes(codigo,nombre,semestre,rostros,foto) VALUES(?,?,?,?,?)", (codigo, nombre, int(sem), nuevo.tobytes(), foto))


# ---------------- Lógica de asistencia ----------------
def mis_materias(usr):
    if usr["rol"] == "admin":
        return df("SELECT * FROM materias ORDER BY semestre, id")
    return df("SELECT m.* FROM materias m JOIN asignaciones a ON a.materia_id=m.id WHERE a.docente_id=? ORDER BY m.semestre, m.id", (usr["id"],))


def etiquetas(mm):
    return {int(r.id): f"S{r.semestre} · {r.nombre}" for r in mm.itertuples()}


def obtener_clase(mid, did, fecha, hora):
    q("INSERT OR IGNORE INTO clases(materia_id,docente_id,fecha,hora) VALUES(?,?,?,?)", (int(mid), int(did), fecha, hora))
    return q("SELECT id FROM clases WHERE materia_id=? AND fecha=? AND hora=?", (int(mid), fecha, hora))[0]["id"]


def faltas_df(mids, desde=None, hasta=None):
    if not mids:
        return pd.DataFrame()
    cond, args = "", list(mids)
    if desde:
        cond += " AND c.fecha>=?"
        args.append(str(desde))
    if hasta:
        cond += " AND c.fecha<=?"
        args.append(str(hasta))
    return df(f"""SELECT e.codigo, e.nombre, m.nombre AS materia, COUNT(*) AS clases,
        SUM(a.presente) AS asistencias, SUM(1 - a.presente) AS faltas
        FROM asistencia a JOIN estudiantes e ON e.id=a.estudiante_id JOIN clases c ON c.id=a.clase_id
        JOIN materias m ON m.id=c.materia_id WHERE m.id IN ({marcas(mids)}){cond}
        GROUP BY e.id, m.id ORDER BY faltas DESC, e.nombre""", tuple(args))


def estado_faltas(n):
    return "Al día" if n <= 1 else "En riesgo" if n <= MAX_FALTAS else "Alerta"


def alertas(mids):
    f = faltas_df(mids)
    if f.empty:
        return []
    g = f[f.faltas > MAX_FALTAS]
    return [f"🚨 {r.nombre} ({r.codigo}) tiene {int(r.faltas)} faltas en {r.materia}" for r in g.itertuples()]


def excel_materia(mid, cid=None):
    m = q("SELECT * FROM materias WHERE id=?", (int(mid),))[0]
    filtro, args = ("AND c.id=?", (int(mid), int(cid))) if cid else ("", (int(mid),))
    cl = df(f"SELECT c.id, c.fecha, c.hora FROM clases c WHERE c.materia_id=? {filtro} "
            "AND c.id IN (SELECT clase_id FROM asistencia) ORDER BY c.fecha, c.hora", args)
    if cl.empty:
        return None
    est = df("SELECT id, codigo, nombre FROM estudiantes WHERE semestre=? ORDER BY nombre", (m["semestre"],))
    asis = {(int(r.clase_id), int(r.estudiante_id)): int(r.presente) for r in df("SELECT * FROM asistencia").itertuples()}
    tabla = est[["codigo", "nombre"]].rename(columns={"codigo": "Código", "nombre": "Nombre"})
    for r in cl.itertuples():
        col = []
        for e in est.id:
            v = asis.get((int(r.id), int(e)))
            col.append("Presente" if v == 1 else "Ausente" if v == 0 else "—")
        tabla[f"{r.fecha} {r.hora}"] = col
    cols = tabla.columns[2:]
    res = tabla[["Código", "Nombre"]].copy()
    res["Clases"] = (tabla[cols] != "—").sum(axis=1)
    res["Asistencias"] = (tabla[cols] == "Presente").sum(axis=1)
    res["Faltas"] = (tabla[cols] == "Ausente").sum(axis=1)
    res["% asistencia"] = (100 * res["Asistencias"] / res["Clases"].replace(0, np.nan)).round(1).fillna(0)
    res["Alerta"] = np.where(res["Faltas"] > MAX_FALTAS, "⚠ más de 3 faltas", "")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tabla.to_excel(w, sheet_name="Asistencia", index=False)
        res.to_excel(w, sheet_name="Resumen", index=False)
        for hoja in ("Asistencia", "Resumen"):
            ws = w.sheets[hoja]
            ws.freeze_panes = "C2"
            for c in ws[1]:
                c.fill = PatternFill("solid", fgColor="0B3D91")
                c.font = Font(bold=True, color="FFFFFF")
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 14 if col[0].column_letter == "A" else 38 if col[0].column_letter == "B" else 17
            for fila in ws.iter_rows(min_row=2):
                for c in fila:
                    if c.value == "Ausente":
                        c.fill = PatternFill("solid", fgColor="F8B4B4")
                    elif c.value == "Presente":
                        c.fill = PatternFill("solid", fgColor="B7E4C7")
        ws = w.sheets["Resumen"]
        i = list(res.columns).index("Faltas")
        for fila in ws.iter_rows(min_row=2):
            n = fila[i].value or 0
            fila[i].fill = PatternFill("solid", fgColor="B7E4C7" if n <= 1 else "FFF3B0" if n <= MAX_FALTAS else "F8B4B4")
    return buf.getvalue()


def editor(tabla, key):
    return st.data_editor(tabla, key=key, hide_index=True, disabled=["Foto", "codigo", "nombre"],
                          column_config={"Foto": st.column_config.ImageColumn("Foto", width="small"), "codigo": "Código", "nombre": "Nombre",
                                         "Presente": st.column_config.CheckboxColumn("Presente")})


def con_foto(tabla, fotos):
    tabla.insert(0, "Foto", [uri_foto(b) for b in fotos])
    return tabla


def pdf_reporte(tabla, subtitulo):
    from fpdf import FPDF

    def t(x):
        return str(x).encode("latin-1", "replace").decode("latin-1")

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(False)
    pdf.add_page()
    x0 = 10
    if ESCUDO:
        try:
            pdf.image(ESCUDO, x=10, y=8, h=18)
            x0 = 32
        except Exception:
            pass
    pdf.set_xy(x0, 9)
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 8, t("ESAVE Checklist - Reporte de faltas"))
    pdf.set_xy(x0, 17)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, t(f"{subtitulo} - Generado: {ahora():%Y-%m-%d %H:%M}"))
    anchos = [24, 66, 70, 17, 24, 17, 24, 35]

    def encabezado_tabla():
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(11, 61, 145)
        pdf.set_text_color(255, 255, 255)
        for ancho, nombre in zip(anchos, tabla.columns):
            pdf.cell(ancho, 8, t(nombre), border=1, fill=True, align="C")
        pdf.ln()
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Helvetica", "", 9)

    pdf.set_y(30)
    encabezado_tabla()
    for fila in tabla.itertuples(index=False):
        if pdf.get_y() > 192:
            pdf.add_page()
            pdf.set_y(12)
            encabezado_tabla()
        faltas = int(fila[5])
        pdf.set_fill_color(*((200, 230, 201) if faltas <= 1 else (255, 243, 196) if faltas <= MAX_FALTAS else (248, 201, 201)))
        for ancho, valor in zip(anchos, fila):
            pdf.cell(ancho, 7, t(valor)[:int(ancho / 1.9)], border=1, fill=True)
        pdf.ln()
    return bytes(pdf.output())


# ---------------- Gráficas ----------------
VERDE, AMARILLO, ROJO = "#22A06B", "#F5B400", "#E5484D"


def grafica(ch):
    try:
        st.altair_chart(ch, width="stretch")
    except TypeError:
        st.altair_chart(ch, use_container_width=True)


def color_pct(v):
    return VERDE if v >= 85 else AMARILLO if v >= 70 else ROJO


def materias_html(g):
    filas = []
    for r in g.sort_values("pct", ascending=False).itertuples():
        c = color_pct(r.pct)
        filas.append(f"<div class='mat'><div class='mat-top'><span>{html.escape(str(r.materia))}</span><b style='color:{c}'>{r.pct:.0f}%</b></div>"
                     f"<div class='bar'><div style='width:{min(100, max(2, r.pct)):.0f}%;background:{c}'></div></div></div>")
    return "".join(filas)


def riesgo_html(r):
    filas = []
    for x in r.itertuples():
        e = estado_faltas(x.faltas)
        bg, fg = (("#12532f", "#c9f5da") if e == "Al día" else ("#6b4e00", "#ffe9a8") if e == "En riesgo" else ("#6b1d20", "#ffd0d2"))
        filas.append(f"<div class='riesgo'><span>{html.escape(str(x.nombre))}</span>"
                     f"<span class='chip' style='background:{bg};color:{fg}'>{e} · {int(x.faltas)} falta{'s' if x.faltas != 1 else ''}</span></div>")
    return "".join(filas)


def mapa_calor(d):
    d = d.copy()
    d["Clase"] = d["fecha"].str[5:] + " " + d["hora"]
    d["Asistencia"] = np.where(d["presente"] == 1, "Presente", "Ausente")
    orden = list(dict.fromkeys(d.sort_values(["fecha", "hora"])["Clase"]))
    ch = alt.Chart(d).mark_rect(stroke="#0A1628", strokeWidth=2, cornerRadius=3).encode(
        x=alt.X("Clase:O", sort=orden, title=None, axis=alt.Axis(labelAngle=-45)),
        y=alt.Y("nombre:N", title=None, axis=alt.Axis(labelLimit=220)),
        color=alt.Color("Asistencia:N", scale=alt.Scale(domain=["Presente", "Ausente"], range=[VERDE, ROJO]),
                        legend=alt.Legend(orient="bottom", title=None)),
        tooltip=[alt.Tooltip("nombre:N", title="Estudiante"), alt.Tooltip("Clase:N"), alt.Tooltip("Asistencia:N")])
    return ch.properties(height=max(120, 26 * d["nombre"].nunique() + 70), width="container")


# ---------------- Estilo y bienvenida ----------------
@st.cache_data
def escudo_uri(lado=160):
    im = Image.open(ESCUDO).convert("RGBA")
    im.thumbnail((lado, lado))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


@st.cache_data
def icono_inicio_uri():
    im = Image.open(ESCUDO).convert("RGBA")
    im.thumbnail((150, 150))
    lienzo = Image.new("RGBA", (180, 180), (255, 255, 255, 255))
    lienzo.paste(im, ((180 - im.width) // 2, (180 - im.height) // 2), im)
    buf = io.BytesIO()
    lienzo.convert("RGB").save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


@st.cache_data
def fondo_uri(ruta):
    """Foto de fondo del login, desenfocada aquí mismo (sirve con cualquier imagen que pongas)."""
    im = Image.open(ruta).convert("RGB")
    im.thumbnail((1280, 1280))
    im = im.filter(ImageFilter.GaussianBlur(5))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=72)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def fondo_login():
    if not FONDO:
        return
    capa = "linear-gradient(rgba(8,20,45,.45),rgba(8,20,45,.68))"
    movil = (f"@media (max-aspect-ratio:1/1){{.stApp,[data-testid=\"stApp\"]{{background:{capa},url(\"{fondo_uri(FONDO_MOVIL)}\") center/cover no-repeat !important}}}}"
             if FONDO_MOVIL else "")
    st.markdown(f"""<style>
    .stApp,[data-testid="stApp"]{{background:{capa},url("{fondo_uri(FONDO)}") center/cover no-repeat !important}}
    {movil}
    [data-testid="stForm"]{{background:rgba(14,30,58,.55);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);border:1px solid rgba(255,255,255,.18);box-shadow:0 12px 40px rgba(0,0,0,.4)}}
    </style>""", unsafe_allow_html=True)


def fecha_texto():
    h = ahora()
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    return f"{dias[h.weekday()].capitalize()} {h.day} de {MESES[h.month - 1].lower()}"


def estilo():
    st.markdown("""<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700;800&display=swap');
    @keyframes fadein{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
    .block-container{padding-top:4.5rem;max-width:960px}
    h1,h2,h3,h4{font-family:'Montserrat',sans-serif;letter-spacing:-.01em}
    [data-testid="stSidebar"]{background:#0D1B33;border-right:1px solid #1E3A66}
    [data-baseweb="tab-list"]{gap:8px;overflow-x:auto;padding-bottom:6px}
    [data-baseweb="tab-highlight"],[data-baseweb="tab-border"]{display:none}
    button[role="tab"]{background:#12233F;border:1px solid #1E3A66;border-radius:999px;padding:8px 16px;height:auto;font-family:'Montserrat',sans-serif;font-weight:600;transition:all .2s}
    button[role="tab"] p{color:#B8C7E0;font-weight:600}
    button[role="tab"]:hover{border-color:#3B82F6}
    button[role="tab"][aria-selected="true"]{background:linear-gradient(135deg,#1D4ED8,#3B82F6);border-color:transparent;box-shadow:0 4px 14px rgba(59,130,246,.45)}
    button[role="tab"][aria-selected="true"] p{color:#fff}
    [data-baseweb="tab-panel"]{animation:fadein .4s}
    [data-testid="stMetric"]{background:linear-gradient(160deg,#14284A,#0F1F3A);border:1px solid #1E3A66;border-radius:16px;padding:14px 16px;box-shadow:0 6px 18px rgba(0,0,0,.25)}
    [data-testid="stMetricValue"]{font-family:'Montserrat',sans-serif;font-weight:700}
    [data-testid="stForm"]{background:#12233F;border:1px solid #1E3A66;border-radius:18px;padding:18px;box-shadow:0 8px 24px rgba(0,0,0,.25)}
    [data-testid="stExpander"]{border:1px solid #1E3A66;border-radius:14px;background:#0F1F3A}
    [data-testid="stAlert"]{border-radius:14px}
    .stButton>button,.stDownloadButton>button,[data-testid="stFormSubmitButton"]>button{border-radius:12px;font-weight:600;transition:transform .15s,box-shadow .15s}
    .stButton>button:hover,.stDownloadButton>button:hover,[data-testid="stFormSubmitButton"]>button:hover{transform:translateY(-1px);box-shadow:0 6px 16px rgba(59,130,246,.35)}
    .esave-head{display:flex;align-items:center;gap:18px;background:radial-gradient(circle at 92% 0%,rgba(59,130,246,.45),transparent 45%),linear-gradient(135deg,#0B2E6B,#123C8A 60%,#1D4ED8);border:1px solid #2B4F8F;border-radius:20px;padding:18px 22px;margin-bottom:16px;box-shadow:0 10px 30px rgba(0,0,0,.35);animation:fadein .5s}
    .esave-head img{height:76px;filter:drop-shadow(0 4px 10px rgba(0,0,0,.45))}
    .esave-head h1{margin:0;padding:0;font-family:'Montserrat',sans-serif;font-weight:800;font-size:1.9rem;color:#fff;letter-spacing:.02em}
    .esave-head p{margin:4px 0 0;color:#BFD4F5;font-size:.9rem}
    .esave-fecha{margin-left:auto;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);border-radius:999px;padding:6px 14px;font-size:.8rem;color:#E8F0FE;white-space:nowrap}
    @media (max-width:640px){.esave-fecha{display:none}.esave-head{padding:14px}.esave-head img{height:56px}.esave-head h1{font-size:1.3rem}}
    .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:10px 0 16px}
    .kpi{background:linear-gradient(160deg,#16305A,#0F2243);border:1px solid #24478A;border-radius:18px;padding:16px;box-shadow:0 8px 20px rgba(0,0,0,.28);animation:fadein .5s}
    .kpi .ico{font-size:1.4rem}
    .kpi .val{font-family:'Montserrat',sans-serif;font-weight:800;font-size:2rem;color:#fff;line-height:1.15}
    .kpi .lab{color:#9DB4DB;font-size:.75rem;text-transform:uppercase;letter-spacing:.07em}
    .kpi.alerta .val{color:#FF8A8E}
    .mat,.riesgo{background:#0F1F3A;border:1px solid #1E3A66;border-radius:14px;padding:10px 14px;margin-bottom:8px;animation:fadein .5s}
    .mat-top,.riesgo{display:flex;justify-content:space-between;align-items:center;gap:8px;font-size:.92rem}
    .bar{height:8px;background:#1C3157;border-radius:99px;margin-top:8px;overflow:hidden}
    .bar div{height:100%;border-radius:99px}
    .chip{border-radius:999px;padding:4px 12px;font-size:.8rem;font-weight:600;white-space:nowrap}
    </style>""", unsafe_allow_html=True)


def encabezado():
    logo = f"<img src='{escudo_uri()}'>" if ESCUDO else "<span style='font-size:52px'>✈️</span>"
    st.markdown(f"<div class='esave-head'>{logo}<div><h1>{APP}</h1>"
                "<p>Escuela de Aviación del Ejército</p></div>"
                f"<div class='esave-fecha'>📅 {fecha_texto()}</div></div>", unsafe_allow_html=True)


def icono_pantalla_inicio():
    """Pone el escudo como ícono al usar 'Añadir a pantalla de inicio' (mejor esfuerzo)."""
    if not ESCUDO:
        return
    components.html(f"""<script>
    for (const w of [window.parent, window.top]) {{
      try {{
        const d = w.document;
        let l = d.querySelector("link[rel='apple-touch-icon']");
        if (!l) {{ l = d.createElement('link'); l.rel = 'apple-touch-icon'; d.head.appendChild(l); }}
        l.href = "{icono_inicio_uri()}";
        let m = d.querySelector("meta[name='apple-mobile-web-app-title']");
        if (!m) {{ m = d.createElement('meta'); m.name = 'apple-mobile-web-app-title'; d.head.appendChild(m); }}
        m.content = "{APP}";
      }} catch (e) {{}}
    }}
    </script>""", height=0)


# ---------------- Pantallas ----------------
def pantalla_login():
    estilo()
    fondo_login()
    encabezado()
    icono_pantalla_inicio()
    mostrar_aviso()
    _, centro, _ = st.columns([1, 2, 1])
    with centro:
        st.subheader("Iniciar sesión")
        with st.form("login"):
            u = st.text_input("Usuario")
            c = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Entrar", type="primary"):
                usr = verificar(u, c)
                if usr:
                    st.session_state.usr = usr
                    st.query_params["s"] = crear_sesion(usr["id"])
                    st.rerun()
                st.error("Usuario o contraseña incorrectos")
        st.caption("¿Olvidaste tu contraseña? Pídele una nueva al administrador.")


def tab_inicio(usr):
    hoy = ahora()
    dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
    st.markdown(f"### Bienvenido, {usr['nombre']}")
    st.caption(f"{dias[hoy.weekday()].capitalize()} {hoy.day} de {MESES[hoy.month - 1].lower()} de {hoy.year}")
    mm = mis_materias(usr)
    mids = [int(i) for i in mm.id] if not mm.empty else []
    if not mids:
        st.info("Aún no tienes materias asignadas. Pídele al administrador que te las asigne.")
        return
    cl = df(f"""SELECT c.hora, m.nombre, (SELECT COUNT(*) FROM asistencia a WHERE a.clase_id=c.id) AS n
        FROM clases c JOIN materias m ON m.id=c.materia_id WHERE c.fecha=? AND c.materia_id IN ({marcas(mids)}) ORDER BY c.hora""",
            (str(hoy.date()), *mids))
    st.write("**Clases de hoy**")
    if cl.empty:
        st.caption("No hay clases programadas para hoy. Puedes programarlas en la pestaña Calendario.")
    for r in cl.itertuples():
        st.write(f"{'✅' if r.n else '🕒'} {r.hora} · {r.nombre} ({'lista tomada' if r.n else 'pendiente'})")
    sems = sorted({int(x) for x in mm.semestre})
    n_est = q(f"SELECT COUNT(*) AS n FROM estudiantes WHERE semestre IN ({marcas(sems)})", tuple(sems))[0]["n"]
    n_cl = q(f"SELECT COUNT(DISTINCT c.id) AS n FROM clases c WHERE c.materia_id IN ({marcas(mids)}) AND c.id IN (SELECT clase_id FROM asistencia)", tuple(mids))[0]["n"]
    f = faltas_df(mids)
    pct = f"{100 * f.asistencias.sum() / f.clases.sum():.1f}%" if not f.empty else "—"
    riesgo = int(f[f.faltas > MAX_FALTAS].codigo.nunique()) if not f.empty else 0
    tiles = [("👥", int(n_est), "Estudiantes", ""), ("📋", int(n_cl), "Listas tomadas", ""), ("📈", pct, "Asistencia general", ""),
             ("🚨", riesgo, "En alerta", " alerta" if riesgo else "")]
    st.markdown("<div class='kpis'>" + "".join(f"<div class='kpi{c}'><div class='ico'>{i}</div><div class='val'>{v}</div><div class='lab'>{l}</div></div>"
                                                for i, v, l, c in tiles) + "</div>", unsafe_allow_html=True)
    for a in alertas(mids):
        st.error(a)
    if not f.empty:
        g = f.groupby("materia")[["asistencias", "clases"]].sum()
        g["pct"] = (100 * g.asistencias / g.clases).round(1)
        st.markdown("### Asistencia por materia")
        st.markdown(materias_html(g.reset_index()), unsafe_allow_html=True)
        st.caption("🟢 85% o más · 🟡 70% a 84% · 🔴 menos de 70%")
    else:
        st.caption("Cuando guardes la primera lista, aquí verás las estadísticas.")
    if usr["rol"] == "admin":
        st.info("💾 Recuerda descargar el respaldo en Admin → Respaldo de datos antes de cada entrega.")


def tab_lista(usr):
    mm = mis_materias(usr)
    if mm.empty:
        st.info("Aún no tienes materias asignadas. Pídele al administrador que te las asigne.")
        return
    etq = etiquetas(mm)
    mid = st.selectbox("Materia", list(etq), format_func=etq.get)
    sem = int(mm[mm.id == mid].semestre.iloc[0])
    c1, c2 = st.columns(2)
    fecha = c1.date_input("Fecha", ahora().date())
    if "hora0" not in st.session_state:
        st.session_state.hora0 = ahora().time().replace(second=0, microsecond=0)
    hora = c2.time_input("Hora", st.session_state.hora0)
    roster = df("SELECT id, codigo, nombre, foto FROM estudiantes WHERE semestre=? ORDER BY nombre", (sem,))
    if roster.empty:
        st.warning(f"No hay estudiantes registrados en semestre {sem}. El administrador debe registrarlos primero.")
        return
    fotos = st.file_uploader("Foto(s) del salón (puedes tomarla con la cámara)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    with st.expander("O usar la cámara directamente"):
        cam = st.camera_input("Tomar foto")
    if cam:
        fotos = list(fotos) + [cam]
    umbral = st.slider("Estricto ← → Flexible", 0.25, 0.70, 0.363, 0.01, help="Valor alto = más estricto (menos errores, pero puede dejar a alguien sin reconocer).")
    b1, b2 = st.columns(2)
    if b1.button("🔍 Reconocer rostros", type="primary", disabled=not fotos):
        presentes, imgs, desconocidos = set(), [], 0
        with st.spinner("Analizando..."):
            for f in fotos:
                img = leer_imagen(f)
                if img is None:
                    continue
                caras, vect = rostros(img)
                res = reconocer(vect, umbral, sem)
                presentes |= {r[0] for r in res if r}
                desconocidos += sum(r is None for r in res)
                imgs.append(anotar(img, caras, res))
        st.session_state.reco = dict(mid=mid, presentes=presentes, imgs=imgs, desconocidos=desconocidos, n=secrets.token_hex(3))
    if b2.button("✍️ Llamar manualmente"):
        st.session_state.reco = dict(mid=mid, presentes=set(), imgs=[], desconocidos=0, n=secrets.token_hex(3))
    r = st.session_state.get("reco")
    if not r or r["mid"] != mid:
        return
    for im in r["imgs"]:
        st.image(im)
    if r["imgs"]:
        st.write(f"✅ Reconocidos: **{len(r['presentes'])}** · ❓ Rostros sin identificar: **{r['desconocidos']}**")
    st.caption("Revisa la lista y corrige a mano lo que haga falta antes de guardar.")
    tabla = con_foto(roster.set_index("id")[["codigo", "nombre"]], roster["foto"])
    tabla["Presente"] = [i in r["presentes"] for i in tabla.index]
    editada = editor(tabla, f"ed_{r['n']}")
    pres = int(editada["Presente"].sum())
    st.info(f"Resumen antes de guardar: {pres} presentes · {len(editada) - pres} ausentes")
    if len(editada) - pres:
        with st.expander("Ver quiénes aparecen ausentes"):
            for n in editada.loc[~editada["Presente"], "nombre"]:
                st.write(f"• {n}")
    if st.button("💾 Guardar asistencia", type="primary"):
        cid = obtener_clase(mid, usr["id"], str(fecha), hora.strftime("%H:%M"))
        filas = []
        for eid, pres in editada["Presente"].items():
            orig = eid in r["presentes"]
            metodo = "facial" if (pres and orig) else "manual" if (pres != orig) else "auto"
            filas.append((cid, int(eid), int(bool(pres)), metodo))
        muchos("INSERT OR REPLACE INTO asistencia(clase_id,estudiante_id,presente,metodo) VALUES(?,?,?,?)", filas)
        total = sum(f[2] for f in filas)
        avisar(f"Asistencia guardada: {total} presentes, {len(filas) - total} ausentes.")
        st.session_state.pop("reco", None)
        st.rerun()


def tab_calendario(usr):
    mm = mis_materias(usr)
    if mm.empty:
        st.info("Aún no tienes materias asignadas.")
        return
    etq = etiquetas(mm)
    hoy = ahora().date()
    c1, c2 = st.columns(2)
    anio = c1.number_input("Año", 2024, 2100, hoy.year)
    mes = c2.selectbox("Mes", range(1, 13), index=hoy.month - 1, format_func=lambda m: MESES[m - 1])
    ids = list(etq)
    cl = df(f"""SELECT c.fecha, c.hora, m.nombre, (SELECT COUNT(*) FROM asistencia a WHERE a.clase_id=c.id) AS n
        FROM clases c JOIN materias m ON m.id=c.materia_id WHERE c.fecha LIKE ? AND c.materia_id IN ({marcas(ids)})
        ORDER BY c.hora""", (f"{int(anio)}-{mes:02d}-%", *ids))
    por_dia = {}
    for r in cl.itertuples():
        por_dia.setdefault(int(r.fecha[8:]), []).append(r)
    h = "<table style='width:100%;border-collapse:collapse;table-layout:fixed;font-size:10px'><tr>"
    h += "".join(f"<th>{d}</th>" for d in ["L", "M", "X", "J", "V", "S", "D"]) + "</tr>"
    for semana in calendar.Calendar(0).monthdayscalendar(int(anio), mes):
        h += "<tr>"
        for d in semana:
            if d == 0:
                h += "<td style='border:1px solid #8883;height:60px'></td>"
                continue
            fondo = "background:#1e90ff33;" if (int(anio), mes, d) == (hoy.year, hoy.month, hoy.day) else ""
            items = "".join(
                f"<div style='background:{'#2ecc7155' if r.n else '#f1c40f55'};border-radius:3px;margin-top:2px;padding:0 2px;"
                f"overflow:hidden;white-space:nowrap;text-overflow:ellipsis'>{r.hora} {html.escape(r.nombre)}</div>"
                for r in por_dia.get(d, []))
            h += f"<td style='border:1px solid #8883;vertical-align:top;height:60px;padding:2px;{fondo}'><b>{d}</b>{items}</td>"
        h += "</tr>"
    st.markdown(h + "</table>", unsafe_allow_html=True)
    st.caption("🟩 lista tomada · 🟨 clase programada · azul = hoy")
    with st.expander("➕ Programar una clase"):
        mid = st.selectbox("Materia", ids, format_func=etq.get, key="prog_m")
        f1, f2 = st.columns(2)
        fecha = f1.date_input("Fecha", hoy, key="prog_f")
        hora = f2.time_input("Hora", dt.time(7, 0), key="prog_h")
        if st.button("Programar"):
            obtener_clase(mid, usr["id"], str(fecha), hora.strftime("%H:%M"))
            st.rerun()


def tab_corregir(usr):
    mm = mis_materias(usr)
    if mm.empty:
        st.info("Aún no tienes materias asignadas.")
        return
    etq = etiquetas(mm)
    mid = st.selectbox("Materia", list(etq), format_func=etq.get, key="cor_m")
    sem = int(mm[mm.id == mid].semestre.iloc[0])
    cl = df("SELECT id, fecha, hora FROM clases WHERE materia_id=? ORDER BY fecha DESC, hora DESC", (int(mid),))
    if cl.empty:
        st.info("Aún no hay clases registradas para esta materia.")
        return
    op = {int(r.id): f"{r.fecha} · {r.hora}" for r in cl.itertuples()}
    cid = st.selectbox("Clase", list(op), format_func=op.get)
    act = df("""SELECT e.id, e.codigo, e.nombre, e.foto, a.presente FROM estudiantes e
        LEFT JOIN asistencia a ON a.estudiante_id=e.id AND a.clase_id=? WHERE e.semestre=? ORDER BY e.nombre""", (int(cid), sem))
    if act.empty:
        st.info("No hay estudiantes en este semestre.")
        return
    act["presente"] = act["presente"].fillna(0).astype(int)
    tabla = con_foto(act.set_index("id")[["codigo", "nombre"]], act["foto"])
    tabla["Presente"] = act.set_index("id")["presente"].astype(bool)
    ver_ed = st.session_state.get(f"cor_v_{cid}", 0)
    editada = editor(tabla, f"cor_{cid}_{ver_ed}")
    if st.button("💾 Guardar correcciones", type="primary"):
        for eid, pres in editada["Presente"].items():
            if bool(pres) != bool(tabla.loc[eid, "Presente"]):
                q("INSERT OR REPLACE INTO asistencia VALUES(?,?,?,?)", (int(cid), int(eid), int(bool(pres)), "manual"))
            else:
                q("INSERT OR IGNORE INTO asistencia VALUES(?,?,?,?)", (int(cid), int(eid), int(bool(pres)), "auto"))
        st.session_state[f"cor_v_{cid}"] = ver_ed + 1
        avisar("Correcciones guardadas.")
        st.rerun()
    # --- eliminar la asistencia de esta clase ---
    with st.expander("🗑️ Eliminar la asistencia de esta clase"):
        st.caption("Borra los presentes y ausentes guardados de esta clase. Después puedes volver a tomar lista.")
        tambien = st.checkbox("Eliminar también la clase del calendario", key=f"del_cal_{cid}")
        if st.button("Eliminar asistencia de esta clase", key=f"del_as_{cid}"):
            st.session_state.conf_clase = int(cid)
        if st.session_state.get("conf_clase") == int(cid):
            st.error(f"¿Eliminar la asistencia del {op[cid]}? No se puede deshacer.")
            d1, d2 = st.columns(2)
            if d1.button("Sí, eliminar", type="primary", key=f"si_as_{cid}"):
                q("DELETE FROM asistencia WHERE clase_id=?", (int(cid),))
                if tambien:
                    q("DELETE FROM clases WHERE id=?", (int(cid),))
                st.session_state.pop("conf_clase", None)
                avisar("Asistencia eliminada." + (" La clase también salió del calendario." if tambien else ""), "🗑️")
                st.rerun()
            if d2.button("Cancelar", key=f"no_as_{cid}"):
                st.session_state.pop("conf_clase", None)
                st.rerun()


def semaforo(tabla):
    def color(fila):
        n = fila["Faltas"]
        c = "background-color:#12532f;color:#c9f5da" if n <= 1 else "background-color:#6b4e00;color:#ffe9a8" if n <= MAX_FALTAS else "background-color:#6b1d20;color:#ffd0d2"
        return [c] * len(fila)
    return tabla.style.apply(color, axis=1, subset=["Faltas", "Estado"])


def tab_reportes(usr):
    mm = mis_materias(usr)
    if mm.empty:
        st.info("Aún no tienes materias asignadas.")
        return
    etq = etiquetas(mm)
    op = {0: "Todas mis materias", **etq}
    sel = st.selectbox("Materia", list(op), format_func=op.get, key="rep_m")
    mids = list(etq) if sel == 0 else [sel]
    for a in alertas(mids):
        st.error(a)
    c1, c2 = st.columns(2)
    buscar = c1.text_input("🔎 Buscar estudiante (nombre o código)", key="rep_b")
    usar = c2.checkbox("Filtrar por fechas", key="rep_uf")
    desde = hasta = None
    if usar:
        d1, d2 = st.columns(2)
        desde = d1.date_input("Desde", ahora().date().replace(day=1), key="rep_d")
        hasta = d2.date_input("Hasta", ahora().date(), key="rep_h")
    f = faltas_df(mids, desde, hasta)
    if f.empty:
        st.info("Todavía no hay asistencia guardada" + (" en ese periodo." if usar else "."))
        return
    if buscar.strip():
        b = sin_tildes(buscar.strip().lower())
        f = f[f.nombre.map(lambda x: b in sin_tildes(x.lower())) | f.codigo.str.contains(buscar.strip(), case=False, regex=False)]
        if f.empty:
            st.info("Ningún estudiante coincide con la búsqueda.")
            return
    f["% asistencia"] = (100 * f.asistencias / f.clases).round(1)
    f["Estado"] = f.faltas.map(estado_faltas)
    tabla = f.rename(columns={"codigo": "Código", "nombre": "Estudiante", "materia": "Materia", "clases": "Clases",
                              "asistencias": "Asistencias", "faltas": "Faltas"})[
        ["Código", "Estudiante", "Materia", "Clases", "Asistencias", "Faltas", "% asistencia", "Estado"]]
    st.subheader("Reporte de faltas")
    m1, m2, m3 = st.columns(3)
    m1.metric("🟢 Al día", int((f.Estado == "Al día").sum()))
    m2.metric("🟡 En riesgo", int((f.Estado == "En riesgo").sum()))
    m3.metric("🔴 Alerta", int((f.Estado == "Alerta").sum()))
    try:
        st.dataframe(semaforo(tabla), hide_index=True)
    except Exception:
        st.dataframe(tabla, hide_index=True)
    st.caption(f"🟢 0 a 1 faltas · 🟡 2 a {MAX_FALTAS} faltas · 🔴 más de {MAX_FALTAS} faltas")
    st.subheader("Estudiantes con más faltas")
    r = f.groupby("nombre", as_index=False)["faltas"].sum().sort_values("faltas", ascending=False).head(10)
    r = r[r["faltas"] > 0]
    if r.empty:
        st.success("Nadie tiene faltas en este periodo.")
    else:
        st.markdown(riesgo_html(r), unsafe_allow_html=True)
    st.subheader("Mapa de calor de asistencia")
    if sel == 0:
        st.caption("Elige una materia para ver el mapa de calor.")
    else:
        cond, args = "", [int(sel)]
        if desde:
            cond += " AND c.fecha>=?"
            args.append(str(desde))
        if hasta:
            cond += " AND c.fecha<=?"
            args.append(str(hasta))
        d = df(f"""SELECT e.nombre, e.codigo, c.fecha, c.hora, a.presente FROM asistencia a
            JOIN clases c ON c.id=a.clase_id JOIN estudiantes e ON e.id=a.estudiante_id
            WHERE c.materia_id=?{cond}""", tuple(args))
        d = d[d["codigo"].isin(f["codigo"])] if not d.empty else d
        if d.empty:
            st.caption("Todavía no hay datos para el mapa.")
        else:
            try:
                grafica(mapa_calor(d))
            except Exception as e:
                st.caption(f"No se pudo dibujar el mapa: {e}")
    with st.expander("🔎 Historial de un estudiante"):
        nombres = dict(zip(f.codigo, f.nombre))
        cod = st.selectbox("Estudiante", list(nombres), format_func=lambda c: f"{nombres[c]} ({c})", key="rep_hist")
        h = df(f"""SELECT c.fecha, c.hora, m.nombre AS materia, CASE a.presente WHEN 1 THEN 'Presente' ELSE 'Ausente' END AS estado, a.metodo
            FROM asistencia a JOIN clases c ON c.id=a.clase_id JOIN materias m ON m.id=c.materia_id JOIN estudiantes e ON e.id=a.estudiante_id
            WHERE e.codigo=? AND m.id IN ({marcas(mids)}) ORDER BY c.fecha DESC, c.hora DESC""", (cod, *mids))
        st.dataframe(h.rename(columns={"fecha": "Fecha", "hora": "Hora", "materia": "Materia", "estado": "Estado", "metodo": "Método"}), hide_index=True)
    try:
        st.download_button("⬇️ Reporte de faltas (PDF)", pdf_reporte(tabla, op[sel]), "reporte_faltas.pdf", "application/pdf")
    except Exception as e:
        st.caption(f"No se pudo generar el PDF: {e}")
    if sel == 0:
        st.caption("Elige una materia para descargar el Excel.")
        return
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    datos = excel_materia(sel)
    if datos:
        st.download_button("⬇️ Excel de toda la materia", datos, f"asistencia_{sel}.xlsx", mime)
    cl = df("SELECT id, fecha, hora FROM clases WHERE materia_id=? AND id IN (SELECT clase_id FROM asistencia) ORDER BY fecha DESC, hora DESC", (int(sel),))
    if not cl.empty:
        op2 = {int(r.id): f"{r.fecha} · {r.hora}" for r in cl.itertuples()}
        cid = st.selectbox("Descargar una sola clase", list(op2), format_func=op2.get)
        st.download_button("⬇️ Excel de esa clase", excel_materia(sel, cid), f"asistencia_{sel}_{cid}.xlsx", mime)


def tab_pensum():
    sem = st.selectbox("Semestre", range(1, 10))
    t = df("""SELECT m.nombre AS "Asignatura", m.creditos AS "Créditos", COALESCE(GROUP_CONCAT(u.nombre, ', '), '—') AS "Docente"
        FROM materias m LEFT JOIN asignaciones a ON a.materia_id=m.id LEFT JOIN usuarios u ON u.id=a.docente_id
        WHERE m.semestre=? GROUP BY m.id ORDER BY m.id""", (sem,))
    st.dataframe(t, hide_index=True)
    st.metric("Total de créditos", int(t["Créditos"].sum()))


def tab_estudiantes():
    est = df("SELECT codigo, nombre, semestre, foto, COALESCE(LENGTH(rostros)/512, 0) AS fotos FROM estudiantes ORDER BY semestre, nombre")
    st.metric("Estudiantes registrados", len(est))
    if not est.empty:
        buscar = st.text_input("🔎 Buscar por nombre o código", key="est_buscar")
        ver = est
        if buscar.strip():
            b = sin_tildes(buscar.strip().lower())
            ver = est[est.nombre.map(lambda x: b in sin_tildes(x.lower())) | est.codigo.str.contains(buscar.strip(), case=False, regex=False)]
        st.dataframe(ver.drop(columns="foto").rename(columns={"codigo": "Código", "nombre": "Nombre", "semestre": "Semestre", "fotos": "Fotos de rostro"}), hide_index=True)
        galeria = [r for r in ver.itertuples() if isinstance(r.foto, (bytes, bytearray))]
        with st.expander("🖼️ Galería de rostros registrados"):
            if not galeria:
                st.caption("Aún no hay miniaturas. Se guardan al registrar un estudiante; para los registrados antes, súbeles fotos de nuevo.")
            cols = st.columns(4)
            for i, r in enumerate(galeria):
                cols[i % 4].image(bytes(r.foto), caption=" ".join(r.nombre.split()[:2]), width=90)
    with st.expander("➕ Registrar estudiante o agregar fotos", expanded=est.empty):
        st.caption("Sube 2 o 3 fotos de frente, con buena luz y una sola persona. Si el código ya existe, las fotos se suman a las que ya tiene.")
        with st.form("nuevo_est", clear_on_submit=True):
            codigo = st.text_input("Código / documento")
            nombre = st.text_input("Nombre completo")
            sem = st.number_input("Semestre", 1, 9, 1)
            fotos = st.file_uploader("Fotos del rostro", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
            if st.form_submit_button("Guardar estudiante", type="primary"):
                vecs, avisos, foto_b = [], [], None
                for f in fotos:
                    img = leer_imagen(f)
                    caras, v = rostros(img) if img is not None else (np.zeros((0, 15)), np.zeros((0, 128)))
                    if len(v) == 0:
                        avisos.append(f"{f.name}: no se detectó ningún rostro")
                        continue
                    if len(v) > 1:
                        avisos.append(f"{f.name}: había {len(v)} rostros, se usó el más grande")
                    idx = int(np.argmax(caras[:, 2] * caras[:, 3]))
                    vecs.append(v[idx])
                    if foto_b is None:
                        foto_b = miniatura(img, caras[idx])
                for a in avisos:
                    st.warning(a)
                if not codigo.strip() or not nombre.strip() or not vecs:
                    st.error("Faltan datos o ninguna foto sirvió.")
                else:
                    guardar_estudiante(codigo.strip(), nombre.strip(), sem, vecs, foto_b)
                    avisar(f"{nombre.strip()} guardado con {len(vecs)} foto(s).")
                    st.rerun()
    if not est.empty:
        with st.expander("🗑️ Eliminar estudiante"):
            cod = st.selectbox("Estudiante", est.codigo, format_func=lambda c: f"{c} · {est.set_index('codigo').nombre[c]}")
            if st.button("Eliminar"):
                e = q("SELECT id FROM estudiantes WHERE codigo=?", (cod,))[0]["id"]
                q("DELETE FROM asistencia WHERE estudiante_id=?", (e,))
                q("DELETE FROM estudiantes WHERE id=?", (e,))
                avisar("Estudiante eliminado.", "🗑️")
                st.rerun()


def acumular_cred(filas):
    nuevo = pd.DataFrame(filas)
    previo = st.session_state.get("cred")
    if previo is not None:
        previo = previo[~previo["Usuario"].isin(nuevo["Usuario"])]
        nuevo = pd.concat([previo, nuevo], ignore_index=True)
    st.session_state.cred = nuevo


def tab_admin():
    docs = df("""SELECT u.id, u.usuario, u.nombre, (SELECT COUNT(*) FROM asignaciones a WHERE a.docente_id=u.id) AS materias
        FROM usuarios u WHERE rol='docente' ORDER BY nombre""")
    st.subheader("Docentes")
    if not docs.empty:
        st.dataframe(docs.drop(columns="id").rename(columns={"usuario": "Usuario", "nombre": "Nombre", "materias": "Materias"}), hide_index=True)
    cred = st.session_state.get("cred")
    if cred is not None:
        st.success("Credenciales generadas. Guárdalas ahora.")
        st.dataframe(cred, hide_index=True)
        st.download_button("⬇️ Descargar credenciales (CSV)", cred.to_csv(index=False).encode("utf-8-sig"), "credenciales_docentes.csv", "text/csv")
        if st.button("Ocultar credenciales"):
            del st.session_state["cred"]
            st.rerun()
    if st.button("Cargar los docentes del plan de estudios (PDF)"):
        filas = []
        for nombre, usuario in DOCENTES_PDF:
            if not q("SELECT 1 FROM usuarios WHERE usuario=?", (usuario,)):
                c = clave_aleatoria()
                crear_usuario(usuario, nombre, "docente", c)
                filas.append({"Nombre": nombre, "Usuario": usuario, "Contraseña": c})
        if filas:
            acumular_cred(filas)
        st.rerun()
    if not docs.empty:
        with st.expander("🔑 Restablecer contraseñas de TODOS los docentes (CSV completo)"):
            st.warning("Esto genera una contraseña nueva para cada docente. Las anteriores dejan de funcionar.")
            ok = st.checkbox("Entiendo, quiero restablecer todas")
            if st.button("Restablecer todas y generar CSV", disabled=not ok):
                filas = []
                for _, r in docs.iterrows():
                    c = clave_aleatoria()
                    cambiar_clave(int(r["id"]), c)
                    filas.append({"Nombre": r["nombre"], "Usuario": r["usuario"], "Contraseña": c})
                st.session_state.pop("cred", None)
                acumular_cred(filas)
                st.rerun()
    with st.form("nuevo_doc", clear_on_submit=True):
        st.write("**Crear docente** (la contraseña se genera sola)")
        nom = st.text_input("Nombre completo")
        usu = st.text_input("Usuario (sin espacios)")
        if st.form_submit_button("Crear docente"):
            if nom.strip() and usu.strip() and not q("SELECT 1 FROM usuarios WHERE usuario=?", (usu.strip().lower(),)):
                c = clave_aleatoria()
                crear_usuario(usu, nom, "docente", c)
                acumular_cred([{"Nombre": nom, "Usuario": usu.strip().lower(), "Contraseña": c}])
                st.rerun()
            st.error("Faltan datos o el usuario ya existe.")
    if not docs.empty:
        nombres = dict(zip(map(int, docs.id), docs.nombre))
        did = st.selectbox("Docente", list(nombres), format_func=nombres.get)
        todas = df("SELECT id, semestre, nombre FROM materias ORDER BY semestre, id")
        etq = etiquetas(todas)
        act = [int(x) for x in df("SELECT materia_id FROM asignaciones WHERE docente_id=?", (did,)).get("materia_id", [])]
        sel = st.multiselect("Materias asignadas", list(etq), default=act, format_func=etq.get, key=f"asig_{did}")
        c1, c2, c3 = st.columns(3)
        if c1.button("Guardar asignación", type="primary"):
            q("DELETE FROM asignaciones WHERE docente_id=?", (did,))
            muchos("INSERT INTO asignaciones VALUES(?,?)", [(did, m) for m in sel])
            avisar("Asignación guardada.")
            st.rerun()
        if c2.button("Nueva contraseña"):
            c = clave_aleatoria()
            cambiar_clave(did, c)
            u = q("SELECT usuario FROM usuarios WHERE id=?", (did,))[0]["usuario"]
            acumular_cred([{"Nombre": nombres[did], "Usuario": u, "Contraseña": c}])
            st.rerun()
        if c3.button("🗑️ Eliminar docente"):
            st.session_state.conf_del = did
        if st.session_state.get("conf_del") == did:
            st.error(f"¿Eliminar a {nombres[did]}? Se borran su usuario y sus materias asignadas. No se puede deshacer.")
            d1, d2 = st.columns(2)
            if d1.button("Sí, eliminar", type="primary"):
                u = q("SELECT usuario FROM usuarios WHERE id=?", (did,))[0]["usuario"]
                q("DELETE FROM asignaciones WHERE docente_id=?", (did,))
                q("DELETE FROM sesiones WHERE usuario_id=?", (did,))
                q("DELETE FROM usuarios WHERE id=? AND rol='docente'", (did,))
                prev = st.session_state.get("cred")
                if prev is not None:
                    prev = prev[prev["Usuario"] != u]
                    if prev.empty:
                        del st.session_state["cred"]
                    else:
                        st.session_state.cred = prev
                st.session_state.pop("conf_del", None)
                avisar("Docente eliminado.", "🗑️")
                st.rerun()
            if d2.button("Cancelar"):
                st.session_state.pop("conf_del", None)
                st.rerun()
    with st.expander("➕ Agregar materia al plan"):
        with st.form("nueva_mat", clear_on_submit=True):
            s = st.number_input("Semestre", 1, 9, 1)
            n = st.text_input("Nombre de la materia")
            cr = st.number_input("Créditos", 1, 6, 3)
            if st.form_submit_button("Agregar") and n.strip():
                q("INSERT OR IGNORE INTO materias(semestre,nombre,creditos) VALUES(?,?,?)", (int(s), n.strip(), int(cr)))
                avisar("Materia agregada.")
                st.rerun()
    with st.expander("💾 Respaldo de datos"):
        st.caption("Descarga una copia de vez en cuando; sirve para restaurar si el servidor se reinicia.")
        with open(DB, "rb") as fh:
            st.download_button("Descargar copia (.db)", fh.read(), "asistencia_respaldo.db")
        sub = st.file_uploader("Restaurar copia", type=["db"], key="rest")
        if sub and st.button("Restaurar copia"):
            with open(DB, "wb") as fh:
                fh.write(sub.getvalue())
            st.session_state.clear()
            st.query_params.clear()
            avisar("Copia restaurada. Inicia sesión de nuevo.")
            st.rerun()


# ---------------- Principal ----------------
init_db()
if "usr" not in st.session_state:
    u_ses = usuario_de_sesion(st.query_params.get("s"))
    if u_ses:
        st.session_state.usr = u_ses
    else:
        pantalla_login()
        st.stop()
usr = st.session_state.usr
es_admin = usr["rol"] == "admin"
with st.sidebar:
    if ESCUDO:
        st.image(ESCUDO, width=90)
    st.write(f"👤 **{usr['nombre']}** ({usr['rol']})")
    nueva = st.text_input("Nueva contraseña", type="password")
    if st.button("Cambiar contraseña"):
        if len(nueva) >= 6:
            cambiar_clave(usr["id"], nueva)
            st.success("Contraseña actualizada.")
        else:
            st.error("Mínimo 6 caracteres.")
    if st.button("Cerrar sesión"):
        cerrar_sesion(st.query_params.get("s"))
        st.query_params.clear()
        st.session_state.clear()
        st.rerun()
estilo()
encabezado()
icono_pantalla_inicio()
mostrar_aviso()
mids = list(etiquetas(mis_materias(usr)))
avisos = alertas(mids)
if avisos:
    lineas = [a.replace("🚨 ", "• ") for a in avisos[:5]]
    if len(avisos) > 5:
        lineas.append(f"…y {len(avisos) - 5} más (míralos en Inicio)")
    st.error(f"🚨 {len(avisos)} alerta(s): más de {MAX_FALTAS} faltas\n\n" + "\n\n".join(lineas))
nombres_tabs = ["🏠 Inicio", "📸 Lista", "📅 Calendario", "✏️ Corregir", "📊 Reportes", "🎓 Pensum"] + (["👥 Estudiantes", "⚙️ Admin"] if es_admin else [])
tabs = st.tabs(nombres_tabs)
with tabs[0]:
    tab_inicio(usr)
with tabs[1]:
    tab_lista(usr)
with tabs[2]:
    tab_calendario(usr)
with tabs[3]:
    tab_corregir(usr)
with tabs[4]:
    tab_reportes(usr)
with tabs[5]:
    tab_pensum()
if es_admin:
    with tabs[6]:
        tab_estudiantes()
    with tabs[7]:
        tab_admin()
