import io, os, subprocess, sys, tempfile, zipfile
import streamlit as st

AQUI = os.path.dirname(os.path.abspath(__file__))
PLANTILLA = os.path.join(AQUI, "plantilla_REG_PRMP_15.docx")

st.set_page_config(page_title="Generador REG-PRMP-15", page_icon="📄")

# Contraseña opcional: si defines PASSWORD en Secrets, se pide antes de usar la app.
try:
    clave = st.secrets.get("PASSWORD", None)
except Exception:
    clave = None
if clave:
    if st.text_input("Contraseña", type="password") != clave:
        st.stop()

st.title("Generador REG-PRMP-15")
st.write("Sube el PDF (o CSV) de la orden de trabajo de MaintainX y descarga los Word.")

archivos = st.file_uploader("PDF o CSV de MaintainX", type=["pdf", "csv"], accept_multiple_files=True)


def procesar(nombre, datos, salida):
    ext = nombre.lower().rsplit(".", 1)[-1]
    script = "generar_reg_prmp15_pdf.py" if ext == "pdf" else "generar_reg_prmp15.py"
    ruta = os.path.join(salida, "_entrada_" + os.path.basename(nombre))
    with open(ruta, "wb") as f:
        f.write(datos)
    r = subprocess.run([sys.executable, os.path.join(AQUI, script), ruta, PLANTILLA, salida],
                       capture_output=True, text=True)
    os.remove(ruta)
    return r


if archivos:
    with tempfile.TemporaryDirectory() as tmp:
        salida = os.path.join(tmp, "salida")
        os.makedirs(salida)
        for a in archivos:
            with st.spinner(f"Procesando {a.name}..."):
                r = procesar(a.name, a.getvalue(), salida)
            if r.returncode != 0:
                st.error(f"Error con {a.name}")
                st.code(r.stderr[-1500:])
            with st.expander(f"Detalle: {a.name}"):
                st.code((r.stdout or "").strip() or "(sin salida)")

        docs = sorted(f for f in os.listdir(salida) if f.endswith(".docx"))
        if not docs:
            st.warning("No se generó ningún Word. Revisa que el archivo sea la orden completada de MaintainX.")
        else:
            st.success(f"Se generaron {len(docs)} archivo(s).")
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for d in docs:
                    z.write(os.path.join(salida, d), d)
            if len(docs) > 1:
                st.download_button("Descargar todo (.zip)", buf.getvalue(), "REG-PRMP-15.zip", "application/zip")
            for d in docs:
                with open(os.path.join(salida, d), "rb") as f:
                    st.download_button(f"Descargar {d}", f.read(), d,
                                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                       key=d)
