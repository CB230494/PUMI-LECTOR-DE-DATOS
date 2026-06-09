import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(
    page_title="Lector de Avances PAO / Programas",
    page_icon="📊",
    layout="wide"
)

MESES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

COLOR_AZUL = "#0B1F3A"
COLOR_DORADO = "#C9A227"
COLOR_GRIS = "#F2F4F7"

st.markdown(
    f"""
    <style>
    .main {{ background-color: #FFFFFF; }}
    .block-container {{ padding-top: 1.5rem; }}
    h1, h2, h3 {{ color: {COLOR_AZUL}; }}
    div[data-testid="stMetric"] {{
        background-color: {COLOR_GRIS};
        padding: 14px;
        border-radius: 12px;
        border-left: 6px solid {COLOR_DORADO};
    }}
    .stButton>button, .stDownloadButton>button {{
        border-radius: 10px;
        font-weight: 700;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# FUNCIONES DE LIMPIEZA Y DETECCIÓN
# =========================================================
def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def limpiar_numero(valor):
    if pd.isna(valor) or valor == "":
        return 0
    if isinstance(valor, str):
        valor = valor.replace("%", "").replace(",", ".").strip()
    try:
        return float(valor)
    except Exception:
        return 0


def normalizar_columnas(cols):
    nuevas = []
    for c in cols:
        c = limpiar_texto(c).upper()
        c = re.sub(r"\s+", " ", c)
        nuevas.append(c)
    return nuevas


def detectar_tipo_libro(archivo):
    xl = pd.ExcelFile(archivo)
    for hoja in xl.sheet_names:
        muestra = pd.read_excel(archivo, sheet_name=hoja, header=None, nrows=3)
        texto = " ".join(muestra.astype(str).fillna("").values.flatten()).upper()
        if "PROGRAMA" in texto and any(m in texto for m in MESES):
            return "DETALLE_REGIONAL"
        if "CÓDIGO" in texto or "CODIGO" in texto:
            return "RESUMEN_NACIONAL"
    return "DESCONOCIDO"


def nombre_delegacion_desde_hoja(nombre_hoja):
    nombre = limpiar_texto(nombre_hoja)
    if nombre.upper().startswith("TOTAL"):
        return nombre
    return nombre

# =========================================================
# LECTURA DEL EXCEL REGIONAL / DETALLADO
# =========================================================
def leer_detalle_regional(archivo):
    xl = pd.ExcelFile(archivo)
    registros = []

    for hoja in xl.sheet_names:
        if hoja.strip().upper().startswith("TOTAL"):
            continue

        raw = pd.read_excel(archivo, sheet_name=hoja, header=None)
        if raw.empty:
            continue

        fila_header = None
        for i in range(min(10, len(raw))):
            fila = " ".join(raw.iloc[i].astype(str).fillna("").str.upper().tolist())
            if "PROGRAMA" in fila and "META" in fila:
                fila_header = i
                break

        if fila_header is None:
            continue

        df = pd.read_excel(archivo, sheet_name=hoja, header=fila_header)
        df.columns = normalizar_columnas(df.columns)
        df = df.dropna(how="all")

        if len(df.columns) < 4:
            continue

        col_programa = df.columns[0]
        col_actividad = df.columns[1]

        col_meta = next((c for c in df.columns if "META" in c), None)
        col_avance = next((c for c in df.columns if "AVANCE" in c), None)

        if not col_meta or not col_avance:
            continue

        df[col_programa] = df[col_programa].ffill()

        for _, row in df.iterrows():
            actividad = limpiar_texto(row.get(col_actividad, ""))
            programa = limpiar_texto(row.get(col_programa, ""))
            if not actividad or actividad.upper() in ["NAN", "TOTAL", "TOTALES"]:
                continue

            meta = limpiar_numero(row.get(col_meta, 0))
            avance = limpiar_numero(row.get(col_avance, 0))

            item = {
                "Tipo": "Detalle Regional",
                "Delegación": nombre_delegacion_desde_hoja(hoja),
                "Programa": programa,
                "Actividad": actividad,
                "Meta": meta,
                "Avance": avance,
                "% Avance": (avance / meta) if meta else 0,
                "Pendiente": max(meta - avance, 0),
                "Estado": "Completa" if meta and avance >= meta else ("En avance" if avance > 0 else "Pendiente")
            }

            for mes in MESES:
                col_mes = next((c for c in df.columns if c == mes), None)
                item[mes.title()] = limpiar_numero(row.get(col_mes, 0)) if col_mes else 0

            registros.append(item)

    return pd.DataFrame(registros)

# =========================================================
# LECTURA DEL EXCEL NACIONAL / RESUMEN
# =========================================================
def leer_resumen_nacional(archivo):
    xl = pd.ExcelFile(archivo)
    registros = []

    for hoja in xl.sheet_names:
        raw = pd.read_excel(archivo, sheet_name=hoja, header=None)
        if raw.empty:
            continue

        fila_header = None
        for i in range(min(8, len(raw))):
            fila = " ".join(raw.iloc[i].astype(str).fillna("").str.upper().tolist())
            if ("CÓDIGO" in fila or "CODIGO" in fila) and "DELEGACIÓN" in fila:
                fila_header = i
                break

        if fila_header is None:
            continue

        df = pd.read_excel(archivo, sheet_name=hoja, header=fila_header)
        df.columns = normalizar_columnas(df.columns)
        df = df.dropna(how="all")

        columnas_base = ["CÓDIGO", "CODIGO", "DELEGACIÓN", "CANTÓN", "PROVINCIA", "PAÍS", "META", "AVANCE", "% AVANCE"]
        columnas_utiles = [c for c in df.columns if c in columnas_base]
        df = df[columnas_utiles].copy()

        if "CODIGO" in df.columns and "CÓDIGO" not in df.columns:
            df = df.rename(columns={"CODIGO": "CÓDIGO"})

        for _, row in df.iterrows():
            delegacion = limpiar_texto(row.get("DELEGACIÓN", ""))
            if not delegacion or delegacion.upper() in ["NAN"]:
                continue

            meta = limpiar_numero(row.get("META", 0))
            avance = limpiar_numero(row.get("AVANCE", 0))
            porc = row.get("% AVANCE", 0)
            porc = limpiar_numero(porc)
            if porc > 1:
                porc = porc / 100
            if porc == 0 and meta:
                porc = avance / meta

            registros.append({
                "Tipo": "Resumen Nacional",
                "Región / Hoja": hoja.strip(),
                "Código": limpiar_texto(row.get("CÓDIGO", "")),
                "Delegación": delegacion,
                "Cantón": limpiar_texto(row.get("CANTÓN", "")),
                "Provincia": limpiar_texto(row.get("PROVINCIA", "")),
                "País": limpiar_texto(row.get("PAÍS", "")),
                "Meta": meta,
                "Avance": avance,
                "% Avance": porc,
                "Pendiente": max(meta - avance, 0),
                "Estado": "Completa" if meta and avance >= meta else ("En avance" if avance > 0 else "Pendiente")
            })

    return pd.DataFrame(registros)

# =========================================================
# EXPORTACIONES
# =========================================================
def generar_excel(df, nombre_hoja="Datos filtrados"):
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja[:31])
        workbook = writer.book
        worksheet = writer.sheets[nombre_hoja[:31]]

        formato_titulo = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": COLOR_AZUL,
            "align": "center", "valign": "vcenter", "border": 1
        })
        formato_porcentaje = workbook.add_format({"num_format": "0.00%"})
        formato_numero = workbook.add_format({"num_format": "#,##0"})

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, formato_titulo)
            ancho = min(max(len(str(value)) + 4, 14), 55)
            worksheet.set_column(col_num, col_num, ancho)

        for idx, col in enumerate(df.columns):
            if "%" in col:
                worksheet.set_column(idx, idx, 14, formato_porcentaje)
            if col in ["Meta", "Avance", "Pendiente"] or col.title() in [m.title() for m in MESES]:
                worksheet.set_column(idx, idx, 14, formato_numero)

        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        worksheet.freeze_panes(1, 0)

    salida.seek(0)
    return salida


def generar_pdf(df, titulo="Reporte de datos filtrados"):
    salida = io.BytesIO()
    doc = SimpleDocTemplate(
        salida,
        pagesize=landscape(letter),
        rightMargin=24,
        leftMargin=24,
        topMargin=24,
        bottomMargin=24
    )

    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle(
        "Titulo",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=16,
        textColor=colors.HexColor(COLOR_AZUL),
        spaceAfter=12
    )
    normal = ParagraphStyle(
        "NormalCustom",
        parent=styles["Normal"],
        alignment=TA_LEFT,
        fontSize=7,
        leading=9
    )

    elementos = []
    elementos.append(Paragraph(titulo, titulo_style))
    elementos.append(Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y')}", normal))
    elementos.append(Spacer(1, 0.15 * inch))

    resumen = f"Total registros: {len(df)} | Meta: {df['Meta'].sum():,.0f} | Avance: {df['Avance'].sum():,.0f} | Avance global: {(df['Avance'].sum() / df['Meta'].sum() if df['Meta'].sum() else 0):.2%}"
    elementos.append(Paragraph(resumen, normal))
    elementos.append(Spacer(1, 0.15 * inch))

    columnas_pdf = [c for c in df.columns if c in [
        "Región / Hoja", "Código", "Delegación", "Programa", "Actividad", "Cantón", "Provincia", "Meta", "Avance", "% Avance", "Pendiente", "Estado"
    ]]
    tabla_df = df[columnas_pdf].copy().head(80)

    if "% Avance" in tabla_df.columns:
        tabla_df["% Avance"] = tabla_df["% Avance"].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "0.00%")

    for col in ["Meta", "Avance", "Pendiente"]:
        if col in tabla_df.columns:
            tabla_df[col] = tabla_df[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "0")

    data = [list(tabla_df.columns)]
    for _, row in tabla_df.iterrows():
        data.append([Paragraph(limpiar_texto(v), normal) for v in row.tolist()])

    tabla = Table(data, repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COLOR_AZUL)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 7),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
    ]))
    elementos.append(tabla)

    if len(df) > 80:
        elementos.append(Spacer(1, 0.12 * inch))
        elementos.append(Paragraph("Nota: el PDF muestra los primeros 80 registros filtrados. El Excel descargable incluye todos los registros.", normal))

    doc.build(elementos)
    salida.seek(0)
    return salida

# =========================================================
# INTERFAZ PRINCIPAL
# =========================================================
st.title("📊 Lector de Avances por Delegación, Programa y Nivel Nacional")
st.caption("Carga el Excel regional o el Excel nacional. La app detecta la estructura y permite filtrar, consultar y descargar resultados.")

archivo = st.file_uploader("📁 Subir archivo Excel", type=["xlsx", "xlsm", "xls"])

if not archivo:
    st.info("Sube un archivo Excel para iniciar la lectura.")
    st.stop()

try:
    tipo = detectar_tipo_libro(archivo)
    archivo.seek(0)

    if tipo == "DETALLE_REGIONAL":
        df = leer_detalle_regional(archivo)
    elif tipo == "RESUMEN_NACIONAL":
        df = leer_resumen_nacional(archivo)
    else:
        st.error("No se pudo reconocer la estructura del archivo. Revisa que tenga encabezados de programa/meta/meses o código/delegación/meta/avance.")
        st.stop()

    if df.empty:
        st.warning("El archivo se leyó, pero no se encontraron registros útiles.")
        st.stop()

    st.success(f"Archivo leído correctamente: {tipo.replace('_', ' ').title()} | Registros detectados: {len(df):,}")

    # =====================================================
    # FILTROS
    # =====================================================
    st.subheader("🔎 Filtros")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        delegaciones = sorted([x for x in df["Delegación"].dropna().unique() if limpiar_texto(x)])
        filtro_delegacion = st.multiselect("Delegación", delegaciones)

    with col2:
        if "Programa" in df.columns:
            programas = sorted([x for x in df["Programa"].dropna().unique() if limpiar_texto(x)])
            filtro_programa = st.multiselect("Programa", programas)
        else:
            filtro_programa = []
            regiones = sorted([x for x in df["Región / Hoja"].dropna().unique() if limpiar_texto(x)]) if "Región / Hoja" in df.columns else []
            filtro_region = st.multiselect("Región", regiones)

    with col3:
        filtro_estado = st.multiselect("Estado", sorted(df["Estado"].dropna().unique()))

    with col4:
        texto_busqueda = st.text_input("Buscar texto", placeholder="Actividad, cantón, provincia...")

    df_filtrado = df.copy()

    if filtro_delegacion:
        df_filtrado = df_filtrado[df_filtrado["Delegación"].isin(filtro_delegacion)]

    if "Programa" in df_filtrado.columns and filtro_programa:
        df_filtrado = df_filtrado[df_filtrado["Programa"].isin(filtro_programa)]

    if "Región / Hoja" in df_filtrado.columns and 'filtro_region' in locals() and filtro_region:
        df_filtrado = df_filtrado[df_filtrado["Región / Hoja"].isin(filtro_region)]

    if filtro_estado:
        df_filtrado = df_filtrado[df_filtrado["Estado"].isin(filtro_estado)]

    if texto_busqueda:
        texto = texto_busqueda.upper().strip()
        mascara = df_filtrado.astype(str).apply(lambda col: col.str.upper().str.contains(texto, na=False)).any(axis=1)
        df_filtrado = df_filtrado[mascara]

    # Filtro por mes solo para archivo regional
    if tipo == "DETALLE_REGIONAL":
        st.markdown("#### 🗓️ Filtro / lectura por mes")
        meses_disponibles = [m.title() for m in MESES if m.title() in df_filtrado.columns]
        meses_sel = st.multiselect("Meses a revisar", meses_disponibles, default=meses_disponibles[:4] if len(meses_disponibles) >= 4 else meses_disponibles)
        if meses_sel:
            mostrar_solo_con_movimiento = st.checkbox("Mostrar solo actividades con movimiento en los meses seleccionados", value=False)
            if mostrar_solo_con_movimiento:
                df_filtrado = df_filtrado[df_filtrado[meses_sel].sum(axis=1) > 0]

    # =====================================================
    # MÉTRICAS
    # =====================================================
    st.subheader("📌 Resumen")
    total_meta = df_filtrado["Meta"].sum()
    total_avance = df_filtrado["Avance"].sum()
    avance_global = total_avance / total_meta if total_meta else 0
    total_pendiente = df_filtrado["Pendiente"].sum()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Registros", f"{len(df_filtrado):,}")
    m2.metric("Meta", f"{total_meta:,.0f}")
    m3.metric("Avance", f"{total_avance:,.0f}")
    m4.metric("% Avance", f"{avance_global:.2%}")

    st.metric("Pendiente", f"{total_pendiente:,.0f}")

    # =====================================================
    # TABLA Y GRÁFICAS
    # =====================================================
    st.subheader("📋 Datos filtrados")
    vista = df_filtrado.copy()
    if "% Avance" in vista.columns:
        vista["% Avance"] = vista["% Avance"].map(lambda x: f"{x:.2%}")
    st.dataframe(vista, use_container_width=True, hide_index=True)

    st.subheader("📊 Visualización rápida")
    c1, c2 = st.columns(2)
    with c1:
        avance_delegacion = df_filtrado.groupby("Delegación", as_index=False)[["Meta", "Avance"]].sum()
        if not avance_delegacion.empty:
            st.bar_chart(avance_delegacion.set_index("Delegación")[["Meta", "Avance"]])
    with c2:
        if "Programa" in df_filtrado.columns:
            avance_programa = df_filtrado.groupby("Programa", as_index=False)[["Meta", "Avance"]].sum()
            if not avance_programa.empty:
                st.bar_chart(avance_programa.set_index("Programa")[["Meta", "Avance"]])
        elif "Región / Hoja" in df_filtrado.columns:
            avance_region = df_filtrado.groupby("Región / Hoja", as_index=False)[["Meta", "Avance"]].sum()
            if not avance_region.empty:
                st.bar_chart(avance_region.set_index("Región / Hoja")[["Meta", "Avance"]])

    # =====================================================
    # DESCARGAS
    # =====================================================
    st.subheader("⬇️ Descargas")
    col_excel, col_pdf = st.columns(2)

    excel_bytes = generar_excel(df_filtrado, "Datos filtrados")
    pdf_bytes = generar_pdf(df_filtrado, "Reporte básico de datos filtrados")

    with col_excel:
        st.download_button(
            "📥 Descargar Excel filtrado",
            data=excel_bytes,
            file_name="datos_filtrados_avances.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_pdf:
        st.download_button(
            "📄 Descargar PDF básico",
            data=pdf_bytes,
            file_name="reporte_basico_avances.pdf",
            mime="application/pdf"
        )

except Exception as e:
    st.error("Ocurrió un error leyendo el archivo.")
    st.exception(e)

