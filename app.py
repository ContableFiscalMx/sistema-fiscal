import os
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st

# 1. LA CONFIGURACIÓN DE PÁGINA DEBE SER LO PRIMERO DE STREAMLIT
st.set_page_config(
    page_title="Sistema de Control Fiscal", page_icon="📊", layout="wide"
)

# 2. CSS COMPLETO PARA OCULTAR MENÚS, LOGOS Y BOTONES DE STREAMLIT
st.markdown(
    """
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .stAppDeployButton {display: none !important;}
        [data-testid="stStatusWidget"] {visibility: hidden !important;}
        [data-testid="stToolbar"] {visibility: hidden !important; display: none !important;}
        .block-container {
            padding-top: 1rem;
            padding-bottom: 1rem;
        }
    </style>
""",
    unsafe_allow_html=True,
)


# --- SISTEMA DE SEGURIDAD Y LOGIN ---
def check_password():
  """Valida la contraseña antes de mostrar el contenido."""

  def password_entered():
    if st.session_state["password"] == "adminfiscal27":
      st.session_state["password_correct"] = True
      del st.session_state["password"]
    else:
      st.session_state["password_correct"] = False

  if "password_correct" not in st.session_state:
    st.title("🔒 Acceso Restringido")
    st.text_input(
        "Ingresa tu contraseña para entrar al Sistema Fiscal:",
        type="password",
        on_change=password_entered,
        key="password",
    )
    return False
  elif not st.session_state["password_correct"]:
    st.title("🔒 Acceso Restringido")
    st.text_input(
        "Ingresa tu contraseña para entrar al Sistema Fiscal:",
        type="password",
        on_change=password_entered,
        key="password",
    )
    st.error("❌ Contraseña incorrecta.")
    return False

  return True


if not check_password():
  st.stop()
# --- FIN DEL SISTEMA DE SEGURIDAD ---


# Función de caché en el nivel superior para optimizar el rendimiento
@st.cache_data
def cargar_catalogo(path):
  df_bruto = pd.read_excel(path, header=None)
  fila_encabezado = 0
  for idx, row in df_bruto.iterrows():
    fila_texto = row.astype(str).str.cat(sep=" ").lower()
    if "claveprodserv" in fila_texto or "c_clave" in fila_texto:
      fila_encabezado = idx
      break
  df = pd.read_excel(path, skiprows=fila_encabezado, dtype=str)
  df.columns = df.columns.astype(str).str.strip()
  return df.fillna("")


# Título de la plataforma
st.title("🏢 Sistema de Control Fiscal")

# Menú de navegación horizontal
opcion_menu = st.radio(
    "Navegación:",
    [
        "📊 Panel General",
        "🧮 Calculadora de Impuestos",
        "📂 Extracción de XML (CFDI 4.0)",
        "📋 Catálogo y Claves SAT",
    ],
    horizontal=True,
    label_visibility="collapsed",
)
st.markdown("---")

# ---------------------------------------------------------
# APARTADO 1: PANEL GENERAL
# ---------------------------------------------------------
if opcion_menu == "📊 Panel General":
  st.title("📊 Resumen Operativo")

  col1, col2, col3 = st.columns(3)
  with col1:
    st.metric(label="Valor UMA 2025", value="$130.14", delta="Vigente")
  with col2:
    st.metric(label="Tasa Recargos", value="1.47%", delta="Mensual")
  with col3:
    st.metric(label="ISN Tamaulipas", value="3%", delta="Estatal")

  st.markdown("---")
  st.subheader("Directorio de Aclaraciones")
  st.info("✉️ : vigilancia.cumplimiento@tamaulipas.gob.mx")

# ---------------------------------------------------------
# APARTADO 2: CALCULADORA DE IMPUESTOS
# ---------------------------------------------------------
elif opcion_menu == "🧮 Calculadora de Impuestos":
  col_izq, col_centro, col_der = st.columns([1, 2, 1])

  with col_centro:
    st.title("🧮 Calculadora de Retenciones")
    st.caption(
        "Determina los importes directos e inversos para RESICO, Honorarios y"
        " Arrendamiento."
    )

    if "calc_cantidad" not in st.session_state:
      st.session_state["calc_cantidad"] = 0.0
    if "calc_tasa_iva" not in st.session_state:
      st.session_state["calc_tasa_iva"] = 16.0

    def limpiar_formulario():
      st.session_state["calc_cantidad"] = 0.0
      st.session_state["calc_tasa_iva"] = 16.0

    regimen = st.selectbox(
        "Seleccionar Régimen Fiscal:",
        [
            "RESICO",
            "Arrendamiento",
            "Servicios Profesionales (Honorarios)",
            "Actividad Empresarial (Comercial)",
        ],
    )

    if regimen == "RESICO":
      opciones_ret = {
          "Ret ISR 1.25% | Ret IVA 2/3 (Persona Moral)": (0.0125, 2 / 3, 16.0),
          "Solo Ret ISR 1.25% (IVA Tasa 0% / Exento)": (0.0125, 0.0, 0.0),
          "Solo Ret ISR 1.25% (Con IVA 16% sin Ret. IVA)": (0.0125, 0.0, 16.0),
          "Sin Retenciones (Público General / PF)": (0.0, 0.0, 16.0),
      }
    elif regimen in ["Arrendamiento", "Servicios Profesionales (Honorarios)"]:
      opciones_ret = {
          "Ret ISR 10% | Ret IVA 2/3 (Persona Moral)": (0.10, 2 / 3, 16.0),
          "Solo Ret ISR 10% (IVA Tasa 0% / Exento)": (0.10, 0.0, 0.0),
          "Solo Ret ISR 10% (Con IVA 16% sin Ret. IVA)": (0.10, 0.0, 16.0),
          "Sin Retenciones (Público General / PF)": (0.0, 0.0, 16.0),
      }
    else:
      opciones_ret = {
          "Operación General (IVA 16%)": (0.0, 0.0, 16.0),
          "Operación Tasa 0% / Exento": (0.0, 0.0, 0.0),
      }

    escenario = st.radio(
        "Configuración de Retenciones:", list(opciones_ret.keys())
    )

    tasa_isr_aplicable = opciones_ret[escenario][0]
    factor_ret_iva = opciones_ret[escenario][1]
    tasa_iva_defecto = opciones_ret[escenario][2]

    st.markdown("---")

    tipo_calculo = st.radio(
        "Seleccionar tipo de Importe:",
        ["Importe Bruto", "Importe Neto (Cálculo Inverso)"],
        horizontal=True,
    )

    col_a, col_b = st.columns(2)
    with col_a:
      es_tasa_cero = tasa_iva_defecto == 0.0
      tasa_iva_input = st.number_input(
          "% IVA*",
          min_value=0.0,
          max_value=100.0,
          value=0.0 if es_tasa_cero else st.session_state["calc_tasa_iva"],
          step=1.0,
          disabled=es_tasa_cero,
          key="calc_tasa_iva" if not es_tasa_cero else None,
      )
    with col_b:
      cantidad = st.number_input(
          "Cantidad ($)*", min_value=0.0, step=100.0, key="calc_cantidad"
      )

    tasa_iva_actual = 0.0 if es_tasa_cero else tasa_iva_input
    tasa_iva_dec = tasa_iva_actual / 100.0
    tasa_ret_iva_dec = tasa_iva_dec * factor_ret_iva

    if "Bruto" in tipo_calculo:
      subtotal = cantidad
    else:
      factor_inverso = (
          1.0 + tasa_iva_dec - tasa_ret_iva_dec - tasa_isr_aplicable
      )
      subtotal = cantidad / factor_inverso if factor_inverso != 0 else 0.0

    importe_iva = subtotal * tasa_iva_dec
    importe_ret_iva = subtotal * tasa_ret_iva_dec
    importe_ret_isr = subtotal * tasa_isr_aplicable
    importe_neto = (
        subtotal + importe_iva - importe_ret_iva - importe_ret_isr
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.text_input("Subtotal", value=f"$ {subtotal:,.2f}", disabled=True)
    st.text_input(
        "(+) Importe de IVA", value=f"$ {importe_iva:,.2f}", disabled=True
    )
    st.text_input(
        "(-) Importe de Retención de IVA",
        value=f"$ {importe_ret_iva:,.2f}",
        disabled=True,
    )

    etiqueta_isr = (
        f"(-) Importe de Retención de ISR ({tasa_isr_aplicable * 100:.2f}%)"
        if tasa_isr_aplicable > 0
        else "(-) Importe de Retención de ISR"
    )
    st.text_input(
        etiqueta_isr, value=f"$ {importe_ret_isr:,.2f}", disabled=True
    )

    st.text_input("Importe Neto", value=f"$ {importe_neto:,.2f}", disabled=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.button(
        "RESET",
        type="primary",
        on_click=limpiar_formulario,
        use_container_width=True,
    )

# ---------------------------------------------------------
# APARTADO 3: EXTRACCIÓN DE XML (CFDI 4.0)
# ---------------------------------------------------------
elif opcion_menu == "📂 Extracción de XML (CFDI 4.0)":
  st.title("📂 Procesador de Nodos XML")
  st.write("Sube los CFDI para extraer UUID, fechas de emisión y retenciones.")

  archivos_xml = st.file_uploader(
      "Seleccionar archivos XML", type=["xml"], accept_multiple_files=True
  )

  if archivos_xml:
    datos_list = []
    for archivo in archivos_xml:
      try:
        tree = ET.parse(archivo)
        root = tree.getroot()
        datos_list.append(
            {"Archivo": archivo.name, "Estado": "Leído correctamente"}
        )
      except Exception as e:
        datos_list.append({"Archivo": archivo.name, "Estado": f"Error: {e}"})

    st.dataframe(pd.DataFrame(datos_list), use_container_width=True)

# ---------------------------------------------------------
# APARTADO 4: CATÁLOGO Y CLAVES SAT
# ---------------------------------------------------------
elif opcion_menu == "📋 Catálogo y Claves SAT":
  st.title("📋 Consulta Rápida SAT")

  archivo_local = "catalogo.xlsx"

  if os.path.exists(archivo_local):
    busqueda = st.text_input("Buscar por clave, descripción o palabra similar:")

    df_catalogo = cargar_catalogo(archivo_local)

    col_clave = next(
        (col for col in df_catalogo.columns if "clave" in col.lower()),
        df_catalogo.columns[0],
    )
    col_desc = next(
        (col for col in df_catalogo.columns if "desc" in col.lower()),
        (
            df_catalogo.columns[1]
            if len(df_catalogo.columns) > 1
            else df_catalogo.columns[0]
        ),
    )
    col_similares = next(
        (
            col
            for col in df_catalogo.columns
            if "similar" in col.lower() or "palabra" in col.lower()
        ),
        None,
    )

    if busqueda:
      filtro = df_catalogo[col_clave].astype(str).str.contains(
          busqueda, case=False
      ) | df_catalogo[col_desc].astype(str).str.contains(busqueda, case=False)
      if col_similares:
        filtro = filtro | df_catalogo[col_similares].astype(str).str.contains(
            busqueda, case=False
        )

      resultados = df_catalogo[filtro]
      st.write(f"Resultados encontrados: {len(resultados)}")
      st.dataframe(resultados, use_container_width=True)
    else:
      st.write("Vista previa del catálogo (Primeros 50 registros):")
      st.dataframe(df_catalogo.head(50), use_container_width=True)
  else:
    st.error(
        "No se encontró el archivo 'catalogo.xlsx' en la carpeta del sistema."
    )
    st.info(
        "Coloca tu archivo de Excel del SAT directamente dentro del"
        " repositorio y nómbralo 'catalogo.xlsx'."
    )
