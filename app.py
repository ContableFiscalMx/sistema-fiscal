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


# Función de caché para el catálogo general del SAT
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


# Función de caché para leer las tablas de ISR desde las pestañas de Excel
@st.cache_data
def cargar_tabla_isr(path_excel, periodo_seleccionado):
  try:
    df = pd.read_excel(path_excel, sheet_name=periodo_seleccionado)
    df.columns = df.columns.str.strip()

    tabla_tuplas = []
    for _, row in df.iterrows():
      lim_inf = float(row["Límite inferior"])

      lim_sup_val = str(row["Límite superior"]).strip().lower()
      if (
          pd.isna(row["Límite superior"])
          or "inf" in lim_sup_val
          or "adelante" in lim_sup_val
          or lim_sup_val == ""
      ):
        lim_sup = float("inf")
      else:
        lim_sup = float(row["Límite superior"])

      cf_val = str(row["Cuota fija"]).replace("$", "").strip()
      cuota_fija = 0.0 if cf_val in ["-", "", "nan", "None"] else float(cf_val)

      pct_val = str(row["% sobre excedente"]).replace("%", "").strip()
      porcentaje_raw = float(pct_val)

      # Si Excel lo guardó como fracción decimal (menor a 1.0), lo convertimos a porcentaje real
      if porcentaje_raw < 1.0:
        porcentaje = porcentaje_raw * 100.0
      else:
        porcentaje = porcentaje_raw

      tabla_tuplas.append((lim_inf, lim_sup, cuota_fija, porcentaje))
    return tabla_tuplas
  except Exception as e:
    return []


# Función auxiliar para calcular el ISR mediante tarifa progresiva
def calcular_isr_tarifa(ingreso, tabla):
    for lim_inf, lim_sup, cuota_fija, porcentaje in tabla:
        if lim_inf <= ingreso <= lim_sup:
            excedente = ingreso - lim_inf
            impuesto_marginal = excedente * (porcentaje / 100.0)
            isr_total = impuesto_marginal + cuota_fija
            return {
                "lim_inf": lim_inf,
                "lim_sup": lim_sup,
                "excedente": excedente,
                "porcentaje": porcentaje,
                "impuesto_marginal": impuesto_marginal,
                "cuota_fija": cuota_fija,
                "isr_total": isr_total,
            }
    if tabla:
        lim_inf, lim_sup, cuota_fija, porcentaje = tabla[-1]
        if ingreso > lim_inf:
            excedente = ingreso - lim_inf
            impuesto_marginal = excedente * (porcentaje / 100.0)
            isr_total = impuesto_marginal + cuota_fija
            return {
                "lim_inf": lim_inf,
                "lim_sup": lim_sup,
                "excedente": excedente,
                "porcentaje": porcentaje,
                "impuesto_marginal": impuesto_marginal,
                "cuota_fija": cuota_fija,
                "isr_total": isr_total,
            }
    return None


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
        st.metric(label="Valor UMA 2026 (Diario)", value="$117.31", delta="Vigente")
    with col2:
        st.metric(label="Tasa Recargos", value="2.07%", delta="Mensual")
    with col3:
        st.metric(label="ISN Tamaulipas", value="3%", delta="Estatal")

    st.markdown("---")

    # --- TABLA DE INGRESOS EXENTOS (ART. 93 LISR) ---
    with st.expander("📌 Ver Tabla de Ingresos Exentos para Trabajadores (Art. 93 LISR 2026)", expanded=True):
        st.caption("Cálculos elaborados con base en el valor UMA diario vigente ($117.31 MXN).")
        
        datos_exentos = [
            {"Concepto": "Aguinaldo", "Exención Ley": "30 UMA", "Fundamento": "Art. 93 Fracc. XIV", "Importe Máx. Exento": "$3,519.30"},
            {"Concepto": "Prima Vacacional", "Exención Ley": "15 UMA", "Fundamento": "Art. 93 Fracc. XIV", "Importe Máx. Exento": "$1,759.65"},
            {"Concepto": "PTU (Reparto de Utilidades)", "Exención Ley": "15 UMA", "Fundamento": "Art. 93 Fracc. XIV", "Importe Máx. Exento": "$1,759.65"},
            {"Concepto": "Prima Dominical", "Exención Ley": "1 UMA por domingo", "Fundamento": "Art. 93 Fracc. XIV", "Importe Máx. Exento": "$117.31 / domingo"},
            {"Concepto": "Jubilaciones / Pensiones", "Exención Ley": "15 UMA diarias", "Fundamento": "Art. 93 Fracc. IV", "Importe Máx. Exento": "$1,759.65 / día"},
            {"Concepto": "Indemnizaciones / Separación", "Exención Ley": "90 UMA por año trabajado", "Fundamento": "Art. 93 Fracc. XIII", "Importe Máx. Exento": "$10,557.90 / año"},
        ]
        
        st.dataframe(pd.DataFrame(datos_exentos), use_container_width=False, hide_index=True)
        
        st.markdown("""
        **Notas sobre Tiempo Extra y Días de Descanso (Art. 93 Fracc. I):**
        * **Trabajadores con Salario Mínimo General (SMG):** 100% exento (sin rebasar los límites de la LFT).
        * **Trabajadores con salario superior al mínimo:** 50% exento, siempre que la exención no rebase de **5 UMA semanales ($586.55 MXN)**.
        """)

    st.markdown("---")
    st.subheader("Directorio de Aclaraciones")
    st.info("✉️ : contablefiscalmx27@gmail.com")


# ---------------------------------------------------------
# APARTADO 2: CALCULADORA DE IMPUESTOS
# ---------------------------------------------------------
elif opcion_menu == "🧮 Calculadora de Impuestos":
    col_izq, col_centro, col_der = st.columns([1, 2, 1])

    with col_centro:
        st.markdown(
    """
    <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 1rem;">
        <span style="font-size: 2.5rem;">🧮</span>
        <h1 style="margin: 0; font-size: 2.1rem; font-weight: 700; color: #ffffff; line-height: 1.2;">Calculadora de Retenciones e ISR</h1>
    </div>
    """,
    unsafe_allow_html=True,
)
        st.caption(
            "Determina retenciones para RESICO, Arrendamiento, Honorarios, "
            "Actividad Empresarial y cálculo de ISR por Tarifas (Sueldos y Salarios)."
        )

        if "calc_cantidad" not in st.session_state:
            st.session_state["calc_cantidad"] = 0.0
        if "calc_tasa_iva" not in st.session_state:
            st.session_state["calc_tasa_iva"] = 16.0

        def limpiar_formulario():
            st.session_state["calc_cantidad"] = 0.0
            st.session_state["calc_tasa_iva"] = 16.0

        regimen = st.selectbox(
            "Seleccionar Régimen Fiscal / Tipo de Cálculo:",
            [
                "RESICO",
                "Arrendamiento",
                "Servicios Profesionales (Honorarios)",
                "Actividad Empresarial (Comercial)",
                "Sueldos y Salarios (ISR por Tarifas)",
            ],
        )

       # --- LÓGICA PARA SUELDOS Y SALARIOS (USANDO EXCEL) ---
        if regimen == "Sueldos y Salarios (ISR por Tarifas)":
          st.markdown("---")
          col_p1, col_p2 = st.columns(2)
          with col_p1:
            periodo_isr = st.selectbox(
                "Seleccionar Periodo de la Tarifa:",
                [
                    "Diaria",
                    "Semanal",
                    "Decenal",
                    "Quincenal",
                    "Mensual",
                    "Bimestral",
                    "Anual",
                ],
            )
          with col_p2:
            ingreso_gravado = st.number_input(
                "Ingreso Gravado del Periodo ($)*",
                min_value=0.0,
                step=100.0,
                key="calc_cantidad",
            )

          archivo_tabla_isr = "tabla_isr.xlsx"
          if os.path.exists(archivo_tabla_isr):
            tabla_cargada = cargar_tabla_isr(archivo_tabla_isr, periodo_isr)

            if tabla_cargada:
              if ingreso_gravado > 0:
                resultado_isr = calcular_isr_tarifa(
                    ingreso_gravado, tabla_cargada
                )
                if resultado_isr:
                  st.markdown("### 📊 Resultado del Cálculo ISR")

                  # Usamos una cuadrícula de 2 columnas para los resultados para eliminar el scroll excesivo
                  res_col1, res_col2 = st.columns(2)

                  with res_col1:
                    st.text_input(
                        "Límite Inferior",
                        value=f"$ {resultado_isr['lim_inf']:,.2f}",
                        disabled=True,
                    )
                    st.text_input(
                        "Excedente Límite Inf.",
                        value=f"$ {resultado_isr['excedente']:,.2f}",
                        disabled=True,
                    )
                    st.text_input(
                        "Impuesto Marginal",
                        value=f"$ {resultado_isr['impuesto_marginal']:,.2f}",
                        disabled=True,
                    )
                    st.text_input(
                        "Cuota Fija",
                        value=f"$ {resultado_isr['cuota_fija']:,.2f}",
                        disabled=True,
                    )

                  with res_col2:
                    lim_sup_txt = (
                        "En adelante (inf)"
                        if resultado_isr["lim_sup"] == float("inf")
                        else f"$ {resultado_isr['lim_sup']:,.2f}"
                    )
                    st.text_input(
                        "Límite Superior", value=lim_sup_txt, disabled=True
                    )
                    st.text_input(
                        "% sobre Excedente",
                        value=f"{resultado_isr['porcentaje']:.2f}%",
                        disabled=True,
                    )
                    # Destacamos el ISR Determinado con métrica visual limpia
                    st.metric(
                        label="ISF / ISR Determinado",
                        value=f"$ {resultado_isr['isr_total']:,.2f}",
                    )
            else:
              st.warning(
                  f"⚠️ La pestaña '{periodo_isr}' en '{archivo_tabla_isr}' está"
                  " vacía o tiene un formato incorrecto."
              )
          else:
            st.error(
                f"❌ No se encontró el archivo '{archivo_tabla_isr}' en la"
                " carpeta del sistema."
            )

        # --- LÓGICA PARA REGÍMENES COMERCIALES / PROFESIONALES ---
        else:
            if regimen == "RESICO":
                opciones_ret = {
                    "Retención ISR 1.25%, Retención IVA 2/3 partes": (0.0125, 2 / 3, 16.0, True, True),
                    "Retención ISR 1.25%, Sin retención de IVA": (0.0125, 0.0, 16.0, True, False),
                    "Retención ISR 1.25%": (0.0125, 0.0, 0.0, False, False),
                }
            elif regimen in ["Arrendamiento", "Servicios Profesionales (Honorarios)"]:
                opciones_ret = {
                    "Ret ISR 10% | Ret IVA 2/3 (Persona Moral)": (0.10, 2 / 3, 16.0, True, True),
                    "Solo Ret ISR 10% (IVA Tasa 0% / Exento)": (0.10, 0.0, 0.0, False, False),
                    "Solo Ret ISR 10% (Con IVA 16% sin Ret. IVA)": (0.10, 0.0, 16.0, True, False),
                }
            else:
                opciones_ret = {
                    "Operación General (IVA 16%)": (0.0, 0.0, 16.0, True, False),
                }

            escenario = st.radio(
                "Configuración de Retenciones:", list(opciones_ret.keys())
            )

            (tasa_isr_aplicable, 
             factor_ret_iva, 
             tasa_iva_defecto, 
             mostrar_iva, 
             mostrar_ret_iva) = opciones_ret[escenario]

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
            
            if mostrar_iva:
                st.text_input(
                    "(+) Importe de IVA", value=f"$ {importe_iva:,.2f}", disabled=True
                )
                
            if mostrar_ret_iva:
                st.text_input(
                    "(-) Importe de Retención de IVA",
                    value=f"$ {importe_ret_iva:,.2f}",
                    disabled=True,
                )

            if tasa_isr_aplicable > 0:
                st.text_input(
                    f"(-) Importe de Retención de ISR ({tasa_isr_aplicable * 100:.2f}%)", 
                    value=f"$ {importe_ret_isr:,.2f}", 
                    disabled=True
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
