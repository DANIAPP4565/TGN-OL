# app.py
# ============================================================
# APP ABII / TGN - REPORTE OPI AUTORIZACIÓN
# Automatiza:
# 1) Login en https://abii.tgn.com.ar/
# 2) Reportes > OPI > Reporte OPI Autorización
# 3) Fecha desde ayer / fecha hasta hoy
# 4) Generar reporte
# 5) Exportar Excel
# ============================================================

import os
import platform
import traceback
from datetime import date, timedelta
from pathlib import Path
from typing import List, Optional, Tuple

import streamlit as st
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ============================================================
# AUTO-INSTALACIÓN / VERIFICACIÓN DE CHROMIUM PLAYWRIGHT
# Streamlit Cloud instala el paquete playwright, pero a veces no descarga
# el navegador Chromium. Este bloque lo descarga automáticamente si falta.
# ============================================================

def asegurar_chromium_playwright() -> str:
    import subprocess
    import sys
    from pathlib import Path

    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            ruta = Path(p.chromium.executable_path)
            if ruta.exists():
                return f"Chromium disponible: {ruta}"
    except Exception:
        pass

    try:
        resultado = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
        if resultado.returncode == 0:
            return "Chromium instalado correctamente por Playwright."
        return (
            "No se pudo instalar Chromium automáticamente.\n"
            f"STDOUT: {resultado.stdout[-1500:]}\n"
            f"STDERR: {resultado.stderr[-1500:]}"
        )
    except Exception as e:
        return f"No se pudo ejecutar playwright install chromium: {repr(e)}"


ESTADO_CHROMIUM = asegurar_chromium_playwright()


# ============================================================
# CONFIGURACIÓN
# ============================================================

URL_ABII = "https://abii.tgn.com.ar/"
# Carpeta donde se guardara automaticamente el Excel final
# En Windows normalmente sera: C:\Users\TuUsuario\Downloads
CARPETA_DESCARGAS = Path.home() / "Downloads"

# Carpeta local para capturas de control/debug
CARPETA_CAPTURAS = Path("capturas_opi")

CARPETA_DESCARGAS.mkdir(exist_ok=True)
CARPETA_CAPTURAS.mkdir(exist_ok=True)


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def fecha_ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def guardar_captura(page, nombre: str) -> Optional[Path]:
    try:
        ruta = CARPETA_CAPTURAS / nombre
        page.screenshot(path=str(ruta), full_page=True)
        return ruta
    except Exception:
        return None


def esperar(page, ms: int = 1000):
    try:
        page.wait_for_timeout(ms)
    except Exception:
        pass


def click_texto(page, textos: List[str], timeout: int = 5000) -> bool:
    ultimo_error = None
    for texto in textos:
        try:
            loc = page.get_by_text(texto, exact=False).first
            loc.wait_for(timeout=timeout)
            loc.click(timeout=timeout)
            return True
        except Exception as e:
            ultimo_error = e
    raise RuntimeError(f"No se encontró texto visible: {textos}. Detalle: {repr(ultimo_error)}")


def click_boton_robusto(page, textos: List[str], timeout: int = 4000) -> bool:
    """Busca botón por rol, texto, input submit, iconos o último botón visible."""
    # 1) Por role button
    for texto in textos:
        try:
            btn = page.get_by_role("button", name=texto, exact=False).first
            btn.wait_for(timeout=timeout)
            btn.click(timeout=timeout)
            return True
        except Exception:
            pass

    # 2) Por texto visible
    for texto in textos:
        try:
            loc = page.get_by_text(texto, exact=False).first
            loc.wait_for(timeout=timeout)
            loc.click(timeout=timeout)
            return True
        except Exception:
            pass

    # 3) Inputs submit/button con value
    for texto in textos:
        try:
            loc = page.locator(f"input[type='submit'][value*='{texto}' i]").first
            loc.wait_for(timeout=timeout)
            loc.click(timeout=timeout)
            return True
        except Exception:
            pass
        try:
            loc = page.locator(f"input[type='button'][value*='{texto}' i]").first
            loc.wait_for(timeout=timeout)
            loc.click(timeout=timeout)
            return True
        except Exception:
            pass

    # 4) Botones comunes por clase/título
    selectores = [
        "button[type='submit']",
        "input[type='submit']",
        "button:has-text('Excel')",
        "a:has-text('Excel')",
        "button:has-text('Export')",
        "a:has-text('Export')",
        "button:has-text('Buscar')",
        "button:has-text('Generar')",
        "button:has-text('Consultar')",
        "a:has-text('Generar')",
        "a:has-text('Buscar')",
        "a:has-text('Consultar')",
    ]
    for sel in selectores:
        try:
            loc = page.locator(sel).first
            loc.wait_for(timeout=timeout)
            loc.click(timeout=timeout)
            return True
        except Exception:
            pass

    # 5) Último botón visible como último recurso
    try:
        botones = page.locator("button:visible")
        count = botones.count()
        if count > 0:
            botones.nth(count - 1).click(timeout=timeout)
            return True
    except Exception:
        pass

    # 6) Enter como último recurso
    try:
        page.keyboard.press("Enter")
        return True
    except Exception:
        pass

    return False


def completar_campo_robusto(page, valor: str, selectores: List[str], timeout: int = 2500) -> bool:
    """Completa un campo buscando en página principal y también dentro de iframes."""
    ultimo_error = None
    contextos = [page] + list(page.frames)

    for ctx in contextos:
        for selector in selectores:
            try:
                if selector.startswith("label="):
                    loc = ctx.get_by_label(selector.replace("label=", ""), exact=False).first
                elif selector.startswith("placeholder="):
                    loc = ctx.get_by_placeholder(selector.replace("placeholder=", ""), exact=False).first
                else:
                    loc = ctx.locator(selector).first

                loc.wait_for(timeout=timeout)
                loc.click(timeout=timeout)
                loc.fill("")
                loc.fill(valor)
                return True
            except Exception as e:
                ultimo_error = e

    raise RuntimeError(f"No se pudo completar campo con valor '{valor}'. Detalle: {repr(ultimo_error)}")


def completar_login_directo(page, usuario: str, clave: str) -> bool:
    """Login robusto: detecta usuario y contraseña aunque cambien IDs/nombres/labels."""
    contextos = [page] + list(page.frames)

    for ctx in contextos:
        try:
            # Buscar campo de contraseña primero
            pass_input = ctx.locator("input[type='password']").first
            pass_input.wait_for(timeout=5000)

            # Buscar input de usuario cercano o primer input visible que no sea password/hidden
            user_candidates = ctx.locator(
                "input:visible:not([type='password']):not([type='hidden']):not([type='submit']):not([type='button'])"
            )
            total = user_candidates.count()
            if total == 0:
                continue

            user_input = user_candidates.first

            user_input.click(timeout=3000)
            user_input.fill("")
            user_input.fill(usuario)

            pass_input.click(timeout=3000)
            pass_input.fill("")
            pass_input.fill(clave)

            return True
        except Exception:
            continue

    return False


def completar_fecha_robusto(page, valor: str, selectores: List[str], timeout: int = 5000) -> bool:
    ultimo_error = None
    for selector in selectores:
        try:
            if selector.startswith("label="):
                loc = page.get_by_label(selector.replace("label=", ""), exact=False).first
            elif selector.startswith("placeholder="):
                loc = page.get_by_placeholder(selector.replace("placeholder=", ""), exact=False).first
            else:
                loc = page.locator(selector).first

            loc.wait_for(timeout=timeout)
            loc.click(timeout=timeout)
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            loc.type(valor, delay=35)
            page.keyboard.press("Tab")
            return True
        except Exception as e:
            ultimo_error = e

    raise RuntimeError(f"No se pudo completar fecha '{valor}'. Detalle: {repr(ultimo_error)}")


def diagnosticar_pantalla(page) -> str:
    """Devuelve textos visibles acotados para saber dónde quedó la app."""
    try:
        txt = page.locator("body").inner_text(timeout=3000)
        txt = txt.strip()
        if len(txt) > 2500:
            txt = txt[:2500] + "..."
        return txt
    except Exception as e:
        return f"No se pudo leer pantalla: {repr(e)}"


# ============================================================
# AUTOMATIZACIÓN PRINCIPAL
# ============================================================

def descargar_reporte_opi(
    usuario: str,
    clave: str,
    fecha_desde: date,
    fecha_hasta: date,
    navegador_oculto: bool = False,
) -> Tuple[Optional[Path], List[Path], str]:

    capturas: List[Path] = []
    archivo_final: Optional[Path] = None

    with sync_playwright() as p:
        # IMPORTANTE:
        # Streamlit Cloud/Linux no tiene pantalla gráfica (XServer).
        # Por eso Chromium DEBE ejecutarse en modo headless=True.
        # En Windows local se permite ver el navegador si el usuario desmarca el checkbox.
        es_windows = platform.system().lower().startswith("win")
        headless_final = True if not es_windows else navegador_oculto

        browser = p.chromium.launch(
            headless=headless_final,
            slow_mo=250 if (es_windows and not headless_final) else 0,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--disable-setuid-sandbox",
                "--disable-software-rasterizer",
            ],
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1400, "height": 900},
        )
        page = context.new_page()
        page.set_default_timeout(20000)

        try:
            # 1) Abrir sitio
            page.goto(URL_ABII, wait_until="domcontentloaded")
            esperar(page, 1500)
            cap = guardar_captura(page, "01_login.png")
            if cap:
                capturas.append(cap)

            # 2) Login robusto: usuario + clave
            login_ok = completar_login_directo(page, usuario, clave)

            if not login_ok:
                # Método alternativo por selectores tradicionales
                completar_campo_robusto(
                    page,
                    usuario,
                    [
                        "input[name='usuario']",
                        "input[name='user']",
                        "input[name='username']",
                        "input[id*='usuario' i]",
                        "input[id*='user' i]",
                        "input[type='text']",
                        "placeholder=Usuario",
                        "placeholder=User",
                        "label=Usuario",
                    ],
                )

                completar_campo_robusto(
                    page,
                    clave,
                    [
                        "input[type='password']",
                        "input[name='password']",
                        "input[name='clave']",
                        "input[id*='password' i]",
                        "input[id*='clave' i]",
                        "placeholder=Contraseña",
                        "placeholder=Clave",
                        "placeholder=Password",
                        "label=Contraseña",
                        "label=Clave",
                    ],
                )

            # 4) Login
            ok_login = click_boton_robusto(page, ["Ingresar", "Entrar", "Acceder", "Login", "Aceptar"], timeout=4000)
            if not ok_login:
                page.keyboard.press("Enter")

            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            esperar(page, 2500)

            cap = guardar_captura(page, "02_post_login.png")
            if cap:
                capturas.append(cap)

            texto_post_login = diagnosticar_pantalla(page)
            if "usuario" in texto_post_login.lower() and "contraseña" in texto_post_login.lower():
                raise RuntimeError("Parece que el login no avanzó. Verificar usuario, contraseña, captcha o permisos.")

            # 5) Reportes
            click_texto(page, ["Reportes", "REPORTES", "reportes"], timeout=10000)
            esperar(page, 1200)
            cap = guardar_captura(page, "03_menu_reportes.png")
            if cap:
                capturas.append(cap)

            # 6) OPI
            click_texto(page, ["OPI", "Opi", "opi"], timeout=10000)
            esperar(page, 1200)
            cap = guardar_captura(page, "04_menu_opi.png")
            if cap:
                capturas.append(cap)

            # 7) Reporte OPI Autorización
            click_texto(
                page,
                [
                    "Reporte OPI Autorización",
                    "Reporte OPI Autorizacion",
                    "OPI Autorización",
                    "OPI Autorizacion",
                    "Autorización",
                    "Autorizacion",
                ],
                timeout=12000,
            )

            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            esperar(page, 2000)

            cap = guardar_captura(page, "05_pantalla_reporte.png")
            if cap:
                capturas.append(cap)

            # 8) Completar fechas
            desde_txt = fecha_ddmmyyyy(fecha_desde)
            hasta_txt = fecha_ddmmyyyy(fecha_hasta)

            completar_fecha_robusto(
                page,
                desde_txt,
                [
                    "input[name='fechaDesde']",
                    "input[name='desde']",
                    "input[id*='fechaDesde' i]",
                    "input[id*='desde' i]",
                    "input[placeholder*='Desde' i]",
                    "input[aria-label*='Desde' i]",
                    "placeholder=Desde",
                    "placeholder=Fecha desde",
                    "label=Desde",
                    "label=Fecha desde",
                ],
            )

            completar_fecha_robusto(
                page,
                hasta_txt,
                [
                    "input[name='fechaHasta']",
                    "input[name='hasta']",
                    "input[id*='fechaHasta' i]",
                    "input[id*='hasta' i]",
                    "input[placeholder*='Hasta' i]",
                    "input[aria-label*='Hasta' i]",
                    "placeholder=Hasta",
                    "placeholder=Fecha hasta",
                    "label=Hasta",
                    "label=Fecha hasta",
                ],
            )

            esperar(page, 800)
            cap = guardar_captura(page, "06_fechas_cargadas.png")
            if cap:
                capturas.append(cap)

            # 9) Generar reporte - versión robusta
            generado = click_boton_robusto(
                page,
                [
                    "Generar",
                    "Buscar",
                    "Consultar",
                    "Ver reporte",
                    "Aceptar",
                    "Filtrar",
                    "Ejecutar",
                    "Mostrar",
                    "Refrescar",
                ],
                timeout=4000,
            )

            if not generado:
                cap = guardar_captura(page, "error_no_boton_generar.png")
                if cap:
                    capturas.append(cap)
                pantalla = diagnosticar_pantalla(page)
                raise RuntimeError(
                    "No se pudo encontrar el botón para generar el reporte.\n\n"
                    f"Texto visible en pantalla:\n{pantalla}"
                )

            try:
                page.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass
            esperar(page, 4000)

            cap = guardar_captura(page, "07_reporte_generado.png")
            if cap:
                capturas.append(cap)

            # 10) Exportar Excel - versión robusta
            exportado = False
            error_export = None

            # Primero intentar con expect_download
            try:
                with page.expect_download(timeout=45000) as download_info:
                    ok_export = click_boton_robusto(
                        page,
                        [
                            "Exportar Excel",
                            "Exportar a Excel",
                            "Excel",
                            "Descargar Excel",
                            "Exportar",
                            "XLS",
                            "XLSX",
                            "CSV",
                        ],
                        timeout=5000,
                    )
                    if not ok_export:
                        raise RuntimeError("No se encontró botón de exportación.")

                descarga = download_info.value
                nombre_archivo = descarga.suggested_filename or f"reporte_opi_{fecha_desde}_{fecha_hasta}.xlsx"
                if not nombre_archivo.lower().endswith((".xlsx", ".xls", ".csv")):
                    nombre_archivo = f"reporte_opi_{fecha_desde}_{fecha_hasta}.xlsx"

                archivo_final = CARPETA_DESCARGAS / nombre_archivo
                descarga.save_as(str(archivo_final))
                exportado = True

            except Exception as e:
                error_export = e

            # Segundo intento: links directos con href a excel/xls/csv
            if not exportado:
                try:
                    links = page.locator("a:visible")
                    total = links.count()
                    for i in range(total):
                        a = links.nth(i)
                        txt = ""
                        href = ""
                        try:
                            txt = a.inner_text(timeout=1000).lower()
                            href = a.get_attribute("href") or ""
                        except Exception:
                            pass

                        if any(x in txt for x in ["excel", "xls", "xlsx", "csv", "exportar"]) or any(x in href.lower() for x in ["excel", "xls", "xlsx", "csv", "export"]):
                            with page.expect_download(timeout=45000) as download_info:
                                a.click(timeout=5000)
                            descarga = download_info.value
                            nombre_archivo = descarga.suggested_filename or f"reporte_opi_{fecha_desde}_{fecha_hasta}.xlsx"
                            archivo_final = CARPETA_DESCARGAS / nombre_archivo
                            descarga.save_as(str(archivo_final))
                            exportado = True
                            break
                except Exception as e:
                    error_export = e

            if not exportado:
                cap = guardar_captura(page, "error_no_export_excel.png")
                if cap:
                    capturas.append(cap)
                pantalla = diagnosticar_pantalla(page)
                raise RuntimeError(
                    "El reporte parece haberse generado, pero no se pudo exportar Excel.\n\n"
                    f"Último error de exportación: {repr(error_export)}\n\n"
                    f"Texto visible en pantalla:\n{pantalla}"
                )

            cap = guardar_captura(page, "08_excel_descargado.png")
            if cap:
                capturas.append(cap)

            return archivo_final, capturas, "Reporte generado y Excel descargado correctamente."

        except PlaywrightTimeoutError as e:
            cap = guardar_captura(page, "error_timeout.png")
            if cap:
                capturas.append(cap)
            raise RuntimeError(f"Tiempo de espera agotado. Detalle: {repr(e)}")

        except Exception as e:
            cap = guardar_captura(page, "error_general.png")
            if cap:
                capturas.append(cap)
            raise RuntimeError(str(e))

        finally:
            try:
                context.close()
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.set_page_config(
    page_title="ABII OPI",
    page_icon="📄",
    layout="centered",
)

st.markdown(
    """
    <style>
    .stButton > button {
        width: 100%;
        background-color: #0b5ed7;
        color: white;
        font-weight: 700;
        border-radius: 10px;
        padding: 0.8rem;
        border: 2px solid #0b5ed7;
    }
    .stDownloadButton > button {
        width: 100%;
        background-color: #198754;
        color: white;
        font-weight: 700;
        border-radius: 10px;
        padding: 0.8rem;
        border: 2px solid #198754;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📄 APP ABII / TGN")
st.subheader("Descarga automática de Reporte OPI Autorización")

st.info(
    "Flujo automático: Login → Reportes → OPI → Reporte OPI Autorización → fechas → generar → exportar Excel."
)

with st.expander("🧩 Estado Playwright / Chromium", expanded=False):
    st.code(ESTADO_CHROMIUM, language="text")
    st.caption("En Streamlit Cloud el navegador debe ejecutarse oculto/headless.")

with st.expander("🔐 Acceso", expanded=True):
    usuario = st.text_input("Usuario", placeholder="Ingrese usuario")
    clave = st.text_input("Contraseña", type="password", placeholder="Ingrese contraseña")

hoy = date.today()
ayer = hoy - timedelta(days=1)

with st.expander("📅 Fechas", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        fecha_desde = st.date_input("Desde", value=ayer, format="DD/MM/YYYY")
    with col2:
        fecha_hasta = st.date_input("Hasta", value=hoy, format="DD/MM/YYYY")

with st.expander("🧪 Opciones de prueba", expanded=True):
    es_windows_ui = platform.system().lower().startswith("win")
    navegador_oculto = st.checkbox(
        "Ejecutar navegador oculto",
        value=True if not es_windows_ui else False,
        disabled=not es_windows_ui,
        help=(
            "En Streamlit Cloud/Linux debe ejecutarse oculto porque no hay pantalla gráfica/XServer. "
            "En Windows local puede desmarcarlo para ver Chrome."
        ),
    )

st.divider()

if st.button("🚀 Generar reporte y descargar Excel"):
    if not usuario.strip():
        st.error("Falta ingresar usuario.")
    elif not clave.strip():
        st.error("Falta ingresar contraseña.")
    elif fecha_desde > fecha_hasta:
        st.error("La fecha desde no puede ser mayor que la fecha hasta.")
    else:
        with st.spinner("Ejecutando automatización. No cierre la ventana de Chrome si aparece..."):
            try:
                archivo, capturas, mensaje = descargar_reporte_opi(
                    usuario=usuario.strip(),
                    clave=clave.strip(),
                    fecha_desde=fecha_desde,
                    fecha_hasta=fecha_hasta,
                    navegador_oculto=navegador_oculto,
                )

                st.success(mensaje)

                if archivo and archivo.exists():
                    with open(archivo, "rb") as f:
                        st.download_button(
                            "⬇️ Descargar Excel",
                            data=f,
                            file_name=archivo.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

                if capturas:
                    with st.expander("🖼️ Capturas de control", expanded=False):
                        for cap in capturas:
                            if cap.exists():
                                st.image(str(cap), caption=cap.name, use_container_width=True)

            except Exception as e:
                st.error("ERROR REAL:")
                st.code(str(e), language="text")
                st.warning("No cierre el Chrome automático durante la ejecución. Si el menú o botón tiene otro nombre, mandar captura de la pantalla donde se detiene.")

                # Mostrar traceback para diagnóstico
                with st.expander("Detalle técnico", expanded=False):
                    st.code(traceback.format_exc(), language="text")

                # Mostrar últimas capturas
                caps = sorted(CARPETA_CAPTURAS.glob("*.png"), key=os.path.getmtime, reverse=True)
                if caps:
                    with st.expander("Últimas capturas", expanded=True):
                        for cap in caps[:8]:
                            st.image(str(cap), caption=cap.name, use_container_width=True)

st.divider()

st.markdown("### Comando correcto para correr")
st.code('py -3.12 -m streamlit run app.py', language="bat")

st.markdown("### Instalación si falta algo en Windows local")
st.code(
    """
py -3.12 -m pip install --upgrade pip
py -3.12 -m pip install streamlit playwright pandas openpyxl python-dotenv
py -3.12 -m playwright install chromium
py -3.12 -m streamlit run app.py
    """,
    language="bat",
)

st.markdown("### requirements.txt para Streamlit Cloud")
st.code(
    """
streamlit
playwright
pandas
openpyxl
python-dotenv
    """,
    language="text",
)

st.markdown("### packages.txt para Streamlit Cloud")
st.code(
    """
libcups2
libnss3
libatk-bridge2.0-0
libgbm1
libxkbcommon0
libxcomposite1
libxdamage1
libxfixes3
libxrandr2
libasound2
libpangocairo-1.0-0
libpango-1.0-0
    """,
    language="text",
)
