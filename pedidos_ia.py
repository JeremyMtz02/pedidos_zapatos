import io
import os
from datetime import datetime
from github import Github
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="PEDIDOS DE ZAPATOS MINGA INC", page_icon="👠", layout="wide"
)

# ----------------------------------------------------
# 1. CONFIGURACIÓN DE GITHUB
# ----------------------------------------------------
# Recuerda configurar tu token en los Secrets de Streamlit o en .streamlit/secrets.toml
GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]

# REEMPLAZA ESTO con tu usuario y nombre del repositorio (ej. "usuario/mi-repo")
REPO_NAME = "JeremyMtz02/pedidos_zapatos"

# Nombre del archivo Excel que se guardará en GitHub
FILE_PATH = "PEDIDOS_PRUEBA.xlsx"

g = Github(GITHUB_TOKEN)
repo = g.get_repo(REPO_NAME)


def cargar_datos_github():
    try:
        content = repo.get_contents(FILE_PATH)
        data = content.decoded_content
        # Leemos el archivo excel desde los bytes de GitHub
        df = pd.read_excel(io.BytesIO(data))
        return df, content.sha
    except Exception:
        # Si el archivo aún no existe en el repositorio, creamos un DataFrame base
        df_nuevo = pd.DataFrame(
            columns=[
                "Pagina",
                "Zapato",
                "Cliente",
                "Total",
                "Abonado",
                "Restante",
            ]
        )
        return df_nuevo, None


def guardar_datos_github(df, sha_actual, mensaje_commit):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    content_bytes = output.getvalue()

    if sha_actual:
        # Actualiza el archivo si ya existe
        repo.update_file(FILE_PATH, mensaje_commit, content_bytes, sha_actual)
    else:
        # Crea el archivo por primera vez
        repo.create_file(FILE_PATH, mensaje_commit, content_bytes)


# ----------------------------------------------------
# 2. MANEJO DE CLIENTES LOCALES
# ----------------------------------------------------
ARCHIVO_CLIENTES = "clientes_lista.txt"
CLIENTES_INICIALES = [
    "Jimena Lopez",
    "Mary Hernandez",
    "Rita Martinez",
    "Consuelo Chavez",
    "Hortencia Flores",
]


def cargar_clientes():
    if os.path.exists(ARCHIVO_CLIENTES):
        with open(ARCHIVO_CLIENTES, "r", encoding="utf-8") as f:
            clientes = [line.strip() for line in f.readlines() if line.strip()]
            return clientes if clientes else CLIENTES_INICIALES
    else:
        with open(ARCHIVO_CLIENTES, "w", encoding="utf-8") as f:
            for c in CLIENTES_INICIALES:
                f.write(f"{c}\n")
        return CLIENTES_INICIALES


if "clientes_lista" not in st.session_state:
    st.session_state["clientes_lista"] = cargar_clientes()

lista_pagina = [x for x in range(1, 10)]
zapatos_dicc = {
    1: ["A1", "A2", "A3"],
    2: ["B1", "B2", "B3"],
    3: ["C1", "C2", "C3"],
    4: ["D1", "D2", "D3"],
    5: ["E1", "E2", "E3"],
    6: ["F1", "F2", "F3"],
    7: ["G1", "G2", "G3"],
    8: ["H1", "H2", "H3"],
    9: ["I1", "I2", "I3"],
    10: ["J1", "J2", "J3"],
}

st.title("👠 Pedidos de Zapatos Minga Inc")

# Cargar los datos actuales desde GitHub al iniciar la app
df_actual, sha_archivo = cargar_datos_github()

# ----------------------------------------------------
# 3. CAPTURA DEL PEDIDO
# ----------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    sel_pag = st.selectbox("Seleccione una página de la lista:", lista_pagina)
    zapatos_opciones = zapatos_dicc.get(sel_pag, ["A1", "A2", "A3"])
    sel_zapato = st.selectbox("Seleccione el par a elegir:", zapatos_opciones)

    # Selección de Cliente
    opcion_nuevo = "➕ Agregar nuevo cliente..."
    opciones_desplegable = st.session_state["clientes_lista"] + [opcion_nuevo]
    sel_cliente = st.selectbox("Seleccione una clienta:", opciones_desplegable)

    cliente_final = sel_cliente
    if sel_cliente == opcion_nuevo:
        nuevo_nombre = st.text_input("Escriba el nombre del nuevo cliente:")
        if st.button("Guardar cliente"):
            nombre_limpio = nuevo_nombre.strip()
            if (
                nombre_limpio
                and nombre_limpio not in st.session_state["clientes_lista"]
            ):
                st.session_state["clientes_lista"].append(nombre_limpio)
                with open(ARCHIVO_CLIENTES, "a", encoding="utf-8") as f:
                    f.write(f"{nombre_limpio}\n")
                st.success(f"Cliente '{nombre_limpio}' guardado.")
                st.rerun()
            else:
                st.warning("Nombre no válido o ya existente.")
        cliente_final = None

with col2:
    monto_total = st.number_input(
        "Monto Total ($):", min_value=0.0, step=50.0, value=0.0
    )
    monto_abonado = st.number_input(
        "Abono Inicial ($):",
        min_value=0.0,
        max_value=float(monto_total) if monto_total > 0 else 0.0,
        step=50.0,
        value=0.0,
    )

    monto_restante = monto_total - monto_abonado
    st.metric(
        label="Restante por pagar",
        value=f"${monto_restante:,.2f}",
        delta=f"-${monto_abonado:,.2f}" if monto_abonado > 0 else None,
    )

st.divider()

# ----------------------------------------------------
# 4. BOTÓN PARA GUARDAR NUEVO PEDIDO
# ----------------------------------------------------
if st.button("💾 Guardar Pedido", use_container_width=True):
    if not cliente_final or cliente_final == opcion_nuevo:
        st.error("❌ Por favor selecciona o agrega un cliente válido.")
    elif monto_total <= 0:
        st.error("❌ El monto total debe ser mayor a 0.")
    else:
        # Crear la nueva fila
        nueva_fila = pd.DataFrame(
            [
                {
                    "Pagina": sel_pag,
                    "Zapato": sel_zapato,
                    "Cliente": cliente_final,
                    "Total": monto_total,
                    "Abonado": monto_abonado,
                    "Restante": monto_restante,
                }
            ]
        )

        # Concatenar a nuestro DataFrame
        df_actual = pd.concat([df_actual, nueva_fila], ignore_index=True)

        # Guardar automáticamente en el repositorio de GitHub
        with st.spinner("Guardando pedido en GitHub..."):
            guardar_datos_github(
                df_actual,
                sha_archivo,
                f"Nuevo pedido guardado: {cliente_final}",
            )

        st.success(f"✅ Pedido guardado en la nube para **{cliente_final}**")
        st.rerun()

st.divider()

# ----------------------------------------------------
# 5. MÓDULO PARA ACTUALIZAR ABONOS EN GITHUB
# ----------------------------------------------------
if not df_actual.empty:
    with st.expander("💳 Registrar un nuevo abono a un pedido existente"):
        opciones_pedidos = [
            f"Fila {idx + 2}: {row['Cliente']} - Pag {row['Pagina']} ({row['Zapato']}) | Deuda: ${row['Restante']}"
            for idx, row in df_actual.iterrows()
        ]

        pedido_seleccionado = st.selectbox(
            "Selecciona el pedido al que abonar:", opciones_pedidos
        )

        fila_idx = opciones_pedidos.index(pedido_seleccionado)
        restante_actual = float(df_actual.loc[fila_idx, "Restante"])

        if restante_actual <= 0:
            st.success("🎉 ¡Este pedido ya está completamente pagado!")
        else:
            nuevo_abono = st.number_input(
                "Monto del nuevo abono ($):",
                min_value=0.01,
                max_value=restante_actual,
                step=50.0,
                value=min(50.0, restante_actual),
            )

            if st.button("➕ Aplicar Abono"):
                # Actualizar montos en el DataFrame
                df_actual.loc[fila_idx, "Abonado"] += nuevo_abono
                df_actual.loc[fila_idx, "Restante"] -= nuevo_abono

                # Guardar cambios en GitHub
                with st.spinner("Actualizando abono en GitHub..."):
                    guardar_datos_github(
                        df_actual,
                        sha_archivo,
                        f"Abono de ${nuevo_abono} a {df_actual.loc[fila_idx, 'Cliente']}",
                    )

                st.success(
                    f"✅ Se abonaron ${nuevo_abono} al pedido de {df_actual.loc[fila_idx, 'Cliente']}."
                )
                st.rerun()

    # Mostrar la tabla actualizada
    st.subheader("📋 Registros almacenados en GitHub:")
    st.dataframe(df_actual, use_container_width=True)
