import os
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="PEDIDOS DE ZAPATOS MINGA INC", page_icon="👠", layout="wide"
)

# ----------------------------------------------------
# 1. ARCHIVOS Y CONFIGURACIÓN INICIAL
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

lista_pagina = [x for x in range(1, 50)]
zapatos_dicc = {1: ["A1", "A2", "A3"], 2: ["B1", "B2", "B3"]}

dia, mes, anio = (
    ("0" + str(datetime.now().day))[-2:],
    ("0" + str(datetime.now().month))[-2:],
    str(datetime.now().year),
)
archivo_nombre = f"PEDIDOS_PRUEBA_{dia}{mes}{anio}.xlsx"

st.title("👠 Pedidos de Zapatos Minga Inc")

# ----------------------------------------------------
# 2. CAPTURA DEL PEDIDO
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
            if nombre_limpio and nombre_limpio not in st.session_state["clientes_lista"]:
                st.session_state["clientes_lista"].append(nombre_limpio)
                with open(ARCHIVO_CLIENTES, "a", encoding="utf-8") as f:
                    f.write(f"{nombre_limpio}\n")
                st.success(f"Cliente '{nombre_limpio}' guardado.")
                st.rerun()
            else:
                st.warning("Nombre no válido o ya existente.")
        cliente_final = None

with col2:
    # NUEVOS CAMPOS: Total y Abonado
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
# 3. BOTÓN PARA GUARDAR NUEVO PEDIDO
# ----------------------------------------------------
if st.button("💾 Guardar Pedido", use_container_width=True):
    if not cliente_final or cliente_final == opcion_nuevo:
        st.error("❌ Por favor selecciona o agrega un cliente válido.")
    elif monto_total <= 0:
        st.error("❌ El monto total debe ser mayor a 0.")
    else:
        # Crear archivo con las nuevas columnas si no existe
        if not os.path.exists(archivo_nombre):
            df_inicial = pd.DataFrame(
                columns=[
                    "Pagina",
                    "Zapato",
                    "Cliente",
                    "Total",
                    "Abonado",
                    "Restante",
                ]
            )
            df_inicial.to_excel(archivo_nombre, index=False)

        # Guardar la fila en Excel
        wb = load_workbook(archivo_nombre)
        ws = wb.active
        ws.append(
            [
                sel_pag,
                sel_zapato,
                cliente_final,
                monto_total,
                monto_abonado,
                monto_restante,
            ]
        )
        wb.save(archivo_nombre)
        wb.close()

        st.success(f"✅ Pedido guardado para **{cliente_final}**")
        st.rerun()

st.divider()

# ----------------------------------------------------
# 4. MÓDULO PARA ACTUALIZAR ABONOS DE PEDIDOS EXISTENTES
# ----------------------------------------------------
if os.path.exists(archivo_nombre):
    df_actual = pd.read_excel(archivo_nombre)

    with st.expander("💳 Registrar un nuevo abono a un pedido existente"):
        if not df_actual.empty:
            # Creamos una lista identificando cada fila por su índice en Excel
            opciones_pedidos = [
                f"Fila {idx + 2}: {row['Cliente']} - Pag {row['Pagina']} ({row['Zapato']}) | Deuda: ${row['Restante']}"
                for idx, row in df_actual.iterrows()
            ]

            pedido_seleccionado = st.selectbox(
                "Selecciona el pedido al que abonar:", opciones_pedidos
            )

            # Obtener el número de fila seleccionado
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
                    # Actualizamos los valores en el DataFrame
                    df_actual.loc[fila_idx, "Abonado"] += nuevo_abono
                    df_actual.loc[fila_idx, "Restante"] -= nuevo_abono

                    # Guardamos el Excel actualizado
                    df_actual.to_excel(archivo_nombre, index=False)
                    st.success(
                        f"✅ Se abonaron ${nuevo_abono} al pedido de {df_actual.loc[fila_idx, 'Cliente']}."
                    )
                    st.rerun()

    # Mostrar la tabla actualizada de Excel
    st.subheader("📋 Registros de hoy:")
    st.dataframe(df_actual, use_container_width=True)