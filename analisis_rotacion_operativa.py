"""
Script de análisis de rotación operativa de afiliadores.

Uso:
1. Coloca este archivo en la misma carpeta que tus Excels.
2. Ajusta la sección PARÁMETROS.
3. Ejecuta:
   python analisis_rotacion_operativa.py

Salidas:
- output/resumen_rotacion_operativa.txt
- output/resumen_rotacion_operativa.csv
- output/detalle_*.csv
- output/grafico_universo_completo.png
- output/grafico_comparativo_escenarios.png
"""

from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# =========================================================
# 1. PARÁMETROS
# =========================================================

ARCHIVO_PERSONAL = Path("PERSONAL_YAPE.xlsx")
ARCHIVO_PRODUCCION = Path("CONSOLIDADO_270426.xlsx")
CARPETA_OUTPUT = Path("output")

ANIO = 2026
MES = 4

# Fecha máxima de producción disponible.
# Ejemplo: si tu consolidado llega hasta 27/04/2026, usa 2026-04-27.
FECHA_CORTE_PRODUCCION = pd.Timestamp("2026-04-27")

# Columnas de PERSONAL
COL_CARGO = "CARGO"
COL_ESTADO = "ESTADO (ACTIVO/INACTIVO)"
COL_CELULAR = "CELULAR"
COL_NOMBRE = "NOMBRE"
COL_INGRESO = "FECHA DE INGRESO"
COL_SALIDA = "FECHA DE SALIDA"

# Columnas de PRODUCCIÓN
COL_ID_EJECUTIVO = "ID EJECUTIVO"
COL_FECHA_PROD = "fecha"

# Escenarios a evaluar:
# El KPI excluye personas con días transaccionados menores al mínimo.
# 0 = universo completo, sin excluir por transacciones.
ESCENARIOS = [
    {
        "nombre": "Universo completo",
        "min_dias_transaccionados": 0,
        "descripcion": "Todos los afiliadores activos registrados en RRHH",
    },
    {
        "nombre": "Excluye sin producción",
        "min_dias_transaccionados": 1,
        "descripcion": "Excluye afiliadores con 0 días transaccionados",
    },
    {
        "nombre": "Excluye < 2 días",
        "min_dias_transaccionados": 2,
        "descripcion": "Excluye afiliadores con 0 o 1 día transaccionado",
    },
    {
        "nombre": "Excluye < 3 días",
        "min_dias_transaccionados": 3,
        "descripcion": "Excluye afiliadores con menos de 3 días transaccionados",
    },
    {
        "nombre": "Excluye < 7 días",
        "min_dias_transaccionados": 7,
        "descripcion": "Excluye afiliadores con menos de 7 días transaccionados",
    },
]


# =========================================================
# 2. FUNCIONES
# =========================================================

def normalizar_texto(serie: pd.Series) -> pd.Series:
    """Convierte a texto, limpia espacios y elimina .0 típico de Excel."""
    return (
        serie
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
    )


def cargar_personal(ruta: Path) -> pd.DataFrame:
    """Carga y normaliza la base de personal."""
    df = pd.read_excel(ruta)
    df.columns = df.columns.str.strip()

    for col in [COL_INGRESO, COL_SALIDA]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in [COL_CARGO, COL_ESTADO, COL_CELULAR, COL_NOMBRE]:
        if col in df.columns:
            df[col] = normalizar_texto(df[col])

    df[COL_CARGO] = df[COL_CARGO].str.upper()
    df[COL_ESTADO] = df[COL_ESTADO].str.upper()

    return df


def cargar_produccion(ruta: Path) -> pd.DataFrame:
    """Carga y normaliza la base de producción."""
    df = pd.read_excel(ruta, dtype=str)
    df.columns = df.columns.str.strip()

    df[COL_ID_EJECUTIVO] = normalizar_texto(df[COL_ID_EJECUTIVO])
    df[COL_FECHA_PROD] = pd.to_datetime(df[COL_FECHA_PROD], errors="coerce")

    return df


def contar_dias_laborales(fecha_inicio, fecha_fin) -> int:
    """Cuenta lunes a viernes entre dos fechas, incluyendo inicio y fin."""
    if pd.isna(fecha_inicio) or pd.isna(fecha_fin):
        return 0

    fecha_inicio = pd.Timestamp(fecha_inicio).normalize()
    fecha_fin = pd.Timestamp(fecha_fin).normalize()

    if fecha_inicio > fecha_fin:
        return 0

    dias = pd.date_range(fecha_inicio, fecha_fin, freq="D")
    return sum(d.weekday() < 5 for d in dias)


def crear_base_rrhh(df_personal: pd.DataFrame, fecha_corte: pd.Timestamp) -> pd.DataFrame:
    """Afiliadores activos en RRHH hasta la fecha de corte."""
    df_base = df_personal[
        (df_personal[COL_CARGO] == "AFILIADOR") &
        (df_personal[COL_ESTADO] == "ACTIVO") &
        (df_personal[COL_INGRESO] <= fecha_corte)
    ].copy()

    return df_base.drop_duplicates(subset=[COL_CELULAR])


def calcular_dias_transaccionados(
    df_prod: pd.DataFrame,
    fecha_inicio: pd.Timestamp,
    fecha_corte: pd.Timestamp,
) -> pd.DataFrame:
    """Cuenta días únicos con al menos una transacción por ejecutivo."""
    df_prod_mes = df_prod[
        (df_prod[COL_FECHA_PROD] >= fecha_inicio) &
        (df_prod[COL_FECHA_PROD] <= fecha_corte) &
        (df_prod[COL_ID_EJECUTIVO].notna())
    ].copy()

    df_prod_mes["FECHA_DIA"] = df_prod_mes[COL_FECHA_PROD].dt.normalize()

    return (
        df_prod_mes
        .groupby(COL_ID_EJECUTIVO)["FECHA_DIA"]
        .nunique()
        .reset_index(name="DIAS_TRANSACCIONADOS")
        .rename(columns={COL_ID_EJECUTIVO: COL_CELULAR})
    )


def preparar_base_rotacion(
    df_personal: pd.DataFrame,
    df_prod: pd.DataFrame,
    fecha_inicio: pd.Timestamp,
    fecha_corte: pd.Timestamp,
) -> pd.DataFrame:
    """Cruza afiliadores activos RRHH con días transaccionados."""
    df_base_rrhh = crear_base_rrhh(df_personal, fecha_corte)
    dias_transaccionados = calcular_dias_transaccionados(df_prod, fecha_inicio, fecha_corte)

    df_base = df_base_rrhh.merge(
        dias_transaccionados,
        on=COL_CELULAR,
        how="left",
    )

    df_base["DIAS_TRANSACCIONADOS"] = (
        df_base["DIAS_TRANSACCIONADOS"]
        .fillna(0)
        .astype(int)
    )

    df_base["DIAS_HABILES_DISPONIBLES"] = df_base[COL_INGRESO].apply(
        lambda fecha_ingreso: contar_dias_laborales(
            max(fecha_ingreso, fecha_inicio),
            fecha_corte,
        )
    )

    return df_base


def calcular_escenario(
    df_base: pd.DataFrame,
    nombre: str,
    min_dias_transaccionados: int,
    descripcion: str,
) -> tuple[dict, pd.DataFrame, pd.DataFrame]:
    """
    Calcula cumplimiento y rotación operativa para un escenario.
    Retorna resumen, base evaluada y base excluida.
    """
    if min_dias_transaccionados == 0:
        df_evaluados = df_base.copy()
        df_excluidos = df_base.iloc[0:0].copy()
    else:
        df_excluidos = df_base[
            df_base["DIAS_TRANSACCIONADOS"] < min_dias_transaccionados
        ].copy()

        df_evaluados = df_base[
            df_base["DIAS_TRANSACCIONADOS"] >= min_dias_transaccionados
        ].copy()

    df_evaluados["CUMPLIMIENTO"] = (
        df_evaluados["DIAS_TRANSACCIONADOS"] /
        df_evaluados["DIAS_HABILES_DISPONIBLES"]
    )

    df_evaluados["CUMPLIMIENTO"] = (
        df_evaluados["CUMPLIMIENTO"]
        .replace([float("inf"), -float("inf")], 0)
        .fillna(0)
        .clip(lower=0, upper=1)
    )

    df_evaluados["ROTACION_OPERATIVA"] = 1 - df_evaluados["CUMPLIMIENTO"]

    df_evaluados["CUMPLIMIENTO_%"] = (
        df_evaluados["CUMPLIMIENTO"] * 100
    ).round(2)

    df_evaluados["ROTACION_OPERATIVA_%"] = (
        df_evaluados["ROTACION_OPERATIVA"] * 100
    ).round(2)

    cumplimiento_promedio = df_evaluados["CUMPLIMIENTO"].mean() * 100 if len(df_evaluados) else 0
    rotacion_promedio = df_evaluados["ROTACION_OPERATIVA"].mean() * 100 if len(df_evaluados) else 0

    congelados = int((df_base["DIAS_TRANSACCIONADOS"] == 0).sum())
    baja_actividad = int(
        (
            (df_base["DIAS_TRANSACCIONADOS"] > 0) &
            (df_base["DIAS_TRANSACCIONADOS"] < min_dias_transaccionados)
        ).sum()
    ) if min_dias_transaccionados > 0 else 0

    resumen = {
        "Escenario": nombre,
        "Descripción": descripcion,
        "Mínimo días transaccionados": min_dias_transaccionados,
        "Headcount RRHH afiliadores activos": len(df_base),
        "Congelados operativos": congelados,
        "Baja actividad excluida": baja_actividad,
        "Total excluidos": len(df_excluidos),
        "Afiliadores evaluados en KPI": len(df_evaluados),
        "Cumplimiento promedio %": round(cumplimiento_promedio, 2),
        "Rotación operativa promedio %": round(rotacion_promedio, 2),
    }

    return resumen, df_evaluados, df_excluidos


def guardar_resumen_txt(resumenes: list[dict], ruta: Path, fecha_inicio, fecha_corte) -> None:
    """Guarda resumen ejecutivo en TXT."""
    lineas = []
    lineas.append("ANÁLISIS DE ROTACIÓN OPERATIVA DE AFILIADORES")
    lineas.append("=" * 70)
    lineas.append(f"Periodo evaluado: {fecha_inicio.strftime('%d/%m/%Y')} al {fecha_corte.strftime('%d/%m/%Y')}")
    lineas.append("")
    lineas.append("Fórmula utilizada:")
    lineas.append("Rotación operativa = 1 - (Días transaccionados / Días hábiles disponibles)")
    lineas.append("")
    lineas.append("Detalle por escenario:")
    lineas.append("=" * 70)

    for r in resumenes:
        lineas.append("")
        lineas.append(r["Escenario"].upper())
        lineas.append("-" * 70)
        lineas.append(r["Descripción"])
        lineas.append(f"Headcount RRHH afiliadores activos: {r['Headcount RRHH afiliadores activos']}")
        lineas.append(f"Congelados operativos: {r['Congelados operativos']}")
        lineas.append(f"Baja actividad excluida: {r['Baja actividad excluida']}")
        lineas.append(f"Total excluidos: {r['Total excluidos']}")
        lineas.append(f"Afiliadores evaluados en KPI: {r['Afiliadores evaluados en KPI']}")
        lineas.append(f"Cumplimiento promedio %: {r['Cumplimiento promedio %']:.2f}%")
        lineas.append(f"Rotación operativa promedio %: {r['Rotación operativa promedio %']:.2f}%")

    lineas.append("")
    lineas.append("Conclusión:")
    lineas.append(
        "La rotación operativa disminuye conforme se excluyen afiliadores con baja o nula "
        "actividad, lo que permite separar el headcount administrativo del personal "
        "realmente operativo."
    )

    ruta.write_text("\n".join(lineas), encoding="utf-8")


def guardar_grafico_universo(resumen: dict, ruta: Path) -> None:
    """Guarda gráfico de torta del universo completo."""
    sns.set_theme(style="whitegrid")

    values = [
        resumen["Cumplimiento promedio %"],
        resumen["Rotación operativa promedio %"],
    ]
    labels = ["Cumplimiento operativo", "Rotación operativa"]

    plt.figure(figsize=(7, 7))
    plt.pie(
        values,
        labels=labels,
        autopct="%1.2f%%",
        startangle=90,
        counterclock=False,
        wedgeprops={"edgecolor": "white"},
    )
    plt.title(
        f"Rotación Operativa General\n"
        f"Universo completo: {resumen['Afiliadores evaluados en KPI']} afiliadores",
        fontsize=13,
        fontweight="bold",
    )
    plt.tight_layout()
    plt.savefig(ruta, dpi=300, bbox_inches="tight")
    plt.close()


def guardar_grafico_comparativo(resumenes: list[dict], ruta: Path) -> None:
    """Guarda subplot 2x2 con los cuatro escenarios depurados."""
    sns.set_theme(style="whitegrid")

    escenarios = [r for r in resumenes if r["Escenario"] != "Universo completo"]

    fig, axes = plt.subplots(2, 2, figsize=(13, 10))
    axes = axes.flatten()

    for ax, r in zip(axes, escenarios):
        values = [
            r["Cumplimiento promedio %"],
            r["Rotación operativa promedio %"],
        ]
        labels = ["Cumplimiento", "Rotación"]

        ax.pie(
            values,
            labels=labels,
            autopct="%1.2f%%",
            startangle=90,
            counterclock=False,
            wedgeprops={"edgecolor": "white"},
        )

        ax.set_title(
            f"{r['Escenario']}\n"
            f"Evaluados: {r['Afiliadores evaluados en KPI']} | "
            f"Excluidos: {r['Total excluidos']}",
            fontsize=11,
            fontweight="bold",
        )

    plt.suptitle(
        "Comparativo de Rotación Operativa por Criterio de Exclusión",
        fontsize=15,
        fontweight="bold",
    )
    plt.tight_layout()
    plt.savefig(ruta, dpi=300, bbox_inches="tight")
    plt.close()


# =========================================================
# 3. EJECUCIÓN
# =========================================================

def main() -> None:
    CARPETA_OUTPUT.mkdir(parents=True, exist_ok=True)

    fecha_inicio = pd.Timestamp(year=ANIO, month=MES, day=1)

    if FECHA_CORTE_PRODUCCION is None:
        fecha_corte = fecha_inicio + pd.offsets.MonthEnd(0)
    else:
        fecha_corte = pd.Timestamp(FECHA_CORTE_PRODUCCION)

    df_personal = cargar_personal(ARCHIVO_PERSONAL)
    df_prod = cargar_produccion(ARCHIVO_PRODUCCION)

    df_base_rotacion = preparar_base_rotacion(
        df_personal=df_personal,
        df_prod=df_prod,
        fecha_inicio=fecha_inicio,
        fecha_corte=fecha_corte,
    )

    resumenes = []
    detalles = {}

    for escenario in ESCENARIOS:
        resumen, df_evaluados, df_excluidos = calcular_escenario(
            df_base=df_base_rotacion,
            nombre=escenario["nombre"],
            min_dias_transaccionados=escenario["min_dias_transaccionados"],
            descripcion=escenario["descripcion"],
        )

        resumenes.append(resumen)
        detalles[escenario["nombre"]] = {
            "evaluados": df_evaluados,
            "excluidos": df_excluidos,
        }

    df_resumen = pd.DataFrame(resumenes)

    # Exportar resumen
    df_resumen.to_csv(
        CARPETA_OUTPUT / "resumen_rotacion_operativa.csv",
        index=False,
        encoding="utf-8-sig",
    )

    guardar_resumen_txt(
        resumenes=resumenes,
        ruta=CARPETA_OUTPUT / "resumen_rotacion_operativa.txt",
        fecha_inicio=fecha_inicio,
        fecha_corte=fecha_corte,
    )

    # Exportar detalles por escenario
    columnas_detalle = [
        COL_NOMBRE,
        COL_CELULAR,
        COL_CARGO,
        COL_ESTADO,
        COL_INGRESO,
        COL_SALIDA,
        "DIAS_TRANSACCIONADOS",
        "DIAS_HABILES_DISPONIBLES",
        "CUMPLIMIENTO_%",
        "ROTACION_OPERATIVA_%",
    ]

    for nombre, bases in detalles.items():
        nombre_archivo = (
            nombre
            .lower()
            .replace(" ", "_")
            .replace("<", "menor_")
            .replace(">", "mayor_")
        )

        df_eval = bases["evaluados"].copy()
        cols_eval = [c for c in columnas_detalle if c in df_eval.columns]
        df_eval[cols_eval].to_csv(
            CARPETA_OUTPUT / f"detalle_evaluados_{nombre_archivo}.csv",
            index=False,
            encoding="utf-8-sig",
        )

        df_exc = bases["excluidos"].copy()
        cols_exc = [c for c in columnas_detalle if c in df_exc.columns]
        df_exc[cols_exc].to_csv(
            CARPETA_OUTPUT / f"detalle_excluidos_{nombre_archivo}.csv",
            index=False,
            encoding="utf-8-sig",
        )

    # Gráficos
    guardar_grafico_universo(
        resumen=resumenes[0],
        ruta=CARPETA_OUTPUT / "grafico_universo_completo.png",
    )

    guardar_grafico_comparativo(
        resumenes=resumenes,
        ruta=CARPETA_OUTPUT / "grafico_comparativo_escenarios.png",
    )

    # Mostrar en consola
    print("=" * 70)
    print("ANÁLISIS FINALIZADO")
    print("=" * 70)
    print(f"Periodo: {fecha_inicio.strftime('%d/%m/%Y')} al {fecha_corte.strftime('%d/%m/%Y')}")
    print(f"Archivos generados en: {CARPETA_OUTPUT.resolve()}")
    print("=" * 70)
    print(df_resumen.to_string(index=False))


if __name__ == "__main__":
    main()
