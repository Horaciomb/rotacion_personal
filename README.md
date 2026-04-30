# Análisis de Rotación Operativa de Afiliadores

Este proyecto permite calcular la **rotación operativa** y el **cumplimiento laboral** de afiliadores, cruzando la base de personal con la base de producción.

El análisis no se basa únicamente en el estado administrativo del personal, sino en la actividad real registrada en producción.

---

## Objetivo

Medir qué tan activo estuvo el equipo de afiliadores durante un mes determinado, considerando:

- afiliadores activos en la base de personal,
- días hábiles disponibles según fecha de ingreso,
- días en los que realmente realizaron afiliaciones,
- exclusión de personal con baja o nula actividad,
- generación de reportes y gráficos automáticos.

La fórmula principal utilizada es:

```text
Rotación operativa = 1 - (Días transaccionados / Días hábiles disponibles)
```

---

## Estructura recomendada del proyecto

```text
proyecto_rotacion/
│
├── analisis_rotacion_operativa.py
├── requirements.txt
├── README.md
│
├── PERSONAL_YAPE.xlsx
├── CONSOLIDADO_270426.xlsx
│
└── output/
```

La carpeta `output/` se crea automáticamente al ejecutar el script.

---

## Crear entorno virtual

Desde la carpeta del proyecto, ejecutar:

### Windows

```bash
python -m venv venv
venv\Scripts\activate
```

### macOS / Linux

```bash
python3 -m venv venv
source venv/bin/activate
```

---

## Instalar dependencias

Con el entorno virtual activado:

```bash
pip install -r requirements.txt
```

Un `requirements.txt` mínimo debería incluir:

```text
pandas
openpyxl
matplotlib
seaborn
```

---

## Configuración del script

En la parte superior del archivo `analisis_rotacion_operativa.py` se deben ajustar estos valores:

```python
ARCHIVO_PERSONAL = Path("PERSONAL_YAPE.xlsx")
ARCHIVO_PRODUCCION = Path("CONSOLIDADO_270426.xlsx")

ANIO = 2026
MES = 4
FECHA_CORTE_PRODUCCION = pd.Timestamp("2026-04-27")
```

### Parámetros principales

- `ARCHIVO_PERSONAL`: archivo Excel con la base de personal.
- `ARCHIVO_PRODUCCION`: archivo Excel con la producción o afiliaciones.
- `ANIO`: año que se desea analizar.
- `MES`: mes que se desea analizar.
- `FECHA_CORTE_PRODUCCION`: fecha máxima hasta la cual se tomará la producción.

---

## Columnas esperadas

### Archivo de personal

El archivo de personal debe contener, como mínimo:

```text
NOMBRE
CELULAR
CARGO
ESTADO (ACTIVO/INACTIVO)
FECHA DE INGRESO
FECHA DE SALIDA
```

### Archivo de producción

El archivo de producción debe contener, como mínimo:

```text
ID EJECUTIVO
fecha
```

El cruce se realiza entre:

```text
PERSONAL_YAPE.xlsx -> CELULAR
CONSOLIDADO_270426.xlsx -> ID EJECUTIVO
```

---

## Ejecutar el análisis

Con el entorno virtual activado:

```bash
python analisis_rotacion_operativa.py
```

---

## Resultados generados

Al ejecutar el script, se genera una carpeta `output/` con archivos como:

```text
output/
│
├── resumen_rotacion_operativa.txt
├── resumen_rotacion_operativa.csv
├── detalle_rotacion_universo_completo.csv
├── detalle_rotacion_min_1_dia.csv
├── detalle_rotacion_min_2_dias.csv
├── detalle_rotacion_min_3_dias.csv
├── detalle_rotacion_min_7_dias.csv
├── grafico_universo_completo.png
└── grafico_comparativo_escenarios.png
```

---

## Escenarios calculados

El script calcula varios escenarios para comparar la rotación operativa:

### 1. Universo completo

Considera todos los afiliadores activos registrados en RRHH.

### 2. Excluyendo personas sin producción

Considera solo afiliadores que realizaron al menos una afiliación.

### 3. Excluyendo personas con menos de 2 días transaccionados

Retira afiliadores que solo tuvieron actividad en 1 día.

### 4. Excluyendo personas con menos de 3 días transaccionados

Retira afiliadores con actividad muy baja entre 1 y 2 días.

### 5. Excluyendo personas con menos de 7 días transaccionados

Evalúa únicamente al núcleo operativo más constante.

---

## Interpretación de indicadores

### Cumplimiento promedio

Mide qué porcentaje de los días hábiles disponibles fueron efectivamente trabajados.

```text
Cumplimiento = Días transaccionados / Días hábiles disponibles
```

### Rotación operativa promedio

Mide el porcentaje de días hábiles no trabajados.

```text
Rotación operativa = 1 - Cumplimiento
```

Una rotación operativa alta puede indicar:

- baja actividad,
- abandono operativo,
- personal congelado,
- falta de seguimiento,
- baja productividad,
- registros activos en RRHH pero sin operación real.

---

## Recomendación de uso

Para seguimiento operativo, se recomienda revisar principalmente los escenarios:

- excluyendo afiliadores sin producción,
- excluyendo afiliadores con menos de 2 días,
- excluyendo afiliadores con menos de 3 días.

Estos escenarios ayudan a limpiar el indicador y evitar que personas sin actividad real distorsionen el análisis.

El escenario de menos de 7 días puede usarse para analizar únicamente el grupo más estable y constante del equipo.

---

## Notas importantes

- El script no modifica los archivos originales.
- Todos los resultados se exportan a la carpeta `output/`.
- Para analizar otro mes, solo se deben cambiar `ANIO`, `MES` y `FECHA_CORTE_PRODUCCION`.
- Las fechas deben estar correctamente registradas en los archivos Excel.
- El campo `CELULAR` en personal debe coincidir con `ID EJECUTIVO` en producción.

---

## Autor

Horacio Molina, AKA El Terrror de las Colegas