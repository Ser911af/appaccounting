# Conciliación de Cartera: Cierre vs Balance

Herramienta Streamlit para conciliar la cartera de apartamentos entre dos fuentes contables: el reporte de **Cierre** y el **Balance** de cuentas 1345\*.

## Qué hace

1. Carga dos archivos Excel (Cierre y Balance) con estructura variable.
2. Permite indicar en qué fila están los encabezados de cada archivo.
3. Detecta automáticamente las columnas relevantes (código, bloque, valor cobro, nuevo saldo, cuenta).
4. Construye una clave única por apartamento en formato `piso-número` (ej. `1-9801`, `2-203`) combinando las columnas de Bloque y Código del Cierre.
5. En el Balance filtra únicamente las cuentas que empiezan con `1345*`.
6. Hace un outer join por clave de apartamento y calcula la diferencia entre el Valor Cobro (Cierre) y el Nuevo Saldo 1345 (Balance).
7. Muestra métricas, tablas detalladas por categoría y permite descargar todos los resultados en Excel.

## Estructura del proyecto

```
appaccounting/
├── app/
│   └── app.py          # Lógica principal de la app
├── requirements.txt
└── README.md
```

## Dependencias

```
streamlit
pandas
openpyxl
xlsxwriter
numpy
```

Instalar:

```bash
pip install -r requirements.txt
```

## Cómo ejecutar

```bash
streamlit run app/app.py
```

La app queda disponible en `http://localhost:8501`.

## Flujo de uso

### Paso 1 — Cargar archivos
Sube el archivo de **Cierre** y el de **Balance** (ambos `.xlsx`).

### Paso 2 — Indicar fila de encabezados
Cada archivo puede tener filas de título o metadatos antes de los encabezados reales. Indica el número de fila donde están los nombres de columna. La vista previa se actualiza en tiempo real.

### Paso 3 — Revisar columnas del Cierre
La app sugiere automáticamente:
- **Inmueble Código**: número del apartamento dentro del bloque.
- **Inmueble Bloque**: número del bloque o torre.
- **Valor Cobro**: monto facturado/cobrado.

Si la sugerencia es incorrecta, cámbiala en el selector.

### Paso 4 — Revisar columnas del Balance
La app sugiere:
- **Clave de apartamento**: columna con valores tipo `1-101`, `2-203`.
- **Nuevo Saldo**: saldo contable del apartamento.
- **Cuenta**: columna de código contable (se filtran solo las que empiezan con `1345`).

### Paso 5 — Ver resultados
Las métricas muestran de un vistazo:

| Métrica | Descripción |
|---|---|
| Aptos en Cierre | Total de apartamentos con clave válida en el Cierre |
| Aptos en Balance (1345\*) | Total en Balance con cuentas 1345\* |
| Diferencias ≠ 0 | Aptos donde Cierre ≠ Balance |
| Solo en Cierre | Cobros sin saldo contable registrado |
| Solo en Balance | Saldo contable sin cobro en Cierre |
| Cobro > 0 y Saldo = 0 | Cobrado pero sin contrapartida 1345 |
| Saldo > 0 y Cobro = 0 | Saldo 1345 sin cobro |
| Ambos > 0 pero diferentes | Discrepancia de montos |

Las tablas detalladas están organizadas en pestañas:
- **Conciliación**: solo los apartamentos con diferencia.
- **Match Total**: el outer join completo.
- **Solo Cierre / Solo Balance**: apartamentos que no cruzan.
- **Agregado Cierre / Balance**: totales por apartamento antes del join.

### Paso 6 — Descargar Excel
El botón al final descarga un `.xlsx` con una hoja por cada tabla (agregado_cierre, agregado_balance, match_total, conciliacion, solo_cierre, solo_balance).

## Lógica de clave de apartamento

La clave se construye como `piso-número`, por ejemplo `2-9803`.

- En el **Cierre**: se extrae el número del Bloque (= piso) y el Código del apartamento y se concatenan.
- En el **Balance**: se busca una columna que ya contenga ese formato directamente.
- El patrón es estricto: solo acepta valores cuya totalidad sea `{1-2 dígitos}{separador}{3-5 dígitos}` para evitar confundir NITs como `43.202.550-3` con claves de apto.

## Notas

- La tolerancia de diferencia es `0.01` (se ignoran diferencias menores a $0.01).
- Los montos soportan formato colombiano con punto como separador de miles y coma como decimal (ej. `1.234.567,89`).
