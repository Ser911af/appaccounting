# AccountEase
# AccountEase: Aplicación de Análisis Contable

AccountEase es una aplicación web diseñada para contadores que analiza reportes proporcionados por la DIAN. A partir de un archivo Excel con una estructura predefinida, la herramienta automatiza cálculos y genera reportes personalizados que son descargables.

## Características

- **Cálculo Automático**: Realiza cálculos como la columna `Base` restando `IVA` a `Total`.
- **Análisis Dinámico**: Filtra y agrupa los datos según `Tipo de documento` y `Grado`.
- **Reportes Mensuales**: Genera tablas con la suma de la columna `Base` organizada por meses.
- **Descarga Fácil**: Permite descargar la tabla consolidada en formato Excel.

## Requisitos del Sistema

- Python 3.9 o superior.
- Entorno virtual configurado (opcional pero recomendado).

### Dependencias

Las siguientes librerías son necesarias para ejecutar la aplicación:

- `streamlit`
- `pandas`
- `openpyxl`

Puedes instalarlas ejecutando:
```bash
pip install -r requirements.txt
```

## Estructura del Proyecto

```plaintext
/workspaces/AccountEase
├── app
│   └── app.py         # Código principal de la aplicación
├── data
│   └── example.xlsx   # Archivo de ejemplo para pruebas
├── requirements.txt   # Lista de dependencias
└── README.md          # Documentación del proyecto
```

## Cómo Ejecutar la Aplicación

1. Clona este repositorio:
   ```bash
   git clone https://github.com/tu-usuario/AccountEase.git
   cd AccountEase
   ```

2. Crea un entorno virtual e instálalo:
   ```bash
   python -m venv venv
   source venv/bin/activate   # En Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. Ejecuta la aplicación:
   ```bash
   streamlit run app/app.py
   ```

4. Abre tu navegador en la URL proporcionada (por defecto: `http://localhost:8501`).

## Uso

1. **Carga de Archivo**: Sube el archivo Excel proporcionado por la DIAN.
2. **Análisis Automático**: La aplicación procesa los datos y genera un reporte basado en:
   - `Tipo de documento`
   - `Grado` (Emitido o Recibido)
   - Sumas mensuales de la columna `Base`.
3. **Descarga del Reporte**: Descarga la tabla consolidada en formato Excel.

### Formato de la Tabla Generada
La tabla consolidada tiene el siguiente formato:

| Tipo Doc     | Grado      | Enero | Febrero | ... | Diciembre | Total |
|--------------|------------|-------|---------|-----|-----------|-------|
| Factura      | Emitido    | 1000  | 2000    | ... | 1500      | 5500  |
| Factura      | Recibido   | 500   | 800     | ... | 400       | 1700  |
| Nota Crédito | Emitido    | 300   | 600     | ... | 200       | 1100  |

## Contribuciones

Si deseas contribuir:

1. Crea un fork del proyecto.
2. Crea una rama para tu feature/bugfix:
   ```bash
   git checkout -b feature-nombre
   ```
3. Realiza tus cambios y haz un commit.
4. Envía un pull request.

## Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo LICENSE para más detalles.

## Contacto

Para consultas o soporte, puedes contactarme en:
- **Email**: tuemail@example.com
- **GitHub**: [Tu Usuario](https://github.com/tu-usuario)

¡Gracias por usar AccountEase!
