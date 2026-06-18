# 🏨 Conciliación de Reservas — Aptour

Aplicación web desarrollada con **Streamlit** para conciliar automáticamente tres informes de reservas y detectar diferencias entre el sistema interno, el archivo del operador mayorista y el sistema de cobros y pagos.

---

## ¿Qué hace esta aplicación?

Permite cruzar y reconciliar información proveniente de tres fuentes distintas:

| Archivo | Descripción |
|---|---|
| `listsistema` | Informe interno del sistema de reservas |
| `listoperador` | Archivo de reservas recibido del operador mayorista |
| `listcobros` | Informe de cobros y pagos del sistema interno |

El resultado es un archivo Excel descargable con la columna **Estado** completada y las discrepancias económicas **marcadas en rojo** para revisión.

---

## Lógica de conciliación

### 1. Preparación de datos
La columna `NOM_CLI` de `listcobros` se divide automáticamente en dos:
- **APELLIDO** — primer token del nombre completo
- **NOMBRES** — resto del nombre

### 2. Cruce por número de reserva *(Paso 1)*
Se compara la columna **Nº de Reserva** de `listoperador` contra todos los códigos que figuran separados por coma en la columna **DETALLE** de `listsistema`.

Si hay coincidencia → se trae el valor de **ID_RES** (sin espacios) a la columna **Estado** de `listoperador`.

### 3. Cruce por nombre de cliente *(Paso 2)*
Para las filas que quedaron sin Estado luego del Paso 1, se compara:
- **APELLIDO** de `listcobros` con algún token de la columna **Cliente** de `listoperador`
- Al menos un token de **NOMBRES** con algún token de **Cliente**

Si coinciden ambas partes → se trae el **ID_RES** correspondiente a la columna **Estado**.

### 4. Marcado de diferencias en rojo
Una vez asignado el ID_RES, se evalúan dos condiciones económicas. Si se cumple **alguna de las dos**, la celda **Estado** se pinta en **rojo**:

- `|COSTOUS (sistema) − Importe (operador)| > 5`
- `|VENTAUS − COBROSUS (cobros)| > 500`

Las celdas que **no tienen coincidencia** en ninguno de los dos pasos quedan vacías.

---

## Tecnologías utilizadas

- [Streamlit](https://streamlit.io/) — interfaz web
- [Pandas](https://pandas.pydata.org/) — procesamiento de datos
- [OpenPyXL](https://openpyxl.readthedocs.io/) — generación y formato del Excel de salida
- [xlrd](https://pypi.org/project/xlrd/) — lectura de archivos `.xls` legacy

---

## Instalación local

### Requisitos
- Python 3.9 o superior

### Pasos

```bash
# 1. Clonar el repositorio
git clone https://github.com/tu-usuario/tu-repo.git
cd tu-repo

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Ejecutar la aplicación
streamlit run app.py
```

La app se abre automáticamente en `http://localhost:8501`

---

## Estructura del repositorio

```
/
├── app.py               # Aplicación principal
├── requirements.txt     # Dependencias de Python
└── README.md            # Este archivo
```

---

## Formatos de archivo soportados

La aplicación acepta tanto `.xlsx` como `.xls` en los tres campos de carga, utilizando automáticamente el motor de lectura adecuado para cada formato.

---

## Salida

Al finalizar el procesamiento se muestra:
- Un resumen con la cantidad de filas conciliadas, marcadas en rojo y sin coincidencia.
- Una vista previa de la tabla resultado con las filas problemáticas resaltadas.
- Un botón para descargar el archivo `conciliacion.xlsx` con formato profesional.

## Link
https://conciliadorst.streamlit.app/

