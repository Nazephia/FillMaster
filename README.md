# FillMaster

Herramienta para **preparar plantillas de Excel de MRW** automáticamente.

A partir de un archivo `.xlsx`, FillMaster puede:

- Rellenar la columna **Estado de destino** según el código de agencia.
- Ajustar los **códigos de agencia** (agregar el cero delante cuando hace falta).
- Dar formato a los **teléfonos** de Venezuela y Perú.

> El archivo original **nunca se modifica**. Siempre se crea una copia nueva.

---

## ¿Cómo funciona?

FillMaster lee la hoja `NACIONAL` del archivo (si no existe, usa la primera hoja) y
trabaja sobre estas columnas:

| Columna | Contenido        | Qué se hace                                  |
|---------|------------------|----------------------------------------------|
| `J`     | Agencia MRW      | Ajusta el código y busca su estado           |
| `L`     | Estado de destino| Se rellena con el estado según la agencia    |
| `G`     | Teléfono         | Se normaliza el formato                      |

La fila de encabezados se detecta automáticamente buscando los textos
`AGENCIA MRW`, `ESTADO DE DESTINO` y `TELEFONO`. Si no se encuentran, se usan
las columnas `J`, `L` y `G` por defecto.

### Estados según el código de agencia

El estado se obtiene con los **miles** del código:

| Código   | Miles | Estado           |
|----------|-------|------------------|
| `1010`   | 1     | Distrito Capital |
| `15000`  | 15    | Miranda          |
| `22001`  | 22    | La Guaira        |

El rango válido es de `1000` a `24999`. Los códigos fuera de rango se reportan
como inválidos y la columna `L` no se toca.

### Formato de teléfonos

| Caso                                  | Resultado        |
|---------------------------------------|------------------|
| `04141234567` (Venezuela correcto)    | Se deja igual    |
| `4141234567` (sin el 0)               | `04141234567`    |
| `+584141234567`                       | `04141234567`    |
| `51987654321` (Perú con 51)           | Se deja igual    |
| `987654321` (Perú sin 51)             | `51987654321`    |
| Cualquier otro formato                | Se marca con `E` |
| Celda con varios teléfonos            | No se modifica   |

---

## Interfaz gráfica (recomendado)

### Opción 1: Ejecutable

No necesita instalar Python.

- **Windows:** ejecutar `FillMaster.exe`
- **Linux:** ejecutar `./FillMaster`

### Opción 2: Desde el código

Requiere **Python 3** con Tkinter (viene incluido en la instalación normal de Python).

```bash
python FillMaster.py
```

### Pasos de uso

1. Presione **Agregar archivos…** y seleccione uno o varios `.xlsx`.
2. Marque las opciones que quiera aplicar:
   - Rellenar «Estado de destino».
   - Ajustar códigos de agencia.
   - Dar formato a los teléfonos.
3. Elija dónde guardar los resultados:
   - Junto al archivo original, o
   - En otra carpeta.
4. Escriba el texto que se agrega al nombre (por defecto `FIX`).
   Ejemplo: `MASTER.xlsx` → `MASTERFIX.xlsx`.
5. Presione **Procesar archivos**.
6. Revise el reporte en pantalla y, si lo desea, guárdelo con **Guardar reporte…**.

La ventana también incluye un botón para cambiar entre **modo claro y oscuro**.

---

## Línea de comandos (CLI)

También puede usarse sin interfaz gráfica:

```bash
# Procesar un archivo concreto
python cli/fillmaster.py MASTER.xlsx

# Procesar todos los .xlsx de la carpeta actual
python cli/fillmaster.py
```

El resultado se guarda junto al original con el sufijo `FIX`:

```
MASTER.xlsx  →  MASTERFIX.xlsx
```

El CLI siempre aplica las tres correcciones y muestra el reporte en la terminal.

---

## Reporte

Al terminar, FillMaster informa:

- Estados escritos.
- Códigos de agencia ajustados.
- Teléfonos reescritos.
- Estados que tenían otro valor y fueron reemplazados.
- Códigos de agencia sin estado válido.
- Teléfonos marcados con `E` para revisar.
- Celdas con varios teléfonos (sin modificar).

---

## Requisitos

- **Python 3.8 o superior** (solo si ejecuta desde el código).
- **Tkinter** para la interfaz gráfica.
- No requiere librerías externas: se usan solo módulos de la biblioteca estándar.

---

## Estructura del proyecto

```
FillMaster/
├── FillMaster.py              # Interfaz gráfica principal
├── cli/
│   └── fillmaster.py          # Versión de línea de comandos
├── gui/
│   └── fillmaster-gui.py      # Copia de la interfaz gráfica
├── executable/
│   └── FillMaster             # Ejecutable para Linux
└── .github/workflows/build.yml # Compilación y publicación automática
```

---

## Compilación del ejecutable

El proyecto usa **PyInstaller**. El flujo de GitHub Actions genera los binarios
automáticamente al crear un tag `v*` (por ejemplo `v1.0.0`):

- `FillMaster.exe` para Windows.
- `FillMaster-linux.tar.gz` para Linux.

Para compilar manualmente:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name FillMaster FillMaster.py
```

El ejecutable se genera en la carpeta `dist/`.
