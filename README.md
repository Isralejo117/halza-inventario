# Inventario Halza Medic

Aplicación de escritorio para registrar los **barridos mensuales de inventario por lote y fecha de caducidad** en el almacén de Halza Medic. Complementa al sistema comercial (Contpaqi), que no maneja de forma ágil la subdivisión de existencias por lote.

Proyecto de prácticas profesionales — Ingeniería en Tecnología de Software, UANL.

## Funciones

- **Captura** de partidas: un mismo producto puede tener varios lotes con distinta existencia y caducidad.
- **Productos que no caducan** (sillas, tijeras, batas…): casilla «No caduca». Internamente se guardan con la fecha 31/12/9999, así quedan al final al ordenar por caducidad y cumplen cualquier mínimo de vida útil.
- **Un código = un producto**: al escribir un código ya registrado se llenan solos su nombre y almacén, y no se permite usarlo para otro producto. Si se corrige el nombre o el almacén, se corrige en todas sus partidas.
- **Validación al escribir**: las fechas se formatean solas (`03062028` → `03/06/2028`) y no se aceptan días o meses que no existen (incluye años bisiestos).
- **Consulta** por mes: cada producto ocupa una fila con su existencia total; si tiene varios lotes, se despliega para ver cada lote con su caducidad y cantidad. Los lotes se pueden editar y eliminar.
- **Aviso de caducidad**: en Consulta se marcan en rojo las partidas con menos de 1 año para caducar (algunas empresas no aceptan productos que caduquen antes).
- **Comparación** lote por lote contra el mes anterior: lotes nuevos, que ya no están, que cambiaron o sin cambio.
- **Reporte PDF** imprimible, separado por almacén, con subtotales y firmas. También hay un **PDF de la comparación** contra el mes anterior, con los colores de cada estado.
- **Exportación a Excel** con filtros, fechas reales, los lotes con menos de 1 año para caducar en rojo y la hoja de comparación.
- **Respaldos**: botón de respaldo manual y respaldo automático diario (se conservan los últimos 30).

## Tecnologías

- Python 3.11
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) — interfaz gráfica
- SQLite (`sqlite3`) — base de datos local en un solo archivo
- [ReportLab](https://www.reportlab.com/) — PDF
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel
- [PyInstaller](https://pyinstaller.org/) — empaquetado como `.exe`

## Cómo ejecutarlo

```
pip install -r requirements.txt
python interfaz.py
```

La primera vez se crea `inventario_halza.db` en la carpeta del programa.

## Cómo crear el `.exe`

Doble clic en `crear_exe.bat`. El programa queda en `dist\HalzaInventario` y no necesita Python instalado.

Para instalarlo, copia esa carpeta a otra ubicación (por ejemplo `C:\HalzaInventario`) y crea un acceso directo a `HalzaInventario.exe`. La base de datos y los respaldos se guardan junto al `.exe`.

## Estructura

| Archivo | Responsabilidad |
|---|---|
| `interfaz.py` | Ventana: pestañas Captura, Consulta y Comparar |
| `inventario.py` | Base de datos: guardar, validar, consultar, comparar y respaldar |
| `formato_fechas.py` | Formato y validación de fechas mientras se escribe |
| `reporte.py` | Reporte PDF |
| `exportar_excel.py` | Exportación a Excel |
| `crear_exe.bat` | Crea el `.exe` con PyInstaller |

Las fechas se guardan internamente como `AAAA-MM-DD` (para que se ordenen correctamente) y se muestran al usuario como `dd/mm/aaaa`.
