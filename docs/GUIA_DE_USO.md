# Guía de uso — Reportes con SQL en Access

## Qué necesitas
- Windows con **Microsoft Access** (2016, 2019, 2021 o Microsoft 365).

## 1. Descarga el proyecto
En la página del repositorio pulsa **Code → Download ZIP** y descomprímelo, por ejemplo en `C:\sql-access-reportes`. Deben quedar las carpetas `datos`, `consultas` y `vba`.

## 2. Crea la base de datos
1. Abre Access → **Base de datos en blanco**.
2. Guárdala **dentro de la carpeta del proyecto** (por ejemplo `C:\sql-access-reportes\Reportes.accdb`). Debe estar ahí porque el módulo busca las carpetas `datos` y `consultas` al lado de la base.
3. Si aparece la barra amarilla de seguridad, pulsa **Habilitar contenido**.

## 3. Carga las tablas y las consultas automáticamente
1. Pulsa **Alt + F11** para abrir el editor de VBA.
2. Menú **Archivo → Importar archivo…** y elige `vba\CrearBase.bas`.
3. Haz doble clic en el módulo **CrearBase**, pon el cursor dentro de `Sub CrearTodo()` y pulsa **F5**.
4. Al terminar aparece el mensaje *"Base lista: 4 tablas y 6 consultas creadas."*

Si **Importar archivo** da error, usa el plan B: **Insertar → Módulo**, abre `CrearBase.bas` con el Bloc de notas, copia todo **menos la primera línea** (`Attribute VB_Name…`) y pégalo en el módulo nuevo.

Si sale el error *"Tipo definido por el usuario no definido"* en `DAO.Recordset`: en el editor de VBA ve a **Herramientas → Referencias** y marca **Microsoft Office … Access database engine Object Library**.

> Puedes ejecutar `CrearTodo` otra vez cuando quieras: borra las tablas y las vuelve a crear desde los CSV.

## 4. Ejecuta las consultas
En el panel de la izquierda verás las consultas `q01_…` a `q06_…`. Haz doble clic en cada una.

| Consulta | Qué te va a pedir |
|---|---|
| q01_ficha_puntos_por_red | La red. Prueba con **RN01**, RN02, RN03 o RS01. Crea la tabla `Tbl_Ficha de Puntos`. Si ya existe, Access pregunta si quieres reemplazarla. |
| q04_cambios_de_estado_por_mes | Fecha inicial y fecha final, por ejemplo **01/01/2026** y **30/06/2026**. |
| Las demás | Nada, se ejecutan directo. |

Para ver o editar el SQL: clic derecho sobre la consulta → **Vista Diseño** → **Ver → Vista SQL**.

## 5. Hazlo a mano (opcional)
Si prefieres no usar VBA:
1. **Datos externos → Nuevo origen de datos → Desde archivo → Texto** e importa cada CSV de `datos` con estos nombres de tabla: `PuntosVenta`, `Direcciones`, `Marcas` y `Geolocalizacion`.
2. **Crear → Diseño de consulta → Vista SQL**, pega el contenido de cada archivo `.sql` y guárdala. Access no acepta las líneas que empiezan con `--`: bórralas antes de guardar.

## Volver a generar los datos (opcional)
Los CSV se crean con `datos/generar_datos.py` (Python 3, sin librerías extra):

```
python datos/generar_datos.py
```

Siempre produce los mismos datos ficticios.
