# Ingreso Masivo Xpress El Salvador
### Versión 3.9.3

Herramienta de escritorio para distribuir paquetes del archivo `INGRESO_MASIVO.xlsx`
hacia los archivos Excel individuales de cada tienda. Desarrollada en Python/Tkinter.

---

## Índice
1. [Descripción general](#descripción-general)
2. [Flujo de trabajo](#flujo-de-trabajo)
3. [Pantalla principal](#pantalla-principal)
4. [Acciones disponibles](#acciones-disponibles)
5. [Motor de procesamiento](#motor-de-procesamiento)
6. [Lógica de negocio](#lógica-de-negocio)
7. [Sistema de actualizaciones](#sistema-de-actualizaciones)
8. [Configuración](#configuración)
9. [Archivos del proyecto](#archivos-del-proyecto)

---

## Descripción general

El programa lee el archivo `INGRESO_MASIVO.xlsx` fila por fila, identifica la
tienda destino de cada paquete y lo inserta en el archivo Excel correspondiente
de esa tienda. Al final actualiza los resultados en el INGRESO_MASIVO con
`LISTO`, `FALTA` o `DUP` por cada fila procesada.

---

## Flujo de trabajo

```
INGRESO_MASIVO.xlsx
    │
    ├─ Paso 1: Leer todas las filas en memoria (una sola pasada)
    ├─ Paso 2: Clasificar filas por tienda (sin abrir ningún Excel)
    ├─ Paso 3: Por cada tienda → abrir su Excel → insertar todos sus paquetes → guardar
    ├─ Paso 4: Guardar archivos de tiendas (con reintento si están bloqueados)
    ├─ Paso 5: Actualizar hoja FALTA con registros no procesados
    └─ Paso 6: Guardar INGRESO_MASIVO con resultados (LISTO / FALTA / DUP)
```

---

## Pantalla principal

### Panel izquierdo (Sidebar)
Muestra los controles y el estado del programa:
- **Sección ACCIONES**: botones para ejecutar las operaciones principales
- **Sección REPORTES**: acceso a vistas de resultados y diagnósticos
- **Footer**: versión del programa y estado de conexión GitHub

### Panel central
- **Tarjetas de resumen**: PROCESADOS / FALTA / DUPLICADOS / SEGUNDOS / ÚLTIMO
- **Barra de progreso**: avance en tiempo real con porcentaje
- **Log de actividad en tiempo real**: cada fila procesada con estado y ETA
- **Panel "Actividad en tiempo real"**: últimas acciones con timestamp
- **Panel "Estado actual"**: mensaje del paso que se está ejecutando
- **Panel "Último proceso"**: resumen del procesamiento más reciente

---

## Acciones disponibles

### ▶ Ejecutar (F5)
Procesa el `INGRESO_MASIVO.xlsx` completo.

**Qué hace:**
- Lee todas las filas desde la fila de inicio configurada
- Clasifica cada fila según su tienda destino
- Detecta automáticamente el archivo Excel de la tienda (por nombre, hoja TIENDAS o cache)
- Inserta cada paquete en la primera fila libre del archivo de tienda
- Marca el resultado en la columna B del INGRESO_MASIVO: `LISTO`, `FALTA` o `DUP`
- Actualiza la hoja `FALTA` con todos los registros no procesados
- Genera log persistente en `ultimo_proceso.log`
- Fuerza sincronización OneDrive en cada archivo guardado

**Resultados posibles por fila:**
| Resultado | Significado |
|-----------|-------------|
| `LISTO`   | Paquete insertado correctamente en la tienda |
| `FALTA`   | No se encontró el archivo de la tienda |
| `DUP`     | El ID ya existe en el archivo destino |

---

### ✔ Verificar
Comprueba que los archivos `.xlsx` de la carpeta de tiendas son accesibles y válidos.

**Qué hace:**
- Escanea todos los archivos `.xlsx` en la carpeta configurada
- Verifica que cada archivo tenga una hoja con los encabezados correctos
- Muestra cuántos archivos son válidos, cuántos tienen error y cuántos están bloqueados

---

### 📂 Indexar tiendas
Genera o actualiza el índice de la columna E (nombre de tienda) de todos los archivos.

**Qué hace:**
- Abre cada archivo `.xlsx` de la carpeta de tiendas
- Lee los valores únicos de la columna E (TIENDA)
- Guarda un mapa `nombre_tienda → archivo` en `cache_cole.json`
- Permite que el motor de procesamiento encuentre tiendas aunque el nombre no coincida exactamente con el nombre del archivo

**Cuándo usarlo:** cuando se agregan tiendas nuevas o se cambia el nombre de los archivos.

---

### 📋 Ver FALTA
Muestra todos los registros que no pudieron ser procesados.

**Qué muestra:**
- Fila original en el INGRESO_MASIVO
- Tienda buscada
- ID del paquete
- Motivo exacto del falta
- Comentario

**Controles:**
- Botón **Actualizar** para recargar sin cerrar la ventana
- Contador: `FALTA: X   DUP: X   Total: X`

---

### 🔍 Diagnosticar
Analiza por qué una tienda específica no está siendo encontrada.

**Cómo usarlo:**
1. Escribe el nombre de la tienda tal como aparece en el INGRESO_MASIVO
2. El diagnóstico muestra en 3 niveles:
   - **Nivel 1**: búsqueda exacta por nombre de archivo
   - **Nivel 2**: búsqueda en hoja TIENDAS del INGRESO_MASIVO
   - **Nivel 3**: búsqueda en cache columna E
3. Si no encuentra, muestra nombres similares disponibles
4. Verifica los encabezados del archivo encontrado
5. Concluye con los pasos exactos para solucionar el problema

---

### ⚙ Cambiar rutas
Configura las rutas principales del programa (protegida con contraseña).

**Permite cambiar:**
- Ruta del archivo `INGRESO_MASIVO.xlsx`
- Ruta de la carpeta de tiendas

---

## Motor de procesamiento

### Arquitectura de 2 pasos (optimizada para 1500+ filas)

**Paso 1 — Clasificación previa (sin abrir ningún Excel)**
- Lee todas las filas del INGRESO_MASIVO en memoria en una sola pasada
- Clasifica cada fila hacia su tienda destino usando:
  - Cache de normalización: cada nombre de tienda se normaliza una sola vez
  - Búsqueda en índice de archivos (nombre exacto)
  - Búsqueda en hoja TIENDAS (mapa de variantes)
  - Búsqueda en cache columna E (nombres dentro de los archivos)

**Paso 2 — Procesamiento por tienda**
- Abre cada archivo Excel una sola vez
- Inserta todos los paquetes de esa tienda de una vez
- Guarda y cierra antes de pasar a la siguiente

### Búsqueda de tienda (3 niveles)
1. **Exacto**: nombre normalizado coincide con nombre del archivo
2. **Hoja TIENDAS**: mapa de variantes de nombres en el INGRESO_MASIVO
3. **Cache col E**: nombres que aparecen dentro de los archivos de tienda

### Detección de hoja válida
- Busca dinámicamente la fila del encabezado en las primeras 15 filas
- Valida que existan los encabezados: `F.RECOLECTA`, `TIENDA`, `ID`, `NOMBRE`, `ZONA`, `TELEFONO`, `PRECIO`
- Compatible con hojas donde el encabezado está en fila 5 o fila 6

### Detección de columnas especiales
El motor detecta automáticamente columnas adicionales según el tipo de libro:
| Tipo de libro | Columna detectada | Función |
|---------------|-------------------|---------|
| Metrogalerías | TIPO SERVICIO / SCGE | Inserta NRM/PLUS/ECO |
| Metropolitanos | MUNICIPIO / CIUDAD | Inserta municipio |
| UT Software | PAQUETES | Inserta cantidad de paquetes |
| Con Orden ID | ORDEN ID | Inserta ID de orden |

### Reglas de duplicado
| Situación | Resultado |
|-----------|-----------|
| ID nuevo, sin comentario PAQ | `LISTO` — inserta 1 paquete |
| ID existente, sin comentario PAQ | `DUP` |
| ID nuevo, con PAQ/TYP | `LISTO` — inserta N paquetes juntos |
| ID existente, con PAQ/TYP | `DUP` — no fragmentar |

### Manejo de paquetes múltiples
Detecta en el comentario (columna N) expresiones como:
`2 PAQ`, `PAQ2`, `18PAQ`, `PAQ-5`, `TYP`, `T/P`, `2TYP`, `PAQ:3`, etc.

`T/P` siempre = 2 paquetes. `TYP` sin número = 2 paquetes.

### Formato de fecha
La columna `F.RECOLECTA` se inserta con formato `D-MMM` (ej: `17-mar`) respetando el formato visual de los archivos de tienda.

### Separadores amarillos
Las filas con color amarillo (`#FFFF00`, `#FFFF33`, `#FFCC00`) en la columna D se tratan como separadores de sección y no se sobreescriben.

### Guardado seguro
- Guarda primero en archivo `.tmp`, luego renombra al original
- Preserva el tema de colores del archivo original (no borra paleta de Excel)
- Reintento automático hasta 5 veces con 3 segundos de espera si el archivo está bloqueado
- Detecta archivos abiertos en Excel (`~$archivo`) antes de procesar
- Fuerza sincronización OneDrive actualizando timestamp después de cada guardado

---

## Lógica de negocio

### Tipos de servicio
Resuelve ~1500 variantes de escritura hacia `NRM`, `PLUS` o `ECO`.
Si no reconoce ninguna variante, asigna `NRM` por defecto.

### Normalización de nombres
Convierte nombres a minúsculas, elimina tildes, signos de puntuación y espacios
para comparar nombres de tiendas sin importar cómo estén escritos.

### Precio cero
Si el ID ya existía en el archivo destino (paquete adicional), el precio se
inserta como `0` para evitar duplicar el cobro.

---

## Sistema de actualizaciones

El programa verifica automáticamente si hay actualizaciones disponibles en GitHub
cada vez que se abre.

### Flujo de actualización
1. Al iniciar consulta `version.json` en GitHub (con anti-caché por timestamp)
2. Compara la versión de GitHub con la versión local
3. Si son distintas → muestra banner verde en el sidebar con la versión disponible
4. El usuario hace clic en **"Actualizar ahora"**
5. Descarga cada archivo listado en `version.json → archivos`
6. Hace backup de los archivos anteriores (`.bak`)
7. Actualiza `version.json` local con la nueva versión
8. Reinicia el programa automáticamente

**Nota:** `config_local.py` nunca se actualiza automáticamente para no borrar las rutas configuradas.

### Publicar una actualización
1. Modificar los archivos que cambien
2. Editar `version.json` en GitHub con la nueva versión y la lista de archivos
3. Subir los archivos modificados a GitHub
4. La próxima vez que los usuarios abran el programa verán la actualización

---

## Configuración

### config_local.py
```python
CARPETA_TIENDAS  = r"C:/ruta/a/carpeta/tiendas"   # carpeta con Excel de tiendas
ARCHIVO_INGRESO  = r"C:/ruta/INGRESO_MASIVO.xlsx"  # archivo origen

FILA_INICIO      = 4    # primera fila de datos en INGRESO_MASIVO
FILA_ENCABEZADO  = 5    # fila de encabezados en archivos de tienda
FILA_DATOS_DEST  = 6    # primera fila de datos en archivos de tienda

COL_RESULTADO    = 2    # columna B — donde se escribe LISTO/FALTA/DUP
COL_DATOS_INI    = 4    # columna D — primera columna de datos a copiar
COL_TIENDA       = 5    # columna E — nombre de tienda
COL_ID           = 6    # columna F — ID del paquete
COL_TIPOSERV     = 11   # columna K — tipo de servicio NRM/PLUS/ECO
COL_MUNICIPIO    = 13   # columna M — municipio destino
COL_COMENTARIO   = 14   # columna N — comentario (PAQ, TYP, etc.)

COLORES_AMARILLO = {"FFFF00", "FFFF33", "FFCC00"}  # colores de separador
```

---

## Archivos del proyecto

| Archivo | Descripción |
|---------|-------------|
| `Ingreso_Masivo_XPES.pyw` | Interfaz gráfica principal (Tkinter) |
| `main_local.py` | Motor de procesamiento — lee, clasifica e inserta paquetes |
| `logica_local.py` | Lógica de negocio — validaciones, duplicados, inserción |
| `config_local.py` | Configuración de rutas y columnas |
| `indexar.py` | Generación y lectura del cache de columna E |
| `servicios_variantes.py` | Mapa de ~1500 variantes de escritura de servicios |
| `test.py` | Herramienta de prueba de diagnóstico |
| `version.json` | Versión local instalada |
| `cache_cole.json` | Cache del índice de columna E |
| `ultimo_proceso.log` | Log del último procesamiento ejecutado |
| `xpress_logo.png` | Logo de la aplicación |
| `xpress_icon.png` | Ícono de la aplicación |
