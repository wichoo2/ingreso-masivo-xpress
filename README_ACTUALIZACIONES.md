# Ingreso Masivo — Guía de configuración GitHub + Actualizaciones

## Paso 1: Crear el repositorio en GitHub

1. Ve a https://github.com y crea una cuenta si no tienes
2. Clic en "New repository"
3. Nombre: `ingreso-masivo-xpress`
4. Visibilidad: **Private** (para que solo tú puedas ver el código)
5. Clic en "Create repository"

---

## Paso 2: Subir los archivos al repositorio

Sube estos archivos a la raíz del repo:
- `Ingreso_Masivo_XPES.pyw`
- `main_local.py`
- `logica_local.py`
- `indexar.py`
- `servicios_variantes.py`
- `config_local.py`
- `test.py`
- `version.json`
- `xpress_logo.png`
- `xpress_icon.png`

---

## Paso 3: Configurar tus datos en el .pyw

Abre `Ingreso_Masivo_XPES.pyw` y edita estas líneas al inicio:

```python
GITHUB_USER   = "TU_USUARIO"    # tu usuario de GitHub
GITHUB_REPO   = "ingreso-masivo-xpress"  # nombre del repo
GITHUB_BRANCH = "main"
```

---

## Paso 4: Generar el exe

1. Pon todos los archivos en una carpeta
2. Doble clic en `build_exe.bat`
3. Espera que termine
4. El exe estará en `dist\IngresoMasivo\IngresoMasivo.exe`
5. Distribuye **toda** la carpeta `dist\IngresoMasivo\` (no solo el .exe)

---

## Cómo publicar una actualización

1. Modifica los archivos que quieras cambiar
2. Edita `version.json` cambiando el número de versión:

```json
{
  "version": "3.1",
  "fecha": "20/03/2026",
  "notas": "- Mejora en el diagnóstico\n- Fix de fecha en col D",
  "archivos": [
    "Ingreso_Masivo_XPES.pyw",
    "main_local.py",
    "logica_local.py",
    "indexar.py",
    "servicios_variantes.py",
    "test.py"
  ]
}
```

3. Sube los archivos modificados + `version.json` a GitHub
4. La próxima vez que los usuarios abran el programa verán el banner verde de actualización

---

## Cómo funciona la actualización en el programa

1. Al abrir el programa consulta `version.json` de GitHub
2. Si la versión es diferente a la instalada → aparece banner verde en el sidebar
3. El usuario hace clic en "Actualizar ahora"
4. Se muestra un diálogo con las novedades
5. Al confirmar, descarga cada archivo nuevo desde GitHub
6. Hace backup de los archivos anteriores (`.bak`)
7. Reinicia el programa automáticamente

**Nota:** `config_local.py` NUNCA se actualiza automáticamente
para no borrar las rutas configuradas por el usuario.
