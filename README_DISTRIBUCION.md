# Cyclope — Distribución como ejecutable (.exe)

Esta guía explica cómo convertir el pipeline en **un único `Cyclope.exe`** que tu
equipo pueda usar sin tener Python ni VS Code.

## Archivos del proyecto

| Archivo | Para qué sirve |
|---|---|
| `cyclope_core.py` | Toda la lógica (extracción + pre-llenado + folders). No se toca. |
| `cyclope_gui.py` | La interfaz gráfica (seleccionar PDF → Ejecutar). |
| `requirements.txt` | Lista de librerías que se instalan al construir. |
| `build_exe.bat` | Script que **tú** corres una vez para generar el `.exe`. |
| `Cyclope_PreFill.ipynb` | El notebook (sigue sirviendo para ti en VS Code). |

> Las 5 librerías son: `pandas`, `openpyxl`, `pdfplumber`, `pypdf`, `lxml`.

---

## Cómo construir el .exe (lo haces TÚ, una sola vez, en Windows)

1. **Reúne en una misma carpeta** (en tu máquina con Windows):
   - `cyclope_core.py`
   - `cyclope_gui.py`
   - `requirements.txt`
   - `build_exe.bat`
   - La carpeta **`Requirements`** con las 7 plantillas (`.docx`, `.pdf`, `.xlsx`).

   > Es importante que la carpeta `Requirements` esté ahí: sus plantillas quedan
   > **empaquetadas dentro del .exe**, así tus compañeros no necesitan tenerlas.

2. **Doble clic en `build_exe.bat`** (o ábrelo desde la terminal).
   - Instala las dependencias + PyInstaller.
   - Empaqueta todo. Tarda 1–3 minutos.

3. Cuando termine, el ejecutable estará en:
   ```
   dist\Cyclope.exe
   ```

4. **Reparte ese único `Cyclope.exe`** a tu equipo (correo, carpeta de red, etc.).

---

## Cómo lo usan tus compañeros (sin instalar nada)

1. Doble clic en `Cyclope.exe`.
2. **Seleccionar PDF(s)** de la Cámara de Comercio.
3. (Opcional) cambiar la carpeta de salida (por defecto: `Escritorio\Cyclope_Output`).
4. (Opcional) activar la **Sección 3** para crear las carpetas de contraparte en la red.
5. Pulsar **▶ Ejecutar**. El progreso aparece en pantalla.
6. Botón **"Abrir carpeta de resultados"** para ver los documentos generados.

---

## Solución de problemas

- **El .exe no abre / se cierra al instante**
  Edita `build_exe.bat` y cambia `--windowed` por `--console`, reconstruye, y
  ejecútalo desde una terminal para ver el mensaje de error.

- **El antivirus marca el .exe**
  Es común con ejecutables de PyInstaller (falso positivo). Si tu política lo
  permite, agrégalo a la lista de confianza; si no, habrá que firmarlo con el
  certificado corporativo (consultar con tu área de seguridad/IT).

- **Cambiaron una plantilla**
  Actualiza el archivo dentro de la carpeta `Requirements` y vuelve a correr
  `build_exe.bat` para regenerar el `.exe`.

- **Probar sin construir el .exe** (en tu máquina, con Python):
  ```bat
  pip install -r requirements.txt
  python cyclope_gui.py
  ```
  (Necesitas la carpeta `Requirements` junto a los `.py`.)
