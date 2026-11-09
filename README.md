# 🗓️ Extractor de Horarios (Python + OCR)

Herramienta que convierte **capturas de pantalla de horarios universitarios** en una **tabla de Excel** lista para analizar.

Tomaba capturas de la [página de horarios de la FES Aragón (UNAM)](https://www.aragon.unam.mx/horarios/horarios/horarios/) y este programa leía cada una para armar `Horarios.xlsx`. Con esa tabla se podía trabajar en otras plataformas o en informes, sin copiar los datos a mano.

Funciona **100 % en local**: usa OCR (Tesseract) y análisis de colores de la imagen, sin servicios en línea ni IA.

## ⚙️ Cómo funciona

1. Lee todas las imágenes de la carpeta `screenshots/`.
2. Encuentra la **línea dorada** de la parte de arriba y lee el nombre de la **carrera** que está encima.
3. Busca el **semestre** en el texto de la imagen y distingue si es sistema escolarizado (`ESC`) o abierto (`SUA`).
4. Detecta las **columnas** a partir del encabezado azul de la tabla.
5. Detecta las **filas** a partir de las líneas oscuras que separan cada renglón.
6. Recorta cada celda, la amplía 3 veces, la pasa a blanco y negro y la lee con OCR.
7. Limpia los datos: deja solo horas (`07:00`) en los días y solo números en clave y cupo.
8. Guarda todo en `Horarios.xlsx`, con encabezado de color, bordes y la primera fila fija.

### Columnas del Excel

`Carrera` · `Acronimo` · `Sistema` · `Semestre` · `Profesor` · `CVE` · `Asignatura` · `Grupo` · `Tipo` · `Lun` · `Mar` · `Mie` · `Jue` · `Vie` · `Sab` · `Salon` · `Cupo`

## ▶️ Cómo ejecutarlo

1. Instala [Python 3](https://www.python.org/downloads/).
2. Instala **Tesseract OCR**:
   - Windows: [instalador de UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
   - Mac: `brew install tesseract`
3. Clona el repositorio e instala las librerías:

```bash
git clone https://github.com/crisag01/Extractor_Horarios.git
cd Extractor_Horarios
pip install -r requirements.txt
```

4. Pon tus capturas de horarios en la carpeta `screenshots/` (`.png`, `.jpg`, `.jpeg`, `.webp` o `.bmp`).
5. Ejecuta:

```bash
python extraer_horarios.py
```

> En Windows, si aparece un error de que no encuentra Tesseract, quita el `#` de la línea `pytesseract.pytesseract.tesseract_cmd = ...` en el código y revisa que la ruta sea la de tu instalación.

### Ejemplo de salida

```text
Carpeta: screenshots
Imágenes encontradas: 3

  [1/3] Procesando: horario_1.png ... OK (12 fila(s))
  [2/3] Procesando: horario_2.png ... OK (9 fila(s))
  [3/3] Procesando: horario_3.png ... OK (14 fila(s))

Archivo guardado: Horarios.xlsx
Total de filas: 35
```

> Las capturas y el Excel generado **no se suben al repositorio**, porque contienen nombres de profesores. Así lo indica el `.gitignore`.

## 👤 Autor

**Cristian Armenta**, ingeniero industrial y analista de datos.
Más proyectos en [crisag.com](https://crisag.com).
