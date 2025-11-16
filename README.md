# 🖥️ Control de Detectores HPGe – GammaVision  
**Automatización de control diario, estabilidad y resolución (FWHM/FWTM)**  
**Jobs & Scripts: Controlcongraf.job, ControlSINgraf.job, graphtest, datatest**

Este sistema automatiza el control de funcionamiento de los **detectores HPGe** operados mediante **GammaVision**, procesando el archivo de control generado en cada medición, verificando centroides, FWHM/FWTM, estabilidad, y generando reportes y gráficos.

Incluye:

- Lectura del archivo **controlgeneral.rpt/.txt**
- Evaluación de tolerancias (centroides y resolución)
- Detección de descalibración
- Escritura automática en **RegistroDetX.xlsx**
- Generación de gráficos de estabilidad
- Alertas guiadas para el usuario
- Copias de respaldo automáticas

---

# 📌 Uso desde GammaVision

## 1. Abrir Job Control
En el detector correspondiente o en Buffer:

**Services → Job Control**

## 2. Ejecutar el Job
Seleccionar:
```bash
C:\ProgramControl\Controlcongraf.job
```
y ejecutarlo.

> ✔ Si no se desean gráficos, usar:  
> `C:\ProgramControl\ControlSINgraf.job`

---

# 📌 Verificaciones posteriores

Tras ejecutar el job, ir a:
```bash
C:\pathcontrol\DetX\
```
donde **X** es el detector (5, 7, etc.)
y donde **pathcontrol** es el path especificado en PathDetX

### Deben generarse:

### ✔ 1. RegistroDetX.xlsx  
Contiene una pestaña por energía del Eu-152, con filas nuevas por fecha:

- Fecha  
- Centroid  
- FWHM  
- FWTM  
- Estado (ok / descalibrado)

### ✔ 2. Gráficos PNG  
Generados automáticamente con `graphtest`:
```bash
121.78.png
244.70.png
344.28.png
```
---

# ⚠ Alertas posibles

### 🔴 **“CALIBRAR y volver a correr el job”**
Centroid fuera de ±0.3 keV.  
No se generan gráficos.

### 🟡 **“FWTM/FWHM fuera de rango”**
Resolución fuera del límite del detector.  
Los gráficos se generan igual.

### 🔵 **“Cerrar el Excel”**
Debe cerrarse **RegistroDetX.xlsx**.  
Si persiste → contactar.

---

# 📁 Estructura de archivos

| Ruta | Descripción |
|------|-------------|
| `C:\GammaControl\controlgeneral.txt` | Archivo generado por GammaVision |
| `C:\ProgramControl\Controlcongraf.job` | Job que genera gráficos |
| `C:\ProgramControl\ControlSINgraf.job` | Job sin gráficos |
| `C:\ProgramInfodet\PathDetX` | Rutas de salida del detector |
| `C:\ProgramInfodet\LimDetX` | Límites de FWHM/FWTM por energía |
| `C:\Librerias\EuControlROI.Lib` | Librería de picos de Eu-152 |
| `reporteatextof.bat` | Convierte RPT → TXT |
| `datatest4.exe` | Escribe los datos en Excel |
| `graphtest5.exe` | Genera gráficos PNG |

---

# 🧠 Lógica del sistema

## 1. Identificación automática del detector
A partir del archivo `controlgeneral.txt`:

- Detector 5  
- Detector 7  
- etc.

## 2. Lectura del archivo TXT
Se extraen:

- Energía  
- CENTROID  
- FWHM  
- FWTM  
- Fecha

## 3. Tolerancias
### ✔ Centroid  
±0.3 keV  
Si falla → “descalibrado” y alerta roja.

### ✔ FWHM / FWTM  
Comparación contra `LimDetX`.

## 4. Escritura en Excel
El script:

- Identifica la siguiente fila libre
- Escribe la nueva medición
- Marca “ok” o “descalibrado”
- Copia el archivo a:
  - `copy_output_file`
  - `backup_file`

## 5. Gráficos
Cada energía produce un PNG:

- Centroid vs Fecha  
- FWHM vs Fecha  
- FWTM vs Fecha  

Con rangos, tolerancias y colores.

---

# 🧰 Principales funciones del código

- `get_detector_number()`  
- `load_detector_config()`  
- `load_detector_pico()`  
- `dic_rango_centro()`  
- `rango_centro()`  
- `check_fwhm_fwtm()`  
- `append_to_worksheet()`  
- `generate_alert()`  
- **graphtest5.py:** lectura + generación de gráficos

---

# 🔧 Dependencias

- Python 2.7  
- pandas  
- numpy  
- openpyxl  
- matplotlib  
- ctypes  
