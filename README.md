# Escáner PDF a doble cara (Windows 11)

Utilidad de escritorio para Windows con interfaz gráfica que permite escanear
documentos **a doble cara** desde cualquier impresora multifunción instalada,
y componer el resultado en un único archivo PDF.

No depende del software propio del fabricante: habla directamente con el
escáner a través de **WIA** (Windows Image Acquisition), el estándar de
Windows para dispositivos de digitalización.

## Características

- **Modo manual (página a página):** escaneas una hoja, el programa te pide
  que le des la vuelta o coloques la siguiente, y repites hasta pulsar
  "Finalizar". Válido tanto para cristal (flatbed) como para ADF hoja a hoja.
- **Selector de origen del papel:** Cristal (flatbed) o Alimentador
  automático (ADF), si tu equipo lo soporta.
- **Modo automático con ADF (doble cara con una sola pila):**
  1. Colocas todo el taco de hojas en el ADF y el programa escanea solo,
     una tras otra, todas las caras delanteras, hasta vaciar el alimentador.
  2. Te pide que cojas el taco completo tal cual ha salido, le des la
     vuelta de un solo golpe (no hoja por hoja) y lo vuelvas a colocar en
     el ADF.
  3. Escanea de nuevo automáticamente las caras traseras (que, al haber
     volteado el taco entero, salen en orden inverso).
  4. El programa reordena internamente esa segunda pasada y la intercala
     con la primera (delantera 1, trasera 1, delantera 2, trasera 2…),
     reconstruyendo el documento original con un solo botón.
- Vista previa en miniatura y contador de páginas en tiempo real.
- Exporta todo a un único PDF, eligiendo dónde guardarlo.

## Requisitos

- Windows 11 (también debería funcionar en Windows 10) con el escáner o
  multifunción ya instalado y funcionando desde el propio sistema.
- Python 3.9 o superior.
- Dependencias de Python:
  ```
  pip install pywin32 pillow
  ```

## Uso

```
python scan_doble_cara.py
```

1. Si Windows detecta varios escáneres, el programa te deja elegir cuál
   usar.
2. Elige el origen del papel (Cristal o ADF) si quieres forzarlo.
3. Para escaneo manual: pulsa **"Escanear página"**, coloca la hoja y
   confirma. Repite dándole la vuelta o colocando la siguiente página.
4. Para escaneo automático a doble cara: pulsa **"Escaneo automático doble
   cara (ADF)"** y sigue las indicaciones en pantalla (pasada de
   delanteras → voltear el taco → pasada de traseras).
5. Pulsa **"Finalizar y guardar PDF"** en cualquier momento para unir las
   páginas capturadas y guardarlas como un `.pdf`.

## Generar un ejecutable (.exe)

El `.exe` debe compilarse en una máquina Windows (PyInstaller no permite
compilar un `.exe` de Windows desde Linux/macOS, y `pywin32` solo funciona
en Windows).

```powershell
pip install pywin32 pillow pyinstaller
python -m PyInstaller --onefile --noconsole --name "EscanerDobleCara" scan_doble_cara.py
```

El ejecutable se genera en `dist\EscanerDobleCara.exe`. Puedes copiar solo
ese archivo donde quieras; no necesita Python instalado en el equipo de
destino (sí necesita el runtime de Visual C++, normalmente ya presente en
Windows 11).

> Si `pyinstaller` da "no se reconoce como un cmdlet", usa `python -m
> PyInstaller ...` como en el comando de arriba, que evita problemas de
> PATH.

## Notas técnicas

- Cada página se captura en color a 300 DPI (ajustable en `RESOLUTION_DPI`
  y `DATA_TYPE_COLOR` dentro del script).
- El origen ADF/Cristal se selecciona probando primero items WIA separados
  por nombre ("Feeder"/"Flatbed") y, si el driver no los expone así, la
  propiedad estándar `WIA_DPS_DOCUMENT_HANDLING_SELECT`. No todos los
  drivers soportan este cambio por software; el programa avisa si no lo
  detecta y sigue con la configuración por defecto del dispositivo.
- El fin de la pila en el ADF se detecta capturando el error WIA
  `WIA_ERROR_PAPER_EMPTY` (código `0x80210003` / `-2145320957`), tanto en
  formato hexadecimal como decimal, además del texto del mensaje en
  español.
- Si el número de caras delanteras y traseras no coincide (por ejemplo,
  la última hoja es a una sola cara), el programa avisa pero compone el
  PDF igualmente intercalando lo que tiene, para que lo revises.

## Licencia

Este proyecto está licenciado bajo la **GNU General Public License v3.0**
(GPL-3.0). Consulta el archivo `LICENSE` del repositorio para el texto
completo.
