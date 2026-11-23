"""
Escaner PDF a doble cara (manual) para Windows 11
---------------------------------------------------
Utilidad con interfaz grafica que permite escanear documentos a doble
cara con un escaner/multifuncion que NO tiene alimentador dúplex
automático: el programa escanea una pagina, te pide que le des la
vuelta a la hoja (o coloques la siguiente), y repite el proceso hasta
que pulses "Finalizar", momento en el que une todas las paginas
escaneadas en un unico PDF y te deja guardarlo donde quieras.

Requisitos (se instalan una sola vez):
    pip install pywin32 pillow

Como se usa:
    1. Conecta e instala tu multifuncion en Windows (ya lo tienes).
    2. Ejecuta:  python scan_doble_cara.py
    3. Si Windows detecta varios escaneres te dejara elegir cual usar.
    4. Pulsa "Escanear pagina", coloca el papel y confirma.
    5. Cuando termines todas las paginas, pulsa "Finalizar y guardar PDF".

Notas tecnicas:
    - Usa la API WIA (Windows Image Acquisition) a traves de pywin32,
      que es la forma estandar de hablar con escaneres/MFPs en Windows
      sin depender del fabricante.
    - Cada pagina se captura en color a 300 DPI (puedes cambiarlo en
      RESOLUTION_DPI y COLOR_MODE mas abajo).
    - El "doble cara" es manual: como la mayoria de multifuncion domestica
      no tiene ADF dúplex, el propio flujo del programa (escanear ->
      pedir que gires la hoja -> escanear) hace el trabajo de un dúplex.
      Si tu equipo SI tiene alimentador automatico a doble cara, puedes
      simplemente escanear cada hoja ya impresa por ambas caras: el WIA
      la devolvera como una sola imagen por cara igualmente, asi que el
      flujo (pagina a pagina) sigue siendo válido.
    - En la ventana puedes elegir el origen del papel: "Cristal (flatbed)"
      o "Alimentador automatico (ADF)". Si tu multifuncion tiene ADF y lo
      seleccionas, cada "Escanear pagina" coge la hoja del alimentador en
      lugar del cristal. No todos los drivers WIA exponen esta opcion por
      separado; si el tuyo no la soporta, el programa te avisa y sigue
      usando la configuracion por defecto del escaner.

Modo automatico (boton "Escaneo automatico doble cara (ADF)"):
    1. Colocas TODO el taco de hojas en el ADF y el programa escanea solas,
       una tras otra, todas las caras delanteras (1, 2, 3, ... en orden),
       hasta que el alimentador se queda sin papel.
    2. Te pide que cojas el taco ENTERO tal como ha salido y le des la
       vuelta de un solo golpe (no hoja por hoja) y lo vuelvas a colocar
       en el ADF. El programa escanea de nuevo automaticamente: esta
       segunda pasada captura las caras traseras, pero en orden inverso
       (empieza por la ultima hoja del documento y termina por la
       primera), que es justo lo que ocurre al voltear el taco completo.
    3. Al terminar, el programa reordena internamente las caras traseras
       (invierte esa segunda lista) y las intercala con las delanteras en
       el orden correcto: delantera 1, trasera 1, delantera 2, trasera 2,
       etc., reconstruyendo el documento original con un solo boton.
"""

import io
import os
import sys
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import win32com.client
except ImportError:
    print("Falta la libreria pywin32. Instalala con:  pip install pywin32")
    sys.exit(1)

try:
    from PIL import Image, ImageTk
except ImportError:
    print("Falta la libreria Pillow. Instalala con:  pip install pillow")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Configuracion de escaneo
# ---------------------------------------------------------------------------
RESOLUTION_DPI = 300
# 1 = blanco y negro, 2 = gris, 4 = color (valores WIA_DPS_CUR_INTENT / ids)
# Usamos las propiedades estandar de WIA (IDs numericos documentados por Microsoft)
WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"

# IDs de propiedades WIA 2.0 para el item de escaneo
PROP_DPI_X = 6147
PROP_DPI_Y = 6148
PROP_DATA_TYPE = 4104  # 0=B/N, 1=gris, 2=color, 3=auto
PROP_BRIGHTNESS = 6154
PROP_CONTRAST = 6155

DATA_TYPE_COLOR = 2

# Propiedad WIA 2.0 para elegir el origen del papel (a nivel de dispositivo)
PROP_DOCUMENT_HANDLING_SELECT = 3088
DOC_HANDLING_FEEDER = 1   # ADF / alimentador automatico
DOC_HANDLING_FLATBED = 2  # Cristal


def set_item_property(item, prop_id, value):
    """Fija una propiedad WIA si el escaner la admite; ignora si no existe."""
    try:
        for p in item.Properties:
            if p.PropertyID == prop_id:
                p.Value = value
                return True
    except Exception:
        pass
    return False


class Escaner:
    """Envoltorio sencillo sobre un dispositivo WIA de tipo escaner."""

    def __init__(self):
        self.manager = win32com.client.Dispatch("WIA.DeviceManager")
        self.device = None
        self.item = None

    def listar_escaneres(self):
        dispositivos = []
        for info in self.manager.DeviceInfos:
            # Type == 1 -> StiDeviceTypeScanner
            try:
                tipo = info.Type
            except Exception:
                tipo = None
            if tipo == 1:
                nombre = info.Properties("Name").Value
                dispositivos.append((nombre, info))
        return dispositivos

    def conectar(self, device_info):
        self.device = device_info.Connect()
        # El item 1 suele ser la bandeja/cristal de escaneo
        self.item = self.device.Items[1]
        set_item_property(self.item, PROP_DPI_X, RESOLUTION_DPI)
        set_item_property(self.item, PROP_DPI_Y, RESOLUTION_DPI)
        set_item_property(self.item, PROP_DATA_TYPE, DATA_TYPE_COLOR)

    def set_source(self, usar_adf):
        """Intenta cambiar el origen del papel a ADF o a cristal.

        Devuelve True si se ha podido aplicar el cambio de alguna forma,
        False si el escaner no expone esa opcion y se queda con su
        configuracion por defecto.
        """
        objetivo = "feeder" if usar_adf else "flatbed"
        encontrado = False

        # 1) Algunos escaneres exponen un item distinto por cada origen
        #    (p.ej. "Flatbed" y "Feeder" dentro de device.Items).
        try:
            for it in self.device.Items:
                nombre_item = ""
                for prop_name in ("Item Name", "Full Item Name", "Name"):
                    try:
                        nombre_item = str(it.Properties(prop_name).Value or "")
                        if nombre_item:
                            break
                    except Exception:
                        continue
                if objetivo in nombre_item.lower():
                    self.item = it
                    encontrado = True
                    break
        except Exception:
            pass

        # 2) Si no hay items separados, probamos la propiedad de dispositivo
        #    WIA_DPS_DOCUMENT_HANDLING_SELECT.
        if not encontrado:
            valor = DOC_HANDLING_FEEDER if usar_adf else DOC_HANDLING_FLATBED
            try:
                for p in self.device.Properties:
                    if p.PropertyID == PROP_DOCUMENT_HANDLING_SELECT:
                        p.Value = valor
                        encontrado = True
                        break
            except Exception:
                pass

        # Volvemos a aplicar resolucion/color al item que vayamos a usar
        set_item_property(self.item, PROP_DPI_X, RESOLUTION_DPI)
        set_item_property(self.item, PROP_DPI_Y, RESOLUTION_DPI)
        set_item_property(self.item, PROP_DATA_TYPE, DATA_TYPE_COLOR)

        return encontrado

    def escanear_pagina(self):
        """Lanza una captura y devuelve una imagen PIL."""
        image_file = self.item.Transfer(WIA_FORMAT_BMP)
        # image_file es un objeto WIA.ImageFile; lo pasamos a bytes
        tmp_path = os.path.join(tempfile.gettempdir(), "wia_tmp_page.bmp")
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        image_file.SaveFile(tmp_path)
        img = Image.open(tmp_path).convert("RGB")
        img.load()
        os.remove(tmp_path)
        return img

    @staticmethod
    def es_adf_vacio(exc):
        """Detecta si una excepcion de Transfer() es simplemente que el
        alimentador se ha quedado sin papel (fin normal de la pila), y no
        un fallo real (atasco, desconexion, etc.)."""
        texto = str(exc).lower()
        marcadores = (
            "80210003",          # WIA_ERROR_PAPER_EMPTY, en hexadecimal
            "-2145320957",       # el mismo codigo, en decimal con signo
            "2145320957",        # por si aparece sin el signo
            "paper_empty",
            "paper empty",
            "no hay papel",
            "out of paper",
            "sin papel",
            "no hay ninguno en el alimentador",
            "no hay ningun documento",
            "no hay ningún documento",
        )
        return any(m in texto for m in marcadores)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Escaner PDF a doble cara")
        self.root.geometry("480x470")
        self.root.resizable(False, False)

        self.escaner = Escaner()
        self.paginas = []  # lista de imagenes PIL en orden

        self._elegir_escaner()
        self._construir_ui()

    # -- Seleccion de escaner -------------------------------------------------
    def _elegir_escaner(self):
        dispositivos = self.escaner.listar_escaneres()
        if not dispositivos:
            messagebox.showerror(
                "No se encontro ningun escaner",
                "Windows no detecta ningun escaner/multifuncion instalado.\n"
                "Comprueba que el equipo esta encendido, conectado y que el "
                "driver/WIA esta instalado, y vuelve a intentarlo.",
            )
            sys.exit(1)

        if len(dispositivos) == 1:
            nombre, info = dispositivos[0]
            self.escaner.conectar(info)
            return

        # Varios escaneres: ventana sencilla de seleccion
        seleccion = tk.StringVar(value=dispositivos[0][0])
        win = tk.Toplevel(self.root)
        win.title("Elige el escaner")
        win.grab_set()
        tk.Label(win, text="Se han encontrado varios escaneres.\nElige cual quieres usar:").pack(padx=20, pady=10)
        combo = ttk.Combobox(win, textvariable=seleccion, state="readonly",
                              values=[n for n, _ in dispositivos], width=40)
        combo.pack(padx=20, pady=5)

        def confirmar():
            nombre_elegido = seleccion.get()
            info_elegido = next(info for n, info in dispositivos if n == nombre_elegido)
            self.escaner.conectar(info_elegido)
            win.destroy()

        tk.Button(win, text="Aceptar", command=confirmar).pack(pady=10)
        self.root.wait_window(win)

    # -- Interfaz --------------------------------------------------------------
    def _construir_ui(self):
        marco = tk.Frame(self.root, padx=20, pady=20)
        marco.pack(fill="both", expand=True)

        origen_frame = tk.LabelFrame(marco, text="Origen del papel", padx=10, pady=5)
        origen_frame.pack(fill="x", pady=(0, 10))

        self.origen_var = tk.StringVar(value="flatbed")
        tk.Radiobutton(
            origen_frame, text="Cristal (flatbed)", variable=self.origen_var,
            value="flatbed", command=self._cambiar_origen,
        ).pack(side="left", padx=(0, 15))
        tk.Radiobutton(
            origen_frame, text="Alimentador automatico (ADF)", variable=self.origen_var,
            value="adf", command=self._cambiar_origen,
        ).pack(side="left")

        self.lbl_estado = tk.Label(
            marco,
            text="Pulsa 'Escanear pagina' para capturar la primera hoja.",
            wraplength=400, justify="left", font=("Segoe UI", 10),
        )
        self.lbl_estado.pack(pady=(0, 10))

        self.lbl_contador = tk.Label(marco, text="Paginas escaneadas: 0",
                                      font=("Segoe UI", 11, "bold"))
        self.lbl_contador.pack(pady=(0, 10))

        self.lbl_preview = tk.Label(marco, borderwidth=1, relief="solid",
                                     width=40, height=10, bg="#f0f0f0",
                                     text="(vista previa)")
        self.lbl_preview.pack(pady=(0, 15))

        botones = tk.Frame(marco)
        botones.pack(fill="x")

        self.btn_escanear = tk.Button(
            botones, text="Escanear pagina", height=2,
            command=self.on_escanear,
        )
        self.btn_escanear.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self.btn_finalizar = tk.Button(
            botones, text="Finalizar y guardar PDF", height=2,
            command=self.on_finalizar, state="disabled",
        )
        self.btn_finalizar.pack(side="left", expand=True, fill="x", padx=(5, 0))

        self.btn_adf_auto = tk.Button(
            marco, text="Escaneo automatico doble cara (ADF)", height=2,
            command=self.on_adf_automatico,
        )
        self.btn_adf_auto.pack(fill="x", pady=(10, 0))

    def _cambiar_origen(self):
        usar_adf = self.origen_var.get() == "adf"
        try:
            soportado = self.escaner.set_source(usar_adf)
        except Exception as exc:
            messagebox.showerror(
                "Error al cambiar el origen",
                f"No se ha podido cambiar el origen del papel:\n{exc}",
            )
            return
        if usar_adf and not soportado:
            messagebox.showwarning(
                "ADF no detectado por separado",
                "Este escaner no expone el alimentador automatico como una "
                "opcion independiente del cristal. Se mantiene la "
                "configuracion por defecto del dispositivo; si tu "
                "multifuncion tiene ADF, puede que lo use igualmente al "
                "detectar papel en el, o puede que tengas que activarlo "
                "desde el propio panel del equipo.",
            )

    # -- Logica de escaneo -------------------------------------------------
    def on_escanear(self):
        numero = len(self.paginas) + 1
        confirmar = messagebox.askokcancel(
            "Preparar pagina",
            f"Coloca la pagina {numero} en el escaner y pulsa Aceptar "
            f"para escanearla.",
        )
        if not confirmar:
            return

        self.btn_escanear.config(state="disabled")
        self.root.update_idletasks()
        try:
            imagen = self.escaner.escanear_pagina()
        except Exception as exc:
            messagebox.showerror(
                "Error al escanear",
                f"No se ha podido completar el escaneo:\n{exc}",
            )
            self.btn_escanear.config(state="normal")
            return

        self.paginas.append(imagen)
        self._actualizar_preview(imagen)
        self.lbl_contador.config(text=f"Paginas escaneadas: {len(self.paginas)}")
        self.lbl_estado.config(
            text=(
                f"Pagina {numero} escaneada correctamente.\n"
                "Dale la vuelta a la hoja (o coloca la siguiente) y pulsa "
                "'Escanear pagina' para continuar, o pulsa 'Finalizar y "
                "guardar PDF' si ya has terminado."
            )
        )
        self.btn_escanear.config(state="normal", text="Escanear pagina siguiente")
        self.btn_finalizar.config(state="normal")

    # -- Modo automatico: ADF + dos pasadas (delanteras / traseras) --------
    def on_adf_automatico(self):
        if self.paginas:
            if not messagebox.askyesno(
                "Sustituir documento actual",
                "El escaneo automatico va a reemplazar las paginas que ya "
                "tienes capturadas en este documento. ¿Quieres continuar?",
            ):
                return

        # Forzamos el origen a ADF (si el escaner lo permite)
        self.origen_var.set("adf")
        try:
            self.escaner.set_source(True)
        except Exception as exc:
            messagebox.showerror(
                "Error al preparar el ADF",
                f"No se ha podido preparar el alimentador automatico:\n{exc}",
            )
            return

        if not messagebox.askokcancel(
            "Paso 1: caras delanteras",
            "Coloca TODO el taco de hojas en el alimentador (ADF), en el "
            "orden correcto, con la pagina 1 encima (o segun indique tu "
            "equipo para que se lea en orden).\n\n"
            "Pulsa Aceptar para escanear automaticamente todas las caras "
            "delanteras, una tras otra, hasta que el alimentador se quede "
            "vacio.",
        ):
            return

        self._bloquear_botones(True)
        try:
            delanteras = self._escanear_pila_automatica("Caras delanteras")
        finally:
            self._bloquear_botones(False)

        if not delanteras:
            messagebox.showwarning(
                "Sin paginas",
                "No se ha detectado papel en el alimentador. Revisa que el "
                "taco este bien colocado e intentalo de nuevo.",
            )
            return

        if not messagebox.askokcancel(
            "Paso 2: caras traseras",
            f"Se han escaneado {len(delanteras)} cara(s) delantera(s).\n\n"
            "Ahora coge el taco ENTERO de hojas tal y como ha salido, "
            "daselo la vuelta de un solo golpe (gira el taco completo, no "
            "hoja por hoja) y colocalo de nuevo en el alimentador.\n\n"
            "Al escanear otra vez, el programa leera las caras traseras "
            "empezando por la ULTIMA hoja del documento y terminando por "
            "la primera; el propio programa se encarga de reordenarlo todo "
            "despues.\n\n"
            "Pulsa Aceptar cuando el taco este listo para la segunda "
            "pasada.",
        ):
            # El usuario prefiere quedarse solo con las caras delanteras
            self.paginas = delanteras
            self._refrescar_tras_escaneo(
                "Documento compuesto solo con las caras delanteras "
                f"({len(delanteras)} pagina(s))."
            )
            return

        self._bloquear_botones(True)
        try:
            traseras_bruto = self._escanear_pila_automatica("Caras traseras")
        finally:
            self._bloquear_botones(False)

        # Pasada 2 sale en orden inverso (ultima hoja -> primera), asi que
        # al darle la vuelta la volvemos a poner en orden natural.
        traseras = list(reversed(traseras_bruto))

        paginas_final = []
        total = max(len(delanteras), len(traseras))
        for i in range(total):
            if i < len(delanteras):
                paginas_final.append(delanteras[i])
            if i < len(traseras):
                paginas_final.append(traseras[i])

        self.paginas = paginas_final

        if len(delanteras) != len(traseras):
            messagebox.showwarning(
                "El numero de caras no coincide",
                f"Se han escaneado {len(delanteras)} cara(s) delantera(s) y "
                f"{len(traseras)} trasera(s). El documento se ha compuesto "
                "igualmente intercalando las paginas disponibles, pero "
                "revisa el PDF resultante por si falta alguna cara o el "
                "orden no es exacto (por ejemplo, si la ultima hoja es a "
                "una sola cara).",
            )

        self._refrescar_tras_escaneo(
            f"Documento compuesto automaticamente: {len(delanteras)} "
            f"cara(s) delantera(s) + {len(traseras)} trasera(s) = "
            f"{len(self.paginas)} pagina(s) en total."
        )

    def _escanear_pila_automatica(self, etiqueta):
        """Escanea paginas en bucle hasta que el ADF se queda sin papel.
        Actualiza la vista previa y el contador en vivo."""
        capturadas = []
        self.lbl_estado.config(text=f"{etiqueta}: escaneando pagina 1...")
        self.root.update()
        while True:
            try:
                imagen = self.escaner.escanear_pagina()
            except Exception as exc:
                if Escaner.es_adf_vacio(exc):
                    break
                messagebox.showerror(
                    "Error durante el escaneo automatico",
                    f"Se ha producido un error al escanear:\n{exc}",
                )
                break
            capturadas.append(imagen)
            self._actualizar_preview(imagen)
            self.lbl_contador.config(
                text=f"Paginas escaneadas: {len(self.paginas) + len(capturadas)}"
            )
            self.lbl_estado.config(
                text=f"{etiqueta}: escaneando pagina {len(capturadas) + 1}..."
            )
            self.root.update()
        return capturadas

    def _bloquear_botones(self, bloqueado):
        estado = "disabled" if bloqueado else "normal"
        self.btn_escanear.config(state=estado)
        self.btn_adf_auto.config(state=estado)
        if bloqueado:
            self.btn_finalizar.config(state="disabled")
        else:
            self.btn_finalizar.config(state="normal" if self.paginas else "disabled")
        self.root.update_idletasks()

    def _refrescar_tras_escaneo(self, mensaje):
        if self.paginas:
            self._actualizar_preview(self.paginas[-1])
        self.lbl_contador.config(text=f"Paginas escaneadas: {len(self.paginas)}")
        self.lbl_estado.config(text=mensaje)
        self.btn_escanear.config(text="Escanear pagina siguiente")
        self.btn_finalizar.config(state="normal" if self.paginas else "disabled")

    def _actualizar_preview(self, imagen):
        miniatura = imagen.copy()
        miniatura.thumbnail((280, 200))
        self._foto = ImageTk.PhotoImage(miniatura)  # referencia viva
        self.lbl_preview.config(image=self._foto, text="")

    def on_finalizar(self):
        if not self.paginas:
            messagebox.showwarning("Sin paginas", "Todavia no has escaneado ninguna pagina.")
            return

        destino = filedialog.asksaveasfilename(
            title="Guardar PDF como",
            defaultextension=".pdf",
            filetypes=[("Documento PDF", "*.pdf")],
            initialfile="documento_escaneado.pdf",
        )
        if not destino:
            return  # el usuario ha cancelado el guardado

        try:
            primera, *resto = self.paginas
            primera.save(destino, save_all=True, append_images=resto, format="PDF")
        except Exception as exc:
            messagebox.showerror("Error al guardar", f"No se ha podido crear el PDF:\n{exc}")
            return

        messagebox.showinfo(
            "PDF creado",
            f"Se ha guardado el PDF con {len(self.paginas)} pagina(s) en:\n{destino}",
        )

        # Reiniciar estado para poder escanear otro documento sin reabrir el programa
        self.paginas = []
        self.lbl_contador.config(text="Paginas escaneadas: 0")
        self.lbl_preview.config(image="", text="(vista previa)")
        self.lbl_estado.config(text="Pulsa 'Escanear pagina' para capturar la primera hoja.")
        self.btn_escanear.config(text="Escanear pagina")
        self.btn_finalizar.config(state="disabled")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
