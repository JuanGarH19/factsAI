"""
extractor.py  ·  FacturaIA v2
------------------------------
Extractor de datos fiscales de facturas de proveedores para el mercado español.

Motor LLM: Groq (llama-3.3-70b-versatile) — gratuito en desarrollo.

Uso básico
----------
    from extractor import FacturaExtractor

    extractor = FacturaExtractor()
    resultado = extractor.extraer_archivo("factura.pdf")
    resultado = extractor.extraer_archivo("factura.jpg")
    resultado = extractor.extraer(texto_plano)

    print(resultado.datos)           # dict completo
    print(resultado.confianza)       # 0.0 – 1.0
    print(resultado.campos_vacios)   # campos no encontrados
    print(resultado.resumen())       # resumen legible en consola

Uso por lotes
-------------
    resultados = extractor.extraer_lote(["f1.pdf", "f2.jpg", "f3.png"])

Dependencias opcionales para OCR mejorado
------------------------------------------
    pip install pytesseract Pillow pdf2image opencv-python-headless
    # Linux:  sudo apt install tesseract-ocr tesseract-ocr-spa poppler-utils
    # macOS:  brew install tesseract poppler
"""

from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from groq import Groq

load_dotenv()

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Constantes
# ─────────────────────────────────────────────

MODELO = "llama-3.3-70b-versatile"

UMBRAL_TEXTO_PDF = 80      # chars útiles mínimos para PDF "con texto"
ANCHO_MIN_OCR    = 1500    # px; imágenes más estrechas se reescalan
MAX_CHARS_LLM    = 7000    # total máximo de chars enviados al LLM
HEAD_CHARS        = 3500   # chars del inicio conservados en truncado
TAIL_CHARS        = 3000   # chars del final conservados en truncado

EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp", ".webp"}
EXTENSIONES_PDF    = {".pdf"}

CAMPOS_REQUERIDOS = [
    "nif_emisor",
    "nombre_emisor",
    "numero_factura",
    "fecha_emision",
    "base_imponible",
    "tipo_iva",
    "cuota_iva",
    "total_factura",
]

CAMPOS_OPCIONALES = [
    "nif_receptor",
    "nombre_receptor",
    "fecha_vencimiento",
    "concepto",
    "tipo_retencion",
    "retencion_irpf",
    "lineas_iva",
    "metodo_pago",
    "iban",
    "moneda",
    "notas",
]

TODOS_LOS_CAMPOS = CAMPOS_REQUERIDOS + CAMPOS_OPCIONALES

# ─────────────────────────────────────────────
# Prompts
# ─────────────────────────────────────────────

SYSTEM_PROMPT = """Eres un sistema experto en fiscalidad española especializado en
extraer datos de facturas de proveedores para clínicas médicas y despachos profesionales.

Tu única tarea es analizar el texto (y las tablas, si las hay) de una factura y extraer
los datos fiscales con la máxima precisión posible.

════════════════════════════════════════════
REGLAS GENERALES
════════════════════════════════════════════
1.  Devuelve ÚNICAMENTE un objeto JSON válido. Sin texto antes ni después.
2.  Si un campo no aparece en la factura, devuelve null para ese campo.
3.  Importes: siempre float sin símbolo €. Ejemplo: 121.00
4.  Fechas: siempre ISO 8601 YYYY-MM-DD. Ejemplo: 2024-03-15
5.  NIF/CIF sin espacios ni guiones. Ejemplo: B12345678
6.  IBAN sin espacios. Ejemplo: ES0000000000000000000000
7.  NUNCA copies el IBAN de ejemplos ni facturas previas.
    Extrae ÚNICAMENTE el IBAN literal de ESTA factura. Si no hay, devuelve null.
8.  No inventes datos. Si no está claro, devuelve null.
9.  Normaliza el nombre del emisor (razón social oficial).

════════════════════════════════════════════
REGLAS ANTI-ALUCINACIÓN — OBLIGATORIAS
════════════════════════════════════════════
Estas reglas tienen PRIORIDAD ABSOLUTA sobre cualquier otra instrucción.

NIF/CIF:
- Si NO ves un NIF o CIF escrito explícitamente en el texto, devuelve null.
- NUNCA construyas, calcules ni deduzcas un NIF. NUNCA uses datos del receptor
  (nombre, dirección) como si fueran del emisor.
- Un NIF inventado es SIEMPRE peor que un null.

Nombre del emisor:
- Usa SOLO el nombre o razón social que aparezca escrito como texto.
- Si el nombre del proveedor aparece como logotipo o imagen (no como texto
  extraído), devuelve null o el texto más cercano al logo, no la dirección.
- NUNCA uses el nombre del receptor o de la dirección de entrega como
  nombre del emisor.

Importes:
- Antes de devolver cualquier importe, comprueba que base + IVA ≈ total.
- Si el total no cuadra con base + IVA (diferencia > 0.10 €), revisa los
  números: es probable que hayas leído mal un decimal (ej: 51,10 → 510).
- En caso de duda entre dos lecturas posibles de un número, elige la que
  hace que base + IVA = total.
- Si aun así no cuadra, devuelve el total tal como aparece y añade una
  nota en el campo "notas" explicando la incoherencia.

Confianza interna:
- Si has devuelto null en nif_emisor O en nombre_emisor porque no estaban
  claros, añade en "notas": "NIF no visible en el documento" o
  "Nombre del emisor no legible".
- Si un importe te genera duda, añádelo también en "notas".

════════════════════════════════════════════
IVA — REGLAS CRÍTICAS
════════════════════════════════════════════
- tipo_iva: entero representativo del tramo principal (21, 10, 4 o 0).
- cuota_iva: suma total de TODAS las cuotas de IVA.
- base_imponible: suma total de TODAS las bases imponibles.
- Si hay MÚLTIPLES tramos de IVA, rellena lineas_iva con un objeto por tramo:
    { "base": 500.00, "tipo": 21, "cuota": 105.00 }
  En tipo_iva pon el tipo del tramo de mayor importe.
- Con un único tipo de IVA, lineas_iva puede ser null.

════════════════════════════════════════════
RETENCIÓN DE IRPF — REGLAS CRÍTICAS
════════════════════════════════════════════
Las facturas de autónomos y profesionales llevan retención que REDUCE el total.
Señales: "Retención IRPF", "R. IRPF", "Ret. 15%", "Ret. 7%", "-15%", "(-)"
- tipo_retencion: porcentaje como entero positivo (15, 7, 2...). null si no hay.
- retencion_irpf: importe retenido como float positivo. Ej: 300.00 (no -300.00)
- Fórmula: total_factura = base_imponible + cuota_iva - retencion_irpf
- Comprueba que el total cuadra con esta fórmula.

════════════════════════════════════════════
NIF / CIF — INSTRUCCIONES ESPECIALES
════════════════════════════════════════════
- Los PDFs pueden tener el texto desordenado: el NIF puede aparecer
  fragmentado, en el pie de página, en vertical o mezclado.
- CIF empresa española: letra + 7 dígitos + control. Ej: B45123456
- NIF persona física: 8 dígitos + letra. Ej: 12345678Z
- NIE extranjero: X/Y/Z + 7 dígitos + letra. Ej: X1234567L
- Si el prompt contiene NIFS_DETECTADOS_EN_TEXTO, úsalos como candidatos.
- CIFs de grandes proveedores conocidos (úsalos si el nombre es inequívoco):
    Telefónica de España S.A.U. / O2 / Movistar → A82018474
    Endesa Energía S.A.U.                        → A81948077
    Iberdrola Clientes S.A.U.                    → A95758389
    Iberdrola S.A.                               → A48010615
    Vodafone España S.A.U.                       → A80907397
    Orange España S.A.U.                         → A82009812
    Naturgy Energy Group S.A.                    → A08015497
    Repsol Electricidad y Gas S.A.               → A78052828
    Amazon EU S.à r.l. Sucursal España           → W0184081H
    El Corte Inglés S.A.                         → A28017895
    Mapfre España S.A.                           → A08055658
    Mutua Madrileña Automovilista                → V28119220
    Adeslas (SegurCaixa Adeslas S.A.)            → A08169815
    Sanitas S.A. de Seguros                      → A28587038
    DKV Seguros                                  → A50003526
    Correos (Sociedad Estatal Correos)           → A83052407
    Renfe Viajeros S.M.E.                        → A86868189
    Iberia L.A.E. S.A.                           → A28017648

════════════════════════════════════════════
FACTURAS DE AUTÓNOMOS — CHECKLIST
════════════════════════════════════════════
✓ El emisor tiene NIF de persona física (8 dígitos + letra).
✓ Suelen incluir retención IRPF (15% general, 7% primeros 3 años).
✓ total = base + IVA - retención.
✓ nombre_emisor puede ser el nombre completo de la persona, no empresa.

════════════════════════════════════════════
FACTURAS CON LOGO COMO NOMBRE DE PROVEEDOR
════════════════════════════════════════════
Algunas facturas muestran el nombre del proveedor solo como imagen/logo,
sin texto extraíble. En ese caso:
- Busca el nombre en el pie de página, en el texto del email o web
  (ej: "tienda.parislibreria.es" → "Librería París").
- Si no encuentras ningún texto que sea claramente el nombre comercial,
  devuelve null en nombre_emisor y explícalo en notas.
- NUNCA uses la dirección postal ni el nombre del receptor como nombre_emisor.

"""

USER_PROMPT_TEMPLATE = """Extrae todos los datos fiscales de la siguiente factura.

TEXTO DE LA FACTURA:
---
{texto}
---
{bloque_tablas}
Devuelve exactamente este JSON (no añadas campos extra):

{{
  "nif_emisor":       "NIF o CIF del proveedor emisor",
  "nombre_emisor":    "Razón social del proveedor",
  "nif_receptor":     "NIF o CIF del receptor, o null",
  "nombre_receptor":  "Razón social del receptor, o null",
  "numero_factura":   "Número o serie/número de la factura",
  "fecha_emision":    "YYYY-MM-DD",
  "fecha_vencimiento":"YYYY-MM-DD o null",
  "concepto":         "Descripción breve del servicio o producto",
  "base_imponible":   0.00,
  "tipo_iva":         21,
  "cuota_iva":        0.00,
  "lineas_iva": [
    {{"base": 0.00, "tipo": 21, "cuota": 0.00}}
  ],
  "tipo_retencion":   null,
  "retencion_irpf":   null,
  "total_factura":    0.00,
  "metodo_pago":      "transferencia / domiciliacion / cheque / efectivo / null",
  "iban":             "ESXX... o null",
  "moneda":           "EUR",
  "notas":            "Datos relevantes no recogidos arriba, o null"
}}
"""


# ─────────────────────────────────────────────
# Tipo de fuente
# ─────────────────────────────────────────────

class TipoFuente:
    PDF_TEXTO     = "pdf_texto"
    PDF_ESCANEADO = "pdf_escaneado"
    IMAGEN        = "imagen"


# ─────────────────────────────────────────────
# Resultado de extracción
# ─────────────────────────────────────────────

@dataclass
class ResultadoExtraccion:
    datos: dict[str, Any]
    texto_original: str
    confianza: float
    tipo_fuente: str = ""
    campos_vacios: list[str] = field(default_factory=list)
    advertencias: list[str] = field(default_factory=list)
    tokens_usados: int = 0
    modelo_usado: str = ""

    def es_valida(self) -> bool:
        return all(self.datos.get(c) is not None for c in CAMPOS_REQUERIDOS)

    def tiene_retencion(self) -> bool:
        return self.datos.get("retencion_irpf") is not None

    def tiene_multi_iva(self) -> bool:
        lineas = self.datos.get("lineas_iva")
        return isinstance(lineas, list) and len(lineas) > 1

    def resumen(self) -> str:
        icono_fuente = {
            TipoFuente.PDF_TEXTO:     "PDF digital",
            TipoFuente.PDF_ESCANEADO: "PDF escaneado (OCR)",
            TipoFuente.IMAGEN:        "Imagen (OCR)",
        }.get(self.tipo_fuente, "")

        lineas = ["─" * 60, f"  FACTURA EXTRAÍDA  [{icono_fuente}]", "─" * 60]

        etiquetas = {
            "nombre_emisor":    "Proveedor      ",
            "nif_emisor":       "NIF/CIF        ",
            "numero_factura":   "N\u00ba Factura     ",
            "fecha_emision":    "Fecha emisi\u00f3n  ",
            "fecha_vencimiento":"Vencimiento    ",
            "concepto":         "Concepto       ",
            "base_imponible":   "Base imponible ",
            "tipo_iva":         "Tipo IVA       ",
            "cuota_iva":        "Cuota IVA      ",
            "tipo_retencion":   "Tipo retenci\u00f3n ",
            "retencion_irpf":   "Retenci\u00f3n IRPF ",
            "total_factura":    "TOTAL          ",
            "metodo_pago":      "Forma de pago  ",
            "iban":             "IBAN           ",
        }
        for campo, etiqueta in etiquetas.items():
            valor = self.datos.get(campo)
            if valor is None:
                if campo in ("tipo_retencion", "retencion_irpf"):
                    continue
                valor = "\u2014"
            elif campo in ("base_imponible", "cuota_iva", "total_factura", "retencion_irpf"):
                valor = f"{float(valor):,.2f} \u20ac"
            elif campo in ("tipo_iva", "tipo_retencion"):
                valor = f"{valor}%"
            lineas.append(f"  {etiqueta}: {valor}")

        lineas_iva = self.datos.get("lineas_iva")
        if isinstance(lineas_iva, list) and len(lineas_iva) > 1:
            lineas.append("  Desglose IVA  :")
            for t in lineas_iva:
                lineas.append(
                    f"    \u2022 Base {t.get('base', 0):,.2f} \u20ac"
                    f" \u00d7 {t.get('tipo', 0)}%"
                    f" = {t.get('cuota', 0):,.2f} \u20ac"
                )

        lineas.append("─" * 60)
        estado = "\u2705 V\u00c1LIDA" if self.es_valida() else "\u26a0\ufe0f  INCOMPLETA"
        lineas.append(f"  Estado         : {estado}  (confianza: {self.confianza:.0%})")
        if self.campos_vacios:
            lineas.append(f"  Campos vac\u00edos  : {', '.join(self.campos_vacios)}")
        if self.advertencias:
            lineas.append("  Advertencias   :")
            for adv in self.advertencias:
                lineas.append(f"    \u26a0 {adv}")
        lineas.append(f"  Tokens usados  : {self.tokens_usados}")
        lineas.append("─" * 60)
        return "\n".join(lineas)

    def a_dict_plano(self) -> dict:
        datos = dict(self.datos)
        if isinstance(datos.get("lineas_iva"), list):
            datos["lineas_iva"] = json.dumps(datos["lineas_iva"], ensure_ascii=False)
        return {
            **datos,
            "_tipo_fuente":  self.tipo_fuente,
            "_confianza":    self.confianza,
            "_valida":       self.es_valida(),
            "_advertencias": "; ".join(self.advertencias) if self.advertencias else "",
        }


# ─────────────────────────────────────────────
# Preprocesador de imágenes OCR
# ─────────────────────────────────────────────

class PreprocesadorImagen:
    """
    Mejora la calidad de una imagen PIL antes de pasarla a Tesseract.

    Pipeline:
      1. Reescalar si es < ANCHO_MIN_OCR px
      2. Escala de grises
      3. Denoising (fastNlMeans)
      4. Umbralización adaptativa Otsu
      5. Corrección de orientación (deskew vía Tesseract OSD)

    Si OpenCV no está instalado, devuelve la imagen sin cambios.
    """

    @staticmethod
    def disponible() -> bool:
        try:
            import cv2  # noqa
            return True
        except ImportError:
            return False

    @classmethod
    def procesar(cls, imagen_pil):
        if not cls.disponible():
            logger.warning(
                "OpenCV no instalado; OCR sin preprocesado. "
                "Instala: pip install opencv-python-headless"
            )
            return imagen_pil

        import cv2
        import numpy as np
        from PIL import Image

        img = np.array(imagen_pil.convert("RGB"))
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        h, w = img.shape[:2]
        if w < ANCHO_MIN_OCR:
            factor = ANCHO_MIN_OCR / w
            img = cv2.resize(img, None, fx=factor, fy=factor,
                             interpolation=cv2.INTER_CUBIC)
            logger.info(f"  Imagen reescalada x{factor:.1f} ({w}px -> {img.shape[1]}px)")

        gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gris = cv2.fastNlMeansDenoising(gris, h=10, templateWindowSize=7, searchWindowSize=21)
        _, binaria = cv2.threshold(gris, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        binaria = cls._deskew(binaria)

        return Image.fromarray(binaria)

    @staticmethod
    def _deskew(img_np):
        try:
            import cv2
            import pytesseract
            osd = pytesseract.image_to_osd(img_np, output_type=pytesseract.Output.DICT)
            angulo = osd.get("rotate", 0)
            if abs(angulo) > 0.5:
                logger.info(f"  Deskew: corrigiendo {angulo} grados.")
                h, w = img_np.shape[:2]
                M = cv2.getRotationMatrix2D((w // 2, h // 2), -angulo, 1.0)
                img_np = cv2.warpAffine(img_np, M, (w, h),
                                        flags=cv2.INTER_CUBIC,
                                        borderMode=cv2.BORDER_REPLICATE)
        except Exception as e:
            logger.debug(f"  Deskew omitido: {e}")
        return img_np


# ─────────────────────────────────────────────
# Lector de archivos
# ─────────────────────────────────────────────

class LectorFactura:
    """
    Detecta el tipo de archivo y extrae texto + tablas.
    Devuelve (texto, tablas_str, tipo_fuente).
    """

    def leer(self, ruta: Path) -> tuple[str, str, str]:
        ext = ruta.suffix.lower()
        if ext in EXTENSIONES_IMAGEN:
            logger.info(f"Imagen detectada ({ext}). Preprocesando + OCR.")
            return self._ocr_imagen(ruta), "", TipoFuente.IMAGEN
        if ext in EXTENSIONES_PDF:
            return self._leer_pdf(ruta)
        raise ValueError(
            f"Formato no soportado: '{ext}'. "
            f"Soportados: PDF, {', '.join(sorted(EXTENSIONES_IMAGEN))}"
        )

    def _leer_pdf(self, ruta: Path) -> tuple[str, str, str]:
        try:
            import pdfplumber
        except ImportError:
            raise ImportError("Instala pdfplumber: pip install pdfplumber")

        texto_total  = ""
        tablas_total = []
        with pdfplumber.open(ruta) as pdf:
            for i, pagina in enumerate(pdf.pages, 1):
                texto_total += (pagina.extract_text() or "") + "\n"
                for tabla in pagina.extract_tables() or []:
                    filas = []
                    for fila in tabla:
                        celdas = [str(c).strip() if c else "" for c in fila]
                        if any(celdas):
                            filas.append(" | ".join(celdas))
                    if filas:
                        tablas_total.append(f"[Tabla pag.{i}]\n" + "\n".join(filas))

        texto_limpio = texto_total.strip()
        n_util = len(re.sub(r"\s", "", texto_limpio))
        tablas_str = "\n\n".join(tablas_total).strip()

        if n_util >= UMBRAL_TEXTO_PDF:
            logger.info(f"PDF digital: {len(texto_limpio)} chars, {len(tablas_total)} tabla(s).")
            return texto_limpio, tablas_str, TipoFuente.PDF_TEXTO

        logger.info(f"PDF escaneado ({n_util} chars utiles). Aplicando OCR...")
        return self._ocr_pdf(ruta), "", TipoFuente.PDF_ESCANEADO

    def _ocr_imagen(self, ruta: Path) -> str:
        try:
            import pytesseract
            from PIL import Image
        except ImportError:
            raise ImportError(
                "Instala: pip install pytesseract Pillow opencv-python-headless\n"
                "Linux: sudo apt install tesseract-ocr tesseract-ocr-spa\n"
                "macOS: brew install tesseract"
            )
        config = "--oem 3 --psm 6"
        original = Image.open(ruta)
        procesada = PreprocesadorImagen.procesar(original)
        texto = pytesseract.image_to_string(procesada, lang="spa+eng", config=config)
        if len(texto.strip()) < 50:
            logger.warning("  Preprocesado produjo poco texto, reintentando con imagen original.")
            texto = pytesseract.image_to_string(original, lang="spa+eng", config=config)
        if not texto.strip():
            raise ValueError(
                f"OCR no extrajo texto de '{ruta.name}'. "
                "Verifica que la imagen sea legible y Tesseract esté instalado."
            )
        logger.info(f"OCR imagen: {len(texto)} chars de '{ruta.name}'.")
        return texto

    def _ocr_pdf(self, ruta: Path) -> str:
        try:
            import pytesseract
            from pdf2image import convert_from_path
        except ImportError:
            raise ImportError(
                "Instala: pip install pytesseract Pillow pdf2image opencv-python-headless\n"
                "Linux: sudo apt install tesseract-ocr tesseract-ocr-spa poppler-utils\n"
                "macOS: brew install tesseract poppler"
            )
        config = "--oem 3 --psm 6"
        imagenes = convert_from_path(ruta, dpi=300)
        logger.info(f"  Convirtiendo {len(imagenes)} pagina(s) a imagen...")
        texto_total = ""
        for i, pil_img in enumerate(imagenes, 1):
            proc = PreprocesadorImagen.procesar(pil_img)
            texto_pag = pytesseract.image_to_string(proc, lang="spa+eng", config=config)
            if len(texto_pag.strip()) < 30:
                logger.warning(f"  Pag.{i}: poco texto con preprocesado, reintentando sin el.")
                texto_pag = pytesseract.image_to_string(pil_img, lang="spa+eng", config=config)
            texto_total += texto_pag + "\n"
            logger.info(f"  Pag.{i}/{len(imagenes)}: {len(texto_pag)} chars.")
        if not texto_total.strip():
            raise ValueError(f"OCR no extrajo texto del PDF escaneado '{ruta.name}'.")
        return texto_total


# ─────────────────────────────────────────────
# Normalización de fechas
# ─────────────────────────────────────────────

_MESES_ES = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12",
    "ene": "01", "feb": "02", "mar": "03", "abr": "04",
    "may": "05", "jun": "06", "jul": "07", "ago": "08",
    "sep": "09", "oct": "10", "nov": "11", "dic": "12",
}

def _normalizar_fecha(valor: str) -> str | None:
    """Convierte múltiples formatos de fecha a YYYY-MM-DD. Devuelve None si falla."""
    v = str(valor).strip()

    if re.match(r"^\d{4}-\d{2}-\d{2}$", v):
        return v

    # DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY
    m = re.match(r"^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})$", v)
    if m:
        d, mo, a = m.group(1), m.group(2), m.group(3)
        if len(a) == 2:
            a = "20" + a
        return f"{a}-{mo.zfill(2)}-{d.zfill(2)}"

    # "15 de marzo de 2024" o "15 marzo 2024"
    m = re.match(r"^(\d{1,2})\s+(?:de\s+)?(\w+)\s+(?:de\s+)?(\d{4})$", v, re.IGNORECASE)
    if m:
        d, mes_txt, a = m.group(1), m.group(2).lower(), m.group(3)
        mes = _MESES_ES.get(mes_txt)
        if mes:
            return f"{a}-{mes}-{d.zfill(2)}"

    # "marzo 2024" -> 01 de ese mes
    m = re.match(r"^(\w+)\s+(\d{4})$", v, re.IGNORECASE)
    if m:
        mes = _MESES_ES.get(m.group(1).lower())
        if mes:
            return f"{m.group(2)}-{mes}-01"

    return None


# ─────────────────────────────────────────────
# Validación y limpieza
# ─────────────────────────────────────────────

class ValidadorFactura:

    NIF_RE    = re.compile(r"^[A-Z0-9]{9}$")
    IBAN_RE   = re.compile(r"^[A-Z]{2}\d{2}[A-Z0-9]{4,}$")
    FECHA_RE  = re.compile(r"^\d{4}-\d{2}-\d{2}$")
    TIPOS_IVA = {0, 4, 5, 10, 21}

    def validar(self, datos: dict) -> tuple[dict, list[str], float]:
        adv: list[str] = []
        datos = self._limpiar_tipos(datos, adv)
        datos = self._normalizar_campos(datos, adv)
        datos = self._validar_lineas_iva(datos, adv)
        self._comprobar_coherencia(datos, adv)
        return datos, adv, self._calcular_confianza(datos, adv)

    def _limpiar_tipos(self, datos: dict, adv: list) -> dict:
        for campo in ("base_imponible", "cuota_iva", "total_factura", "retencion_irpf"):
            v = datos.get(campo)
            if v is not None:
                try:
                    datos[campo] = round(float(str(v).replace(",", ".")), 2)
                except (ValueError, TypeError):
                    adv.append(f"'{campo}' no es un numero valido: {v!r}")
                    datos[campo] = None
        for campo in ("tipo_iva", "tipo_retencion"):
            v = datos.get(campo)
            if v is not None:
                try:
                    datos[campo] = int(float(str(v).replace("%", "").strip()))
                except (ValueError, TypeError):
                    adv.append(f"'{campo}' no es un entero valido: {v!r}")
                    datos[campo] = None
        return datos

    def _normalizar_campos(self, datos: dict, adv: list) -> dict:
        for campo in ("nif_emisor", "nif_receptor"):
            v = datos.get(campo)
            if v:
                vc = re.sub(r"[\s\-\.]", "", str(v)).upper()
                datos[campo] = vc
                if not self.NIF_RE.match(vc):
                    adv.append(f"'{campo}' tiene formato inusual: {vc!r}")

        if datos.get("iban"):
            iban = re.sub(r"\s", "", str(datos["iban"])).upper()
            datos["iban"] = iban
            if not self.IBAN_RE.match(iban):
                adv.append(f"'iban' tiene formato inusual: {iban!r}")

        for campo in ("fecha_emision", "fecha_vencimiento"):
            v = datos.get(campo)
            if v and not self.FECHA_RE.match(str(v)):
                norm = _normalizar_fecha(str(v))
                if norm:
                    adv.append(f"'{campo}' normalizada: {v!r} -> {norm!r}")
                    datos[campo] = norm
                else:
                    adv.append(f"'{campo}' no reconocida: {v!r}")
                    datos[campo] = None
            v2 = datos.get(campo)
            if v2:
                try:
                    datetime.strptime(str(v2), "%Y-%m-%d")
                except ValueError:
                    adv.append(f"'{campo}' no es una fecha real: {v2!r}")
                    datos[campo] = None

        tipo = datos.get("tipo_iva")
        if tipo is not None and tipo not in self.TIPOS_IVA:
            adv.append(
                f"'tipo_iva' = {tipo}% no es estandar espanol "
                f"(validos: {sorted(self.TIPOS_IVA)})"
            )

        if not datos.get("moneda"):
            datos["moneda"] = "EUR"

        ret = datos.get("retencion_irpf")
        if ret is not None and ret < 0:
            datos["retencion_irpf"] = abs(ret)
            adv.append("'retencion_irpf' era negativa; convertida a positiva.")

        return datos

    def _validar_lineas_iva(self, datos: dict, adv: list) -> dict:
        lineas = datos.get("lineas_iva")
        if not isinstance(lineas, list) or not lineas:
            datos["lineas_iva"] = None
            return datos

        validas = []
        for tramo in lineas:
            if not isinstance(tramo, dict):
                continue
            try:
                validas.append({
                    "base":  round(float(str(tramo.get("base",  0)).replace(",", ".")), 2),
                    "tipo":  int(float(str(tramo.get("tipo",  0)).replace("%", ""))),
                    "cuota": round(float(str(tramo.get("cuota", 0)).replace(",", ".")), 2),
                })
            except (ValueError, TypeError):
                adv.append(f"Tramo de IVA invalido ignorado: {tramo}")

        datos["lineas_iva"] = validas if validas else None

        if validas and len(validas) > 1:
            if datos.get("base_imponible") is None:
                datos["base_imponible"] = round(sum(t["base"]  for t in validas), 2)
                adv.append("'base_imponible' calculada sumando lineas_iva.")
            if datos.get("cuota_iva") is None:
                datos["cuota_iva"] = round(sum(t["cuota"] for t in validas), 2)
                adv.append("'cuota_iva' calculada sumando lineas_iva.")

        return datos

    def _comprobar_coherencia(self, datos: dict, adv: list):
        base  = datos.get("base_imponible")
        cuota = datos.get("cuota_iva")
        total = datos.get("total_factura")
        tipo  = datos.get("tipo_iva")
        ret   = datos.get("retencion_irpf") or 0.0

        if base is not None and tipo is not None and cuota is None:
            datos["cuota_iva"] = round(base * tipo / 100, 2)
            cuota = datos["cuota_iva"]
            adv.append("'cuota_iva' calculada automaticamente (base x tipo_iva).")

        if base is not None and cuota is not None and total is not None:
            esperado = round(base + cuota - ret, 2)
            if abs(esperado - total) > 0.05:
                detalle = (
                    f"{base} + {cuota} - {ret} (retencion) = {esperado}"
                    if ret else f"base ({base}) + IVA ({cuota}) = {esperado}"
                )
                adv.append(
                    f"Incoherencia: {detalle}, pero total_factura = {total}. "
                    "Verifica la factura."
                )

    def _calcular_confianza(self, datos: dict, adv: list) -> float:
        presentes = sum(1 for c in CAMPOS_REQUERIDOS if datos.get(c) is not None)
        base      = presentes / len(CAMPOS_REQUERIDOS)
        penaliz   = min(len(adv) * 0.04, 0.30)
        return max(0.0, round(base - penaliz, 2))


# ─────────────────────────────────────────────
# Extractor principal
# ─────────────────────────────────────────────

class FacturaExtractor:
    """
    Extrae datos estructurados de facturas en cualquier formato.

    Detecta automaticamente PDF digital, PDF escaneado o imagen,
    aplica preprocesado OCR mejorado si es necesario, y llama al LLM
    para estructurar los datos en el esquema fiscal espanol.
    """

    def __init__(
        self,
        modelo: str = MODELO,
        max_tokens: int = 1500,
        temperatura: float = 0.0,
    ):
        self.cliente     = Groq()
        self.modelo      = modelo
        self.max_tokens  = max_tokens
        self.temperatura = temperatura
        self.validador   = ValidadorFactura()
        self.lector      = LectorFactura()

    # ── API pública ───────────────────────────

    def extraer_archivo(self, ruta: str | Path) -> ResultadoExtraccion:
        """Punto de entrada principal. Acepta PDF, JPG, PNG y otros formatos de imagen."""
        ruta = Path(ruta)
        if not ruta.exists():
            raise FileNotFoundError(f"Archivo no encontrado: {ruta}")
        texto, tablas_str, tipo_fuente = self.lector.leer(ruta)
        resultado = self._extraer_de_texto(texto, tablas_str)
        resultado.tipo_fuente = tipo_fuente
        return resultado

    def extraer(self, texto: str) -> ResultadoExtraccion:
        """Extrae datos a partir de texto plano ya obtenido."""
        if not texto or not texto.strip():
            raise ValueError("El texto de la factura no puede estar vacio.")
        return self._extraer_de_texto(texto, "")

    def extraer_lote(
        self,
        rutas: list[str | Path],
        continuar_en_error: bool = True,
    ) -> list[dict]:
        """Procesa multiples archivos. Devuelve lista de {archivo, resultado, error}."""
        resultados = []
        for ruta in rutas:
            nombre = Path(ruta).name
            try:
                resultado = self.extraer_archivo(ruta)
                resultados.append({"archivo": nombre, "resultado": resultado, "error": None})
                logger.info(
                    f"OK {nombre} [{resultado.tipo_fuente}] "
                    f"(confianza: {resultado.confianza:.0%})"
                )
            except Exception as exc:
                logger.error(f"ERROR en '{nombre}': {exc}")
                if not continuar_en_error:
                    raise
                resultados.append({"archivo": nombre, "resultado": None, "error": str(exc)})

        exitosos = sum(1 for r in resultados if r["error"] is None)
        logger.info(f"Lote: {exitosos}/{len(rutas)} facturas procesadas.")
        return resultados

    def extraer_pdf(self, ruta: str | Path) -> ResultadoExtraccion:
        """Alias de compatibilidad con v1."""
        return self.extraer_archivo(ruta)

    # ── Privado ───────────────────────────────

    _NIF_RE = re.compile(
        r"\b(?:[A-HJ-NP-SUVW]\d{7}[0-9A-J]|\d{8}[A-Z]|[XYZ]\d{7}[A-Z])\b",
        re.IGNORECASE,
    )

    def _extraer_de_texto(self, texto: str, tablas_str: str) -> ResultadoExtraccion:
        texto_proc = self._preprocesar_texto(texto)
        datos_raw, tokens = self._llamar_api(texto_proc, tablas_str)
        datos, adv, confianza = self.validador.validar(datos_raw)
        vacios = [c for c in TODOS_LOS_CAMPOS if datos.get(c) is None]
        return ResultadoExtraccion(
            datos=datos,
            texto_original=texto,
            confianza=confianza,
            campos_vacios=vacios,
            advertencias=adv,
            tokens_usados=tokens,
            modelo_usado=self.modelo,
        )

    def _preprocesar_texto(self, texto: str) -> str:
        texto = re.sub(r"[ \t]+", " ", texto)
        texto = re.sub(r"\n{3,}", "\n\n", texto)

        # Reconstruir NIFs fragmentados por columnas en PDFs
        texto_comp = re.sub(
            r"\b([A-Z])\s+(\d)\s+(\d)\s+(\d)\s+(\d)\s+(\d)\s+(\d)\s+(\d)\s+(\d)\b",
            lambda m: "".join(m.groups()),
            texto,
        )
        candidatos = list(dict.fromkeys(
            [c.upper() for c in
             self._NIF_RE.findall(texto) + self._NIF_RE.findall(texto_comp)]
        ))
        if candidatos:
            logger.info(f"NIFs detectados por regex: {candidatos}")
            anotacion = (
                "\n\nNIFS_DETECTADOS_EN_TEXTO:\n"
                + ", ".join(candidatos)
                + "\nUsa estos candidatos para nif_emisor y/o nif_receptor.\n"
            )
        else:
            logger.warning("No se detecto ningun NIF/CIF por regex.")
            anotacion = ""

        # Truncado inteligente: cabecera + pie
        if len(texto) > MAX_CHARS_LLM:
            logger.warning(f"Texto largo ({len(texto)} chars). Truncando inteligentemente.")
            texto = texto[:HEAD_CHARS] + "\n[...]\n" + texto[-TAIL_CHARS:]

        return (texto + anotacion).strip()

    def _llamar_api(self, texto: str, tablas_str: str) -> tuple[dict, int]:
        bloque_tablas = ""
        if tablas_str:
            bloque_tablas = (
                "\nTABLAS DETECTADAS EN EL PDF:\n---\n"
                + tablas_str[:2000]
                + "\n---\n"
            )
        prompt = USER_PROMPT_TEMPLATE.format(texto=texto, bloque_tablas=bloque_tablas)

        for intento in range(3):
            try:
                resp = self.cliente.chat.completions.create(
                    model=self.modelo,
                    max_tokens=self.max_tokens,
                    temperature=self.temperatura,
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user",   "content": prompt},
                    ],
                )
                tokens    = resp.usage.total_tokens
                contenido = self._extraer_json(resp.choices[0].message.content.strip())
                datos     = json.loads(contenido)
                for campo in TODOS_LOS_CAMPOS:
                    datos.setdefault(campo, None)
                return datos, tokens

            except json.JSONDecodeError as e:
                if intento < 2:
                    logger.warning(f"JSON malformado (intento {intento+1}/3): {e}")
                    continue
                raise ValueError(f"Respuesta no valida tras 3 intentos: {e}")
            except Exception as e:
                raise RuntimeError(f"Error en la API de Groq: {e}")

        return {}, 0

    @staticmethod
    def _extraer_json(texto: str) -> str:
        m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", texto, re.DOTALL)
        if m:
            return m.group(1)
        m = re.search(r"\{.*\}", texto, re.DOTALL)
        if m:
            return m.group(0)
        return texto


# ─────────────────────────────────────────────
# Función de conveniencia
# ─────────────────────────────────────────────

def extraer_factura(ruta_o_texto: str, modelo: str = MODELO) -> ResultadoExtraccion:
    """
    Atajo para uso rapido. Acepta ruta de archivo o texto plano.

        from extractor import extraer_factura
        resultado = extraer_factura("factura.pdf")
        resultado = extraer_factura("factura.jpg")
    """
    extractor = FacturaExtractor(modelo=modelo)
    p = Path(ruta_o_texto)
    if p.exists() and p.suffix.lower() in EXTENSIONES_PDF | EXTENSIONES_IMAGEN:
        return extractor.extraer_archivo(p)
    return extractor.extraer(ruta_o_texto)


# ─────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s")

    DEMO_EMPRESA = """
    FARMACIA LOPEZ Y ASOCIADOS S.L.   CIF: B45123456
    Calle Mayor, 14 - 45002 Toledo
    FACTURA N: 2024/0342  Fecha: 15/03/2024  Vencimiento: 30/03/2024
    Cliente: Clinica Dental Ruiz S.L.  NIF: B78901234
    Material fungible odontologico   850,00
    Productos antisepticos           120,00
    BASE IMPONIBLE                   970,00
    IVA 21%                          203,70
    TOTAL FACTURA                  1.173,70
    Pago: Transferencia  IBAN: ES12 3456 7890 1234 5678 9012
    """

    DEMO_AUTONOMO = """
    Ana Garcia Martinez   NIF: 45678901B   Disenadora Grafica
    Calle Paloma 7, 28012 Madrid
    FACTURA N: 2024/018   Fecha: 5 de febrero de 2024
    Cliente: Clinica Dental Ruiz S.L.  NIF: B78901234
    Diseno identidad corporativa enero 2024
    Base imponible         2.000,00
    IVA 21%                  420,00
    Retencion IRPF 15%      -300,00
    TOTAL A PAGAR          2.120,00
    IBAN: ES91 2100 0418 4502 0005 1332
    """

    DEMO_MULTIIVA = """
    SUMINISTROS MEDICOS PENINSULARES S.L.   CIF: B87654321
    Poligono Industrial Norte, Nave 12 - 08820 El Prat de Llobregat
    FACTURA: SMP-2024-00789   Fecha: 20 enero 2024
    Cliente: Clinica Dental Ruiz S.L.  NIF: B78901234
    Equipo radiografia digital   3.500,00   21%   735,00
    Guantes desechables (caja)     150,00   21%    31,50
    Anestesico dental (vial)       400,00   10%    40,00
    Apositos esteriles             200,00    4%     8,00
    TOTAL BASE IMPONIBLE         4.250,00
    TOTAL CUOTA IVA                814,50
    TOTAL FACTURA                5.064,50
    Pago: Transferencia   IBAN: ES76 0049 1500 0527 1019 2350
    """

    demos = {
        "--demo":          ("empresa",   DEMO_EMPRESA),
        "--demo-autonomo": ("autonomo",  DEMO_AUTONOMO),
        "--demo-multiiva": ("multi-IVA", DEMO_MULTIIVA),
    }

    if len(sys.argv) < 2:
        print(
            "\nUso:\n"
            "  python extractor.py <factura.pdf>      -> PDF digital o escaneado\n"
            "  python extractor.py <factura.jpg>      -> imagen JPG/PNG\n"
            "  python extractor.py --demo             -> demo empresa\n"
            "  python extractor.py --demo-autonomo    -> demo autonomo con retencion\n"
            "  python extractor.py --demo-multiiva    -> demo multiples tipos IVA\n"
            "  python extractor.py f1.pdf f2.jpg      -> lote de archivos\n"
            "  python extractor.py factura.pdf --json -> salida en JSON\n"
        )
        sys.exit(1)

    for flag, (etiqueta, texto_demo) in demos.items():
        if sys.argv[1] == flag:
            print(f"\nDemo: factura de {etiqueta}\n")
            r = FacturaExtractor().extraer(texto_demo)
            print(r.resumen())
            print("\nDatos JSON:")
            print(json.dumps(r.datos, ensure_ascii=False, indent=2))
            sys.exit(0)

    mostrar_json = "--json" in sys.argv
    archivos = [a for a in sys.argv[1:] if not a.startswith("--")]
    extractor = FacturaExtractor()

    if len(archivos) == 1:
        resultado = extractor.extraer_archivo(archivos[0])
        print(resultado.resumen() if not mostrar_json else
              json.dumps(resultado.datos, ensure_ascii=False, indent=2))
    else:
        resultados = extractor.extraer_lote(archivos)
        for r in resultados:
            print(f"\n{'─'*40}\n# {r['archivo']}")
            if r["error"]:
                print(f"ERROR: {r['error']}")
            elif mostrar_json:
                print(json.dumps(r["resultado"].datos, ensure_ascii=False, indent=2))
            else:
                print(r["resultado"].resumen())
