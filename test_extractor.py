"""
test_extractor.py  ·  FacturaIA v2
------------------------------------
Suite de tests y métricas de precisión para FacturaExtractor.

¿Qué hace?
----------
  1. Escanea data/test/ buscando pares  factura.pdf + factura.json
     (el .json es el "ground truth" que tú validas a mano).
  2. Ejecuta el extractor sobre cada factura.
  3. Compara campo a campo el resultado con el ground truth.
  4. Imprime un informe de precisión por campo + tabla de errores.
  5. Opcionalmente exporta el informe completo a Excel.

Estructura de carpetas esperada
--------------------------------
  data/
  └── test/
      ├── factura_amazon.pdf
      ├── factura_amazon.json       ← ground truth
      ├── factura_iberdrola.pdf
      ├── factura_iberdrola.json
      ├── factura_autonomo.jpg
      ├── factura_autonomo.json
      └── ...

Formato del ground truth (.json)
----------------------------------
  Crea un .json por cada factura con los valores CORRECTOS que esperas.
  Solo necesitas incluir los campos que quieres testear; el resto se ignora.

  {
    "nif_emisor":      "A82018474",
    "nombre_emisor":   "Telefónica de España S.A.U.",
    "numero_factura":  "OM1VMCJ0495490",
    "fecha_emision":   "2026-03-01",
    "base_imponible":  97.93,
    "tipo_iva":        21,
    "cuota_iva":       20.57,
    "total_factura":   118.50,
    "retencion_irpf":  null,
    "metodo_pago":     "domiciliacion"
  }

Generación asistida del ground truth
--------------------------------------
  Si aún no tienes los .json, usa el modo --generar-gt para que el extractor
  procese las facturas y genere borradores .json que tú solo tienes que revisar
  y corregir a mano:

    python test_extractor.py --generar-gt

  Esto NO lanza los tests; solo crea los borradores para que los valides.

Uso
---
  # Correr todos los tests
  python test_extractor.py

  # Correr solo una factura concreta
  python test_extractor.py --archivo factura_amazon.pdf

  # Exportar informe a Excel
  python test_extractor.py --excel

  # Generar borradores de ground truth (sin lanzar tests)
  python test_extractor.py --generar-gt

  # Cambiar la carpeta de tests
  python test_extractor.py --dir /otra/carpeta

Tolerancias de comparación
----------------------------
  - Campos numéricos (importes): diferencia ≤ 0.05 € → OK
  - Campos de texto (NIF, nombre...): coincidencia exacta tras normalización
    (mayúsculas, sin espacios extra, sin tildes opcionales)
  - Fechas: comparación exacta en formato YYYY-MM-DD
  - Campos null en el ground truth: se ignoran (no penalizan ni puntúan)
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Configuración
# ─────────────────────────────────────────────

DIR_TEST_DEFAULT = Path("data/test")

# Campos numéricos con tolerancia de 0.05 €
CAMPOS_NUMERICOS = {
    "base_imponible", "cuota_iva", "total_factura",
    "retencion_irpf", "tipo_iva", "tipo_retencion",
}

# Campos que se comparan con normalización de texto
CAMPOS_TEXTO = {
    "nif_emisor", "nif_receptor", "nombre_emisor", "nombre_receptor",
    "numero_factura", "concepto", "metodo_pago", "iban", "moneda",
}

# Campos de fecha (comparación exacta ISO)
CAMPOS_FECHA = {"fecha_emision", "fecha_vencimiento"}

# Campos que NO se testean automáticamente (demasiado libres o estructurados)
CAMPOS_IGNORADOS = {"lineas_iva", "notas"}

TOLERANCIA_EUROS = 0.05

# Colores ANSI para la consola
class C:
    VERDE   = "\033[92m"
    ROJO    = "\033[91m"
    AMARILLO= "\033[93m"
    AZUL    = "\033[94m"
    GRIS    = "\033[90m"
    BOLD    = "\033[1m"
    RESET   = "\033[0m"

    @staticmethod
    def ok(texto):    return f"{C.VERDE}{texto}{C.RESET}"
    @staticmethod
    def fail(texto):  return f"{C.ROJO}{texto}{C.RESET}"
    @staticmethod
    def warn(texto):  return f"{C.AMARILLO}{texto}{C.RESET}"
    @staticmethod
    def info(texto):  return f"{C.AZUL}{texto}{C.RESET}"
    @staticmethod
    def bold(texto):  return f"{C.BOLD}{texto}{C.RESET}"
    @staticmethod
    def gris(texto):  return f"{C.GRIS}{texto}{C.RESET}"


# ─────────────────────────────────────────────
# Resultado de un test individual
# ─────────────────────────────────────────────

@dataclass
class ResultadoCampo:
    campo: str
    esperado: Any
    obtenido: Any
    ok: bool
    motivo: str = ""   # descripción del fallo si ok=False

@dataclass
class ResultadoTest:
    archivo: str
    campos: list[ResultadoCampo] = field(default_factory=list)
    error_extraccion: str = ""        # si el extractor lanzó excepción
    confianza_extractor: float = 0.0
    tipo_fuente: str = ""
    tokens_usados: int = 0

    @property
    def ok(self) -> bool:
        return not self.error_extraccion and all(c.ok for c in self.campos)

    @property
    def n_ok(self) -> int:
        return sum(1 for c in self.campos if c.ok)

    @property
    def n_total(self) -> int:
        return len(self.campos)

    @property
    def precision(self) -> float:
        return self.n_ok / self.n_total if self.n_total else 0.0


# ─────────────────────────────────────────────
# Comparador de campos
# ─────────────────────────────────────────────

def _normalizar_texto(valor: Any) -> str:
    """Normaliza para comparación: mayúsculas, sin espacios extra, sin tildes."""
    if valor is None:
        return ""
    s = str(valor).strip().upper()
    s = re.sub(r"\s+", " ", s)
    # quitar tildes opcionales (ej: "DOMICILIACION" == "DOMICILIACIÓN")
    s = "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )
    return s

def _comparar_campo(campo: str, esperado: Any, obtenido: Any) -> ResultadoCampo:
    """Compara un campo con la tolerancia adecuada según su tipo."""

    # Si el esperado es null en el ground truth → ignorar
    if esperado is None:
        return ResultadoCampo(campo, esperado, obtenido, ok=True, motivo="ignorado (null en GT)")

    # El extractor no encontró el valor
    if obtenido is None:
        return ResultadoCampo(
            campo, esperado, obtenido, ok=False,
            motivo=f"extractor devolvió null, esperaba {esperado!r}"
        )

    # ── Numérico ──────────────────────────────
    if campo in CAMPOS_NUMERICOS:
        try:
            esp = float(esperado)
            obt = float(obtenido)
            if abs(esp - obt) <= TOLERANCIA_EUROS:
                return ResultadoCampo(campo, esperado, obtenido, ok=True)
            else:
                return ResultadoCampo(
                    campo, esperado, obtenido, ok=False,
                    motivo=f"esperado {esp}, obtenido {obt} (diff={abs(esp-obt):.2f})"
                )
        except (ValueError, TypeError):
            return ResultadoCampo(
                campo, esperado, obtenido, ok=False,
                motivo=f"no se pudo comparar como número: {obtenido!r}"
            )

    # ── Fecha ─────────────────────────────────
    if campo in CAMPOS_FECHA:
        e = str(esperado).strip()
        o = str(obtenido).strip()
        ok = (e == o)
        return ResultadoCampo(
            campo, esperado, obtenido, ok=ok,
            motivo="" if ok else f"esperado {e!r}, obtenido {o!r}"
        )

    # ── Texto ─────────────────────────────────
    e = _normalizar_texto(esperado)
    o = _normalizar_texto(obtenido)
    ok = (e == o)

    # Para nombres de empresa: permitir que el obtenido CONTENGA el esperado
    # (ej: "TELEFÓNICA DE ESPAÑA" dentro de "TELEFÓNICA DE ESPAÑA S.A.U.")
    if not ok and campo in ("nombre_emisor", "nombre_receptor", "concepto"):
        ok = (e in o) or (o in e)
        motivo = "" if ok else f"esperado {e!r}, obtenido {o!r}"
    else:
        motivo = "" if ok else f"esperado {e!r}, obtenido {o!r}"

    return ResultadoCampo(campo, esperado, obtenido, ok=ok, motivo=motivo)


# ─────────────────────────────────────────────
# Ejecutor de tests
# ─────────────────────────────────────────────

class EjecutorTests:

    def __init__(self, dir_test: Path):
        self.dir_test = dir_test
        self._extractor = None   # lazy init (evita importar Groq si no hay tests)

    @property
    def extractor(self):
        if self._extractor is None:
            from core.extractor import FacturaExtractor
            self._extractor = FacturaExtractor()
        return self._extractor

    def descubrir_pares(self) -> list[tuple[Path, Path]]:
        """Encuentra pares (factura, ground_truth.json) en dir_test."""
        ext_soportadas = {".pdf", ".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp", ".webp"}
        pares = []
        for factura in sorted(self.dir_test.iterdir()):
            if factura.suffix.lower() not in ext_soportadas:
                continue
            gt = factura.with_suffix(".json")
            if gt.exists():
                pares.append((factura, gt))
            else:
                logger.warning(
                    f"  Sin ground truth: {factura.name} "
                    f"(crea {gt.name} para incluirla en los tests)"
                )
        return pares

    def ejecutar_uno(self, factura: Path, gt_path: Path) -> ResultadoTest:
        """Extrae y compara contra el ground truth."""
        resultado = ResultadoTest(archivo=factura.name)

        # Cargar ground truth
        try:
            with open(gt_path, encoding="utf-8") as f:
                ground_truth: dict = json.load(f)
        except Exception as e:
            resultado.error_extraccion = f"Error leyendo ground truth: {e}"
            return resultado

        # Ejecutar extractor
        try:
            ext_resultado = self.extractor.extraer_archivo(factura)
            resultado.confianza_extractor = ext_resultado.confianza
            resultado.tipo_fuente         = ext_resultado.tipo_fuente
            resultado.tokens_usados       = ext_resultado.tokens_usados
            datos_obtenidos               = ext_resultado.datos
        except Exception as e:
            resultado.error_extraccion = str(e)
            return resultado

        # Comparar campo a campo
        campos_a_testear = (
            set(ground_truth.keys())
            - CAMPOS_IGNORADOS
        )
        for campo in sorted(campos_a_testear):
            esperado = ground_truth.get(campo)
            obtenido = datos_obtenidos.get(campo)
            resultado.campos.append(_comparar_campo(campo, esperado, obtenido))

        return resultado

    def ejecutar_todos(self, pares: list[tuple[Path, Path]]) -> list[ResultadoTest]:
        resultados = []
        total = len(pares)
        for i, (factura, gt) in enumerate(pares, 1):
            print(f"  [{i}/{total}] {factura.name}...", end=" ", flush=True)
            r = self.ejecutar_uno(factura, gt)
            if r.error_extraccion:
                print(C.fail(f"ERROR — {r.error_extraccion}"))
            else:
                icono = C.ok("OK") if r.ok else C.warn(f"{r.n_ok}/{r.n_total} campos")
                print(f"{icono}  (confianza: {r.confianza_extractor:.0%})")
            resultados.append(r)
        return resultados


# ─────────────────────────────────────────────
# Generador de ground truth (modo --generar-gt)
# ─────────────────────────────────────────────

class GeneradorGroundTruth:
    """
    Procesa las facturas sin ground truth y crea borradores .json
    que el usuario debe revisar y corregir a mano.
    """

    def __init__(self, dir_test: Path):
        self.dir_test = dir_test

    def generar(self):
        from core.extractor import FacturaExtractor, EXTENSIONES_IMAGEN, EXTENSIONES_PDF
        extractor = FacturaExtractor()
        ext_soportadas = EXTENSIONES_PDF | EXTENSIONES_IMAGEN
        generados = 0

        for factura in sorted(self.dir_test.iterdir()):
            if factura.suffix.lower() not in ext_soportadas:
                continue
            gt_path = factura.with_suffix(".json")
            if gt_path.exists():
                print(C.gris(f"  Saltando {factura.name} (ya tiene ground truth)"))
                continue

            print(f"  Procesando {factura.name}...", end=" ", flush=True)
            try:
                resultado = extractor.extraer_archivo(factura)
                # Crear borrador con todos los campos (excepto los ignorados)
                borrador = {
                    k: v for k, v in resultado.datos.items()
                    if k not in CAMPOS_IGNORADOS
                }
                # Marcar como borrador para que el usuario lo sepa
                borrador["_BORRADOR"] = (
                    "REVISA Y CORRIGE ESTE ARCHIVO. "
                    "Elimina esta clave cuando esté validado."
                )
                with open(gt_path, "w", encoding="utf-8") as f:
                    json.dump(borrador, f, ensure_ascii=False, indent=2)
                print(C.ok(f"Borrador creado → {gt_path.name}"))
                generados += 1
            except Exception as e:
                print(C.fail(f"Error: {e}"))

        print(f"\n  {generados} borradores generados en {self.dir_test}")
        print(
            "\n  SIGUIENTE PASO:\n"
            "  1. Abre cada .json generado\n"
            "  2. Corrige los campos incorrectos comparándolos con la factura real\n"
            "  3. Elimina la clave '_BORRADOR' cuando esté validado\n"
            "  4. Corre: python test_extractor.py\n"
        )


# ─────────────────────────────────────────────
# Informe de resultados
# ─────────────────────────────────────────────

class InformeTests:

    def __init__(self, resultados: list[ResultadoTest]):
        self.resultados = resultados

    # ── Consola ───────────────────────────────

    def imprimir(self):
        self._imprimir_tabla_resumen()
        self._imprimir_precision_por_campo()
        self._imprimir_errores_detallados()
        self._imprimir_resumen_final()

    def _imprimir_tabla_resumen(self):
        print(f"\n{C.bold('═' * 70)}")
        print(C.bold(f"  RESULTADOS DE TESTS — {datetime.now().strftime('%d/%m/%Y %H:%M')}"))
        print(C.bold('═' * 70))
        print(
            f"  {'Archivo':<35} {'Estado':<12} "
            f"{'Campos':<10} {'Confianza':<10} Fuente"
        )
        print("  " + "─" * 68)
        for r in self.resultados:
            if r.error_extraccion:
                estado  = C.fail("ERROR")
                campos  = "—"
                conf    = "—"
            elif r.ok:
                estado  = C.ok("✓ PERFECTO")
                campos  = C.ok(f"{r.n_ok}/{r.n_total}")
                conf    = f"{r.confianza_extractor:.0%}"
            else:
                estado  = C.warn("⚠ PARCIAL")
                campos  = C.warn(f"{r.n_ok}/{r.n_total}")
                conf    = f"{r.confianza_extractor:.0%}"
            print(
                f"  {r.archivo:<35} {estado:<22} "
                f"{campos:<20} {conf:<10} {r.tipo_fuente}"
            )

    def _imprimir_precision_por_campo(self):
        # Agregar resultados por campo
        stats: dict[str, dict] = {}
        for r in self.resultados:
            if r.error_extraccion:
                continue
            for c in r.campos:
                if c.motivo == "ignorado (null en GT)":
                    continue
                if c.campo not in stats:
                    stats[c.campo] = {"ok": 0, "total": 0, "fallos": []}
                stats[c.campo]["total"] += 1
                if c.ok:
                    stats[c.campo]["ok"] += 1
                else:
                    stats[c.campo]["fallos"].append((r.archivo, c.motivo))

        if not stats:
            return

        print(f"\n{C.bold('  PRECISIÓN POR CAMPO')}")
        print("  " + "─" * 60)

        # Ordenar: campos requeridos primero, luego por precisión ascendente
        from core.extractor import CAMPOS_REQUERIDOS
        def sort_key(item):
            campo, s = item
            req = 0 if campo in CAMPOS_REQUERIDOS else 1
            prec = s["ok"] / s["total"] if s["total"] else 1
            return (req, prec)

        for campo, s in sorted(stats.items(), key=sort_key):
            ok    = s["ok"]
            total = s["total"]
            prec  = ok / total if total else 0
            barra = self._barra(prec, ancho=20)
            req   = " *" if campo in CAMPOS_REQUERIDOS else "  "
            color = C.ok if prec >= 0.9 else (C.warn if prec >= 0.6 else C.fail)
            print(
                f"  {req}{campo:<22} {barra} "
                f"{color(f'{prec:.0%}'):>18}  ({ok}/{total})"
            )
        print(f"\n  {C.gris('* = campo requerido')}")

    def _imprimir_errores_detallados(self):
        hay_errores = any(
            not c.ok and c.motivo != "ignorado (null en GT)"
            for r in self.resultados
            for c in r.campos
        ) or any(r.error_extraccion for r in self.resultados)

        if not hay_errores:
            return

        print(f"\n{C.bold('  DETALLE DE FALLOS')}")
        print("  " + "─" * 60)

        for r in self.resultados:
            if r.error_extraccion:
                print(f"\n  {C.fail('✘')} {C.bold(r.archivo)}")
                print(f"    {C.fail('Error de extracción:')} {r.error_extraccion}")
                continue

            fallos = [c for c in r.campos if not c.ok and c.motivo != "ignorado (null en GT)"]
            if not fallos:
                continue

            print(f"\n  {C.warn('⚠')} {C.bold(r.archivo)}")
            for c in fallos:
                print(f"    {c.campo:<22} → {C.fail(c.motivo)}")

    def _imprimir_resumen_final(self):
        total    = len(self.resultados)
        errores  = sum(1 for r in self.resultados if r.error_extraccion)
        perfectos= sum(1 for r in self.resultados if r.ok and not r.error_extraccion)
        parciales= total - errores - perfectos

        # Precisión global (solo sobre facturas sin error de extracción)
        todos_campos = [
            c for r in self.resultados if not r.error_extraccion
            for c in r.campos
            if c.motivo != "ignorado (null en GT)"
        ]
        prec_global = (
            sum(1 for c in todos_campos if c.ok) / len(todos_campos)
            if todos_campos else 0
        )

        tokens_total = sum(r.tokens_usados for r in self.resultados)

        print(f"\n{C.bold('═' * 70)}")
        print(C.bold("  RESUMEN FINAL"))
        print(C.bold('═' * 70))
        print(f"  Facturas testeadas  : {total}")
        print(f"  Perfectas (100%)    : {C.ok(perfectos)}")
        print(f"  Parciales           : {C.warn(parciales)}")
        print(f"  Errores extracción  : {C.fail(errores)}")
        print(f"  Precisión global    : {C.bold(f'{prec_global:.1%}')}")
        print(f"  Tokens API usados   : {tokens_total:,}")
        print(C.bold('═' * 70))

        if prec_global >= 0.95:
            print(C.ok("  🎉 Excelente. El extractor está listo para producción."))
        elif prec_global >= 0.80:
            print(C.warn("  ⚠  Buena base. Revisa los fallos antes de producción."))
        else:
            print(C.fail("  ✘  Precisión insuficiente. Revisa el prompt y el validador."))
        print()

    @staticmethod
    def _barra(proporcion: float, ancho: int = 20) -> str:
        llenos  = round(proporcion * ancho)
        vacios  = ancho - llenos
        if proporcion >= 0.9:
            color = C.VERDE
        elif proporcion >= 0.6:
            color = C.AMARILLO
        else:
            color = C.ROJO
        return f"{color}{'█' * llenos}{'░' * vacios}{C.RESET}"

    # ── Excel ─────────────────────────────────

    def exportar_excel(self, ruta: Path):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
            from openpyxl.utils import get_column_letter
        except ImportError:
            print(C.warn("openpyxl no instalado. Instala con: pip install openpyxl"))
            return

        wb = Workbook()

        # ── Hoja 1: Resultados por factura ────
        ws1 = wb.active
        ws1.title = "Resultados"
        self._excel_hoja_resultados(ws1)

        # ── Hoja 2: Precisión por campo ───────
        ws2 = wb.create_sheet("Precisión por Campo")
        self._excel_hoja_precision(ws2)

        # ── Hoja 3: Errores detallados ────────
        ws3 = wb.create_sheet("Errores Detallados")
        self._excel_hoja_errores(ws3)

        wb.save(ruta)
        print(C.ok(f"\n  Informe Excel exportado → {ruta}"))

    def _excel_estilo_base(self):
        from openpyxl.styles import Border, Font, Side
        thin  = Side(style="thin", color="CCCCCC")
        borde = Border(left=thin, right=thin, top=thin, bottom=thin)
        return borde, Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)

    def _excel_titulo(self, ws, texto, n_cols):
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
        c = ws["A1"]
        c.value     = texto
        c.font      = Font(name="Arial", size=12, bold=True, color="FFFFFF")
        c.fill      = PatternFill("solid", start_color="0D3B66")
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

    def _excel_cabecera(self, ws, row, columnas):
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        borde, _, _ = self._excel_estilo_base()
        fill = PatternFill("solid", start_color="1F4E79")
        font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        for c, (etiqueta, ancho) in enumerate(columnas, 1):
            cell = ws.cell(row=row, column=c, value=etiqueta)
            cell.font = font; cell.fill = fill; cell.border = borde
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(c)].width = ancho
        ws.row_dimensions[row].height = 28

    def _excel_hoja_resultados(self, ws):
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        borde, font, bold = self._excel_estilo_base()

        self._excel_titulo(ws, f"FacturaIA — Resultados de Tests  ·  {datetime.now().strftime('%d/%m/%Y %H:%M')}", 7)

        cols = [
            ("Archivo", 35), ("Estado", 14), ("Campos OK", 11),
            ("Precisión", 11), ("Confianza", 11), ("Tipo Fuente", 18), ("Error", 40),
        ]
        self._excel_cabecera(ws, 2, cols)

        FILL_OK   = PatternFill("solid", start_color="E2EFDA")
        FILL_WARN = PatternFill("solid", start_color="FFF2CC")
        FILL_ERR  = PatternFill("solid", start_color="FCE4D6")

        for r_idx, r in enumerate(self.resultados, 3):
            if r.error_extraccion:
                fill   = FILL_ERR
                estado = "ERROR"
                campos = "—"
                prec   = "—"
            elif r.ok:
                fill   = FILL_OK
                estado = "PERFECTO"
                campos = f"{r.n_ok}/{r.n_total}"
                prec   = f"{r.precision:.0%}"
            else:
                fill   = FILL_WARN
                estado = "PARCIAL"
                campos = f"{r.n_ok}/{r.n_total}"
                prec   = f"{r.precision:.0%}"

            fila = [
                r.archivo, estado, campos, prec,
                f"{r.confianza_extractor:.0%}" if not r.error_extraccion else "—",
                r.tipo_fuente,
                r.error_extraccion or "",
            ]
            for c_idx, valor in enumerate(fila, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=valor)
                cell.font = font; cell.fill = fill; cell.border = borde
                cell.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[r_idx].height = 18

    def _excel_hoja_precision(self, ws):
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        borde, font, bold = self._excel_estilo_base()

        self._excel_titulo(ws, "Precisión por Campo", 5)

        cols = [
            ("Campo", 25), ("Requerido", 12), ("OK", 8),
            ("Total", 8), ("Precisión", 12),
        ]
        self._excel_cabecera(ws, 2, cols)

        stats: dict[str, dict] = {}
        for r in self.resultados:
            if r.error_extraccion:
                continue
            for c in r.campos:
                if c.motivo == "ignorado (null en GT)":
                    continue
                if c.campo not in stats:
                    stats[c.campo] = {"ok": 0, "total": 0}
                stats[c.campo]["total"] += 1
                if c.ok:
                    stats[c.campo]["ok"] += 1

        try:
            from core.extractor import CAMPOS_REQUERIDOS
        except Exception:
            CAMPOS_REQUERIDOS = []

        FILL_OK   = PatternFill("solid", start_color="E2EFDA")
        FILL_WARN = PatternFill("solid", start_color="FFF2CC")
        FILL_ERR  = PatternFill("solid", start_color="FCE4D6")

        def sort_key(item):
            campo, s = item
            req  = 0 if campo in CAMPOS_REQUERIDOS else 1
            prec = s["ok"] / s["total"] if s["total"] else 1
            return (req, prec)

        for r_idx, (campo, s) in enumerate(sorted(stats.items(), key=sort_key), 3):
            ok    = s["ok"]
            total = s["total"]
            prec  = ok / total if total else 0
            req   = "Sí" if campo in CAMPOS_REQUERIDOS else "No"
            fill  = FILL_OK if prec >= 0.9 else (FILL_WARN if prec >= 0.6 else FILL_ERR)
            letra_prec = get_column_letter(5)
            letra_ok   = get_column_letter(3)
            letra_tot  = get_column_letter(4)

            fila = [campo, req, ok, total, None]
            for c_idx, valor in enumerate(fila, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=valor)
                cell.font = font; cell.fill = fill; cell.border = borde
                cell.alignment = Alignment(horizontal="left", vertical="center")

            # Precisión como fórmula
            prec_cell = ws.cell(row=r_idx, column=5)
            prec_cell.value  = f"={letra_ok}{r_idx}/{letra_tot}{r_idx}"
            prec_cell.number_format = "0.0%"
            prec_cell.font   = bold
            prec_cell.fill   = fill
            prec_cell.border = borde
            prec_cell.alignment = Alignment(horizontal="center")
            ws.row_dimensions[r_idx].height = 18

    def _excel_hoja_errores(self, ws):
        from openpyxl.styles import Alignment, Font, PatternFill
        borde, font, bold = self._excel_estilo_base()
        FILL_ERR = PatternFill("solid", start_color="FCE4D6")

        self._excel_titulo(ws, "Errores Detallados", 4)
        cols = [("Archivo", 35), ("Campo", 25), ("Esperado", 30), ("Obtenido / Motivo", 45)]
        self._excel_cabecera(ws, 2, cols)

        r_idx = 3
        for r in self.resultados:
            if r.error_extraccion:
                fila = [r.archivo, "ERROR DE EXTRACCIÓN", "—", r.error_extraccion]
                for c_idx, valor in enumerate(fila, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=valor)
                    cell.font = bold; cell.fill = FILL_ERR; cell.border = borde
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                ws.row_dimensions[r_idx].height = 18
                r_idx += 1
                continue

            fallos = [c for c in r.campos if not c.ok and c.motivo != "ignorado (null en GT)"]
            for c in fallos:
                fila = [r.archivo, c.campo, str(c.esperado), c.motivo]
                for c_idx, valor in enumerate(fila, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=valor)
                    cell.font = font; cell.fill = FILL_ERR; cell.border = borde
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                ws.row_dimensions[r_idx].height = 18
                r_idx += 1


# ─────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────

def main():
    logging.basicConfig(
        level=logging.WARNING,   # silenciar logs del extractor durante los tests
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    parser = argparse.ArgumentParser(
        description="FacturaIA — Suite de tests de precisión",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--dir", "-d",
        default=str(DIR_TEST_DEFAULT),
        help=f"Carpeta con facturas y ground truths (default: {DIR_TEST_DEFAULT})",
    )
    parser.add_argument(
        "--archivo", "-a",
        default=None,
        help="Testear solo un archivo concreto (nombre, no ruta completa)",
    )
    parser.add_argument(
        "--excel", "-e",
        action="store_true",
        help="Exportar informe a Excel en data/test/informe_tests.xlsx",
    )
    parser.add_argument(
        "--generar-gt",
        action="store_true",
        help=(
            "Generar borradores de ground truth para facturas sin .json.\n"
            "No lanza los tests. Revisa y corrige los borradores a mano."
        ),
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Mostrar logs del extractor durante los tests",
    )
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.INFO)

    dir_test = Path(args.dir)
    if not dir_test.exists():
        print(C.fail(f"\nError: la carpeta de tests no existe: {dir_test}"))
        print(
            f"Crea la carpeta y añade tus facturas con sus ground truths:\n"
            f"  mkdir -p {dir_test}\n"
            f"  cp mis_facturas/*.pdf {dir_test}/\n"
            f"  python test_extractor.py --generar-gt\n"
        )
        sys.exit(1)

    # ── Modo: generar ground truth ─────────────────────────────────────────
    if args.generar_gt:
        print(C.bold(f"\n  Generando borradores de ground truth en {dir_test}...\n"))
        GeneradorGroundTruth(dir_test).generar()
        sys.exit(0)

    # ── Modo: ejecutar tests ───────────────────────────────────────────────
    ejecutor = EjecutorTests(dir_test)
    pares    = ejecutor.descubrir_pares()

    if not pares:
        print(C.warn(
            f"\nNo se encontraron pares factura+ground_truth en {dir_test}.\n"
            f"Ejecuta primero:  python test_extractor.py --generar-gt\n"
        ))
        sys.exit(0)

    # Filtrar por archivo si se especificó
    if args.archivo:
        pares = [(f, gt) for f, gt in pares if f.name == args.archivo]
        if not pares:
            print(C.fail(f"\nArchivo '{args.archivo}' no encontrado en {dir_test}."))
            sys.exit(1)

    print(C.bold(f"\n  FacturaIA — Suite de Tests"))
    print(C.bold(f"  {len(pares)} factura(s) a testear en {dir_test}\n"))

    resultados = ejecutor.ejecutar_todos(pares)

    informe = InformeTests(resultados)
    informe.imprimir()

    if args.excel:
        ruta_excel = dir_test / f"informe_tests_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        informe.exportar_excel(ruta_excel)

    # Exit code: 0 si todo OK, 1 si hay fallos
    hay_fallos = any(not r.ok or r.error_extraccion for r in resultados)
    sys.exit(1 if hay_fallos else 0)


if __name__ == "__main__":
    main()
