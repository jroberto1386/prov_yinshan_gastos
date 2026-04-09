# ─────────────────────────────────────────────
#  motor.py  —  Provisión de Gastos / Proveedores
#  v2 — Bajo consumo de RAM para Render free tier
#
#  Cambios vs v1:
#  - Lee el Excel de facturas con openpyxl read_only (fila por fila)
#    en lugar de pd.read_excel() que carga todo en un DataFrame
#  - Genera el Excel de pólizas con write_only mode (sin mantener
#    el workbook completo en RAM)
#  - Un solo pase: parsea y escribe simultáneamente, sin lista
#    intermedia de 4,415 dicts acumulados
#  - pandas solo se usa para cargar el catálogo de proveedores
#    (~5,891 filas, manejable) — no para las facturas
# ─────────────────────────────────────────────

import io
import os
import subprocess
import tempfile
from datetime import datetime, date

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.cell import WriteOnlyCell

from config import (
    CTA_IVA_16, CTA_IVA_8, CTA_IEPS,
    CTA_RET_IVA, CTA_RET_ISR_GENERAL, CTA_RET_ISR_ARRENDAM,
    CTA_GASTO_DEFAULT, PALABRAS_ARRENDAMIENTO,
    TIPO_POL, ID_DIARIO,
)

# ── 22 filas de encabezado CONTPAQi ───────────
HEADERS = [
    ["Egreso(EG)", "IdDocumentoDe", "TipoDocumento", "Folio", "Fecha", "FechaAplicacion",
     "CodigoPersona", "BeneficiarioPagador", "IdCuentaCheques", "CodigoMoneda",
     "Total", "Referencia", "Origen", "BancoDestino", "CuentaDestino",
     "OtroMetodoDePago", "Guid", None, None, "TipoCambio",
     "UUIDRep", "NodoPago", "CodigoMonedaTipoCambio", "NumAsoc"],
    ["deposito.1(DE)", "IdDocumentoDe", "TipoDocumento", "Folio", "Fecha", "Ejercicio",
     "Periodo", "FechaAplicacion", "EjercicioAp", "PeriodoAp",
     "IdCuentaCheques", "NatBancaria", "Naturaleza", "Total", "Referencia",
     "Concepto", "EsConciliado", "IdMovEdoCta", "EjercicioPol", "PeriodoPol",
     "TipoPol", "NumPol", "FormaDeposito", "IdPoliza", "Origen",
     "IdDocumento", "PolizaAgrupada", "UsuarioCrea", "UsuarioModifica", "tieneCFD", "Guid"],
    ["ingreso.1(IN)", "IdDocumentoDe", "TipoDocumento", "Folio", "Fecha", "FechaAplicacion",
     "CodigoPersona", "BeneficiarioPagador", "IdCuentaCheques", "CodigoMoneda",
     "Total", "Referencia", "Origen", "BancoOrigen", "CuentaOrigen",
     "OtroMetodoDePago", "Guid", None, None, "TipoCambio",
     "NumeroCheque", "UUIDRep", "NodoPago", "CodigoMonedaTipoCambio", "NumAsoc"],
    ["Datos para CONTPAQi Factura Electrónica®(FE)", "RutaAnexo", "ArchivoAnexo"],
    ["Movimiento de póliza(M1)", "IdCuenta", "Referencia", "TipoMovto", "Importe",
     "IdDiario", "ImporteME", "Concepto", "IdSegNeg", "Guid", "FechaAplicacion"],
    ["Devolución de IVA (IETU)(W)", "IETUDeducible", "IETUModificado"],
    ["Devolución de IVA(V)", "IdProveedor", "ImpTotal", "PorIVA", "ImpBase",
     "ImpIVA", "CausaIVA", "ExentoIVA", "Serie", "Folio", "Referencia",
     "OtrosImptos", "ImpSinRet", "IVARetenido", "ISRRetenido", "GranTotal",
     "EjercicioAsignado", "PeriodoAsignado", "IdCuenta", "IVAPagNoAcred",
     "UUID", None, "IEPS"],
    ["Asociación de nodo de pago(AP)", "UUIDRep", "NumNodoPago", "GuidReferencia", "AplicationType"],
    ["Periodo de causación de IVA(R)", "EjercicioAsignado", "PeriodoAsignado"],
    ["Póliza(P)", "Fecha", "TipoPol", "Folio", "Clase", "IdDiario",
     "Concepto", "SistOrig", "Impresa", "Ajuste", "Guid"],
    ["Asociación movimiento(AM)", "UUID"],
    ["Comprobantes(MC)", "IdCuentaFlujoEfectivo", "IdSegmentoNegCtaFlujo", "Fecha",
     "Serie", "Folio", "UUID", "ClaveRastreo", "Referencia", "IdProveedor",
     "CodigoConceptoIETU", "ImpNeto", "ImpNetoME", "IdCuentaNeto",
     "IdSegmentoNegNeto", "PorIVA", "ImporteIVA", "ImporteIVAME",
     "IVATasaExcenta", "IdCuentaIVA", "IdSegmentoNegIVA", "NombreImpuesto",
     "ImpImpuesto", "ImpImpuestoME", "IdCuentaImpuesto", "IdSegmentoNegImp",
     "ImpOtrosGastos", "ImpOtrosGastosME", "IdCuentaOtrosGastos",
     "IdSegmentoNegOtrosGastos", "IVARetenido", "IVARetenidoME",
     "IdCuentaRetIVA", "IdSegmentoNegRetIVA", "ISRRetenido", "ISRRetenidoME",
     "IdCuentaRetISR", "IdSegmentoNegRetISR", "NombreOtrasRetenciones",
     "ImpOtrasRetenciones", "ImpOtrasRetencionesME", "IdCuentaOtrasRetenciones",
     "IdSegmentoNegOtrasRet", "BaseIVADIOT", "BaseIETU", "IVANoAcreditable",
     "ImpTotalErogacion", "IVAAcreditable", "ImpExtra1", "ImpExtra2",
     "IdCategoria", "IdSubCategoria", "TipoCambio", "IdDocGastos",
     "EsCapturaCompleta", "FolioStr"],
    ["Movimiento de póliza(M)", "IdCuenta", "Referencia", "TipoMovto", "Importe",
     "IdDiario", "ImporteME", "Concepto", "IdSegNeg"],
    ["Dispersiones de pago(DP)", "UUID", "UUIDRep", "GuidRef", "NumNodoPago",
     "FechaPago", "TotalPago", "TipoCambio", "TotalPagoComprobante"],
    ["Devolución de IVA (IETU)(W2)", "IETUDeducible", "IETUAcreditable",
     "IETUModificado", "IdConceptoIETU"],
    ["Movimientos de impuestos(I)", "IdPersona", "EjercicioAsignado",
     "PeriodoAsignado", "IdCuenta", "AplicaImpuesto", "Serie", "Folio",
     "Referencia", "UUID", "Origen", "Computable", "TipoMovimiento",
     "TipoFactor", "Impuesto", "ObjetoImpuesto", "NombreImpLocal",
     "TasaOCuota", "ImpBase", "ImpImpuesto", "ImpTotal", "Desglosado",
     "IVANoAcred", "AcumulaIETU", "IdConceptoIETU", "IETUDeducible",
     "IETUModificado", "IETUAcreditable", "GuidMov", "GuidMovPadre",
     "Migrado", "ConceptoIVA", "SubconceptoIVA", "ClasificadorIVA",
     "ProporcionDIOT", "DeducibleDIOT"],
    ["Asociación documento(AD)", "UUID"],
    ["Cheque(CH)", "IdDocumentoDe", "TipoDocumento", "Folio", "Fecha",
     "FechaAplicacion", "CodigoPersona", "BeneficiarioPagador",
     "IdCuentaCheques", "CodigoMoneda", "Total", "Referencia", "Origen",
     "CuentaDestino", "BancoDestino", "Guid", None, "OtroMetodoDePago",
     "TipoCambio", "UUIDRep", "NodoPago", "CodigoMonedaTipoCambio", "NumAsoc"],
    ["IngresosNoDepositados.1(DI)", "IdDocumentoDe", "TipoDocumento", "Folio",
     "Fecha", "Ejercicio", "Periodo", "FechaAplicacion", "EjercicioAp",
     "PeriodoAp", "CodigoPersona", "BeneficiarioPagador", "NatBancaria",
     "Naturaleza", "CodigoMoneda", "CodigoMonedaTipoCambio", "TipoCambio",
     "Total", "Referencia", "Concepto", "EsAsociado", "UsuAutorizaPresupuesto",
     "PosibilidadPago", "EsProyectado", "Origen", "IdChequeOrigen",
     "TipoCambioDeposito", "IdDocumento", "EsAnticipo", "EsTraspasado",
     "UsuarioCrea", "UsuarioModifica", "tieneCFD", "Guid",
     "CuentaOrigen", "BancoOrigen", "OtroMetodoDePago", "NumeroCheque", "NumAsoc"],
    ["Causación de IVA (Concepto de IETU)(E)", "IdConceptoIETU"],
    ["Causación de IVA (IETU)(D)", "IVATasa15NoAcred", "IVATasa10NoAcred",
     "IETU", "Modificado", "Origen", "TotTasa16", "BaseTasa16", "IVATasa16",
     "IVATasa16NoAcred", "TotTasa11", "BaseTasa11", "IVATasa11",
     "IVATasa11NoAcred", "TotTasa8", "BaseTasa8", "IVATasa8", "IVATasa8NoAcred"],
    ["Causación de IVA(C)", "Tipo", "TotTasa15", "BaseTasa15", "IVATasa15",
     "TotTasa10", "BaseTasa10", "IVATasa10", "TotTasa0", "BaseTasa0",
     "TotTasaExento", "BaseTasaExento", "TotOtraTasa", "BaseOtraTasa",
     "IVAOtraTasa", "ISRRetenido", "TotOtros", "IVARetenido",
     "Captado", "NoCausar", "IEPS"],
]

# ── Encabezados catálogo de cuentas ───────────
_HEADERS_CUENTAS = [
    ["Grupo estadístico(E)", "CtaSup"],
    ["Cuenta contable(C)", "Codigo", "Nombre", "NomIdioma", "CtaSup", "Tipo",
     "EsBaja", "CtaMayor", "CtaEfectivo", "FechaRegistro", "SistOrigen",
     "IdMoneda", "DigAgrup", "IdSegNeg", "SegNegMovtos", "Consume", "IdAgrupadorSAT"],
    ["Rubros NIF(RF)", "Codigo"],
    ["Cuenta de flujo de efectivo(F)", "Codigo"],
    ["F", "10200000"],
    ["C", "200000000", "Pasivo", "Liabilities", "000000000", "D", 0, 2, 0, "2019-11-11", 11, "   1", 0, "0", 0, 0, "0"],
    ["C", "200010000", "Pasivo a corto plazo", "Short-term liabilities", "200000000", "D", 0, 3, 0, "2019-11-11", 11, "   1", 0, "0", 0, 0, "0"],
    ["C", "201000000", "Proveedores", "Suppliers", "200010000", "D", 0, 1, 0, "2019-11-11", 11, "   1", 0, "0", 0, 0, "201"],
    ["C", "201010000", "Proveedores nacionales", "Nationals Suppliers", "201000000", "D", 0, 2, 0, "2019-11-11", 11, "   1", 0, "0", 0, 0, "201.01"],
]

_CUENTA_FIJA = {
    "cta_sup": "201010000", "tipo": "D", "es_baja": 0,
    "cta_mayor": 2, "cta_efectivo": 0, "sist_origen": 11,
    "id_moneda": "   1", "dig_agrup": 0, "id_seg_neg": "0",
    "seg_neg_movtos": 0, "consume": 0, "id_agrup_sat": "201.01",
}


# ── Helpers ───────────────────────────────────

def _f(val, default=0.0):
    if val is None:
        return default
    try:
        v = float(val)
        import math
        return default if math.isnan(v) else v
    except (TypeError, ValueError):
        return default


def _s(val, default=""):
    if val is None:
        return default
    s = str(val).strip()
    return default if s.lower() in ("nan", "none", "") else s


def _es_arrendamiento(texto):
    t = texto.lower()
    return any(p in t for p in PALABRAS_ARRENDAMIENTO)


def _xls_a_xlsx_bytes(raw_bytes):
    """Convierte bytes de .xls a .xlsx usando LibreOffice. Limpia archivos temporales."""
    with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as tmp:
        tmp.write(raw_bytes)
        tmp_path = tmp.name
    out_dir = tempfile.mkdtemp()
    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "xlsx",
             "--outdir", out_dir, tmp_path],
            capture_output=True, timeout=90,
        )
        xlsx_path = os.path.join(
            out_dir, os.path.basename(tmp_path).replace(".xls", ".xlsx")
        )
        with open(xlsx_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        xlsx_path_try = os.path.join(
            out_dir, os.path.basename(tmp_path).replace(".xls", ".xlsx")
        )
        if os.path.exists(xlsx_path_try):
            os.unlink(xlsx_path_try)
        if os.path.exists(out_dir):
            os.rmdir(out_dir)


# ── Carga del catálogo de proveedores ─────────
# pandas solo aquí — el catálogo es pequeño (~6K filas) y manejable.

def cargar_catalogo(catalogo_bytes):
    """
    Devuelve:
      - catalogo: dict RFC → {cta_proveedor, cta_gasto, nombre, codigo}
      - ultimo_codigo: int
      - ultima_cuenta: int (última cuenta secuencial 201010001-201016000)
      - plantilla_row: list (primera fila P1, para generar_altas)
      - header_rows: list de lists (primeras 6 filas del catálogo, para generar_altas)
    """
    try:
        df = pd.read_excel(io.BytesIO(catalogo_bytes), header=None, engine="openpyxl")
    except Exception:
        xlsx_bytes = _xls_a_xlsx_bytes(catalogo_bytes)
        df = pd.read_excel(io.BytesIO(xlsx_bytes), header=None, engine="openpyxl")

    provs = df[df[0] == "P1"].copy()

    catalogo = {}
    for _, r in provs.iterrows():
        rfc = _s(r[3])
        if not rfc:
            continue
        try:
            cta_prov = int(float(r[6]))
        except (ValueError, TypeError):
            cta_prov = None
        try:
            cta_gasto = int(float(r[26])) if not pd.isna(r[26]) else CTA_GASTO_DEFAULT
        except (ValueError, TypeError):
            cta_gasto = CTA_GASTO_DEFAULT
        try:
            codigo = int(float(r[1]))
        except (ValueError, TypeError):
            codigo = 0

        catalogo[rfc] = {
            "nombre": _s(r[2]),
            "codigo": codigo,
            "cta_proveedor": cta_prov,
            "cta_gasto": cta_gasto,
        }

    codigos = [v["codigo"] for v in catalogo.values() if v["codigo"] > 0]
    ultimo_codigo = max(codigos) if codigos else 0

    provs2 = provs.copy()
    provs2["_codigo"] = pd.to_numeric(provs2[1], errors="coerce")
    provs2["_cuenta"] = pd.to_numeric(provs2[6], errors="coerce")
    provs_seq = provs2.dropna(subset=["_codigo", "_cuenta"])
    provs_seq = provs_seq[
        (provs_seq["_cuenta"] >= 201010001) &
        (provs_seq["_cuenta"] <= 201016000)
    ].sort_values("_codigo", ascending=False)
    ultima_cuenta = int(provs_seq.iloc[0]["_cuenta"]) if len(provs_seq) else 201015882

    # Plantilla (primera fila P1) para generar_altas
    plantilla_row = provs.iloc[0].tolist() if len(provs) else []

    # Primeras 6 filas del catálogo para encabezados de altas
    header_rows = [df.iloc[i].tolist() for i in range(min(6, len(df)))]

    # Liberar el DataFrame grande
    del df, provs, provs2, provs_seq

    return catalogo, ultimo_codigo, ultima_cuenta, plantilla_row, header_rows


# ── Procesamiento + generación en un solo pase ─
#
# ESTRATEGIA DE MEMORIA:
#   1. Leer facturas con openpyxl read_only → fila por fila, sin DataFrame
#   2. Escribir pólizas con openpyxl write_only → sin mantener workbook en RAM
#   3. Un solo pase simultáneo: no se acumula la lista de 4K dicts
#   4. nuevos_provs_dict es pequeño (~42 entradas) — sí se acumula, no hay problema

def procesar_excel(excel_bytes, catalogo_bytes, num_poliza_inicio, callback=None):
    """
    Lee facturas y genera el Excel de pólizas en un solo pase.
    Devuelve (polizas_bytes, nuevos_provs_list, stats, ultimo_codigo).
    """
    catalogo, ultimo_codigo, ultima_cuenta, plantilla_row, header_rows = \
        cargar_catalogo(catalogo_bytes)

    # Contadores para stats (no acumulamos dicts)
    cnt = {
        "procesadas": 0, "facturas": 0, "notas_credito": 0,
        "con_iva8": 0, "con_ret_iva": 0, "con_ret_isr": 0,
        "arrendamiento": 0, "proveedores_nuevos_set": set(),
    }

    nuevos_provs_dict = {}  # RFC → {nombre, rfc, cta_proveedor}

    def obtener_cuenta_nueva(rfc, nombre):
        if rfc not in nuevos_provs_dict:
            nueva_cta = ultima_cuenta + len(nuevos_provs_dict) + 1
            nuevos_provs_dict[rfc] = {
                "nombre": nombre, "rfc": rfc, "cta_proveedor": nueva_cta,
            }
        return nuevos_provs_dict[rfc]["cta_proveedor"]

    # ── Abrir Excel de facturas en modo read_only ──
    wb_in = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    ws_in = wb_in.active
    rows_iter = ws_in.iter_rows(values_only=True)
    next(rows_iter)  # saltar fila de headers

    # Contar total aproximado para callback (no usamos max_row en read_only — puede ser None)
    # Usamos un contador incremental en su lugar
    total_procesado = 0

    # ── Abrir Excel de salida en modo write_only ──
    wb_out = openpyxl.Workbook(write_only=True)
    ws_out = wb_out.create_sheet("Pólizas Gastos")

    hfill = PatternFill("solid", fgColor="D9D9D9")
    hfont = Font(bold=True, size=9)

    # Escribir 22 filas de encabezado
    for hrow in HEADERS:
        cells = []
        for v in hrow:
            c = WriteOnlyCell(ws_out, value=v)
            c.font = hfont
            c.fill = hfill
            cells.append(c)
        ws_out.append(cells)

    num_pol = num_poliza_inicio

    # ── Pase único: leer fila → procesar → escribir → siguiente ──
    for row in rows_iter:
        total_procesado += 1
        if callback:
            callback(total_procesado, total_procesado)  # total desconocido en read_only

        tipo = _s(row[3])
        if tipo not in ("Factura", "NotaCredito"):
            continue

        es_nota = (tipo == "NotaCredito")
        signo   = -1 if es_nota else 1

        rfc_emisor   = _s(row[12])
        nombre       = _s(row[13])
        serie        = _s(row[8])
        folio        = _s(row[9])
        uuid         = _s(row[10]).upper()
        concepto_txt = _s(row[40])

        # Fecha — siempre datetime real con yyyymmdd
        fecha_raw = row[4]
        try:
            if isinstance(fecha_raw, datetime):
                fecha = fecha_raw
            else:
                fecha = datetime.strptime(str(fecha_raw)[:10], "%Y-%m-%d")
        except (ValueError, TypeError):
            try:
                fecha = datetime.strptime(str(fecha_raw)[:10], "%d/%m/%Y")
            except Exception:
                fecha = None

        subtotal    = _f(row[20])
        descuento   = _f(row[21])
        iva16_col   = _f(row[23])
        ret_iva     = _f(row[24])
        ret_isr     = _f(row[25])
        ish         = _f(row[26])
        total_cfdi  = _f(row[27])
        tot_traslad = _f(row[29])
        iva8        = _f(row[56])

        iva_declarado = round(iva16_col + iva8, 2)
        if abs(iva_declarado - tot_traslad) > 0.05 and tot_traslad > 0:
            iva16 = round(tot_traslad - iva8, 2)
        else:
            iva16 = iva16_col

        gasto_neto = round(subtotal - descuento + ish, 2)
        neto_prov  = round(total_cfdi, 2)

        folio_limpio = folio if folio and folio.lower() not in ("", "nan") else ""
        serie_limpia = serie if serie and serie.lower() not in ("", "nan") else ""
        if folio_limpio:
            ref = f"F-{serie_limpia}{folio_limpio}".strip()
        else:
            uuid_sg = uuid.replace("-", "")
            ref = f"F-{uuid_sg[-5:]}"

        prov = catalogo.get(rfc_emisor)
        if prov:
            cta_prov        = prov["cta_proveedor"]
            cta_gasto       = prov["cta_gasto"]
            nombre_catalogo = prov["nombre"] or nombre
        else:
            cta_prov        = obtener_cuenta_nueva(rfc_emisor, nombre)
            cta_gasto       = CTA_GASTO_DEFAULT
            nombre_catalogo = nombre

        arr         = _es_arrendamiento(concepto_txt)
        cta_ret_isr = CTA_RET_ISR_ARRENDAM if arr else CTA_RET_ISR_GENERAL
        concepto_p  = f"Provision {nombre_catalogo} {ref}"

        gasto_neto_s = round(signo * gasto_neto, 2)
        iva16_s      = round(signo * iva16, 2)
        iva8_s       = round(signo * iva8, 2)
        ret_iva_s    = round(signo * ret_iva, 2)
        ret_isr_s    = round(signo * ret_isr, 2)
        neto_prov_s  = round(signo * neto_prov, 2)

        # ── Fila P — write_only con fecha yyyymmdd ──
        c_p = WriteOnlyCell(ws_out, value="P")
        c_fecha = WriteOnlyCell(ws_out, value=fecha)
        if isinstance(fecha, datetime):
            c_fecha.number_format = "yyyymmdd"
        ws_out.append([
            c_p, c_fecha,
            int(TIPO_POL), int(num_pol), 1, int(ID_DIARIO),
            concepto_p, 12, 0, 0,
        ])

        # ── M1 Gasto (Debe) ──
        ws_out.append(["M1", int(cta_gasto), ref, 0, gasto_neto_s, int(ID_DIARIO), 0, nombre_catalogo])

        # ── M1 IVA 16% (Debe) ──
        if abs(iva16_s) > 0:
            ws_out.append(["M1", int(CTA_IVA_16), ref, 0, iva16_s, int(ID_DIARIO), 0, nombre_catalogo])

        # ── M1 IVA 8% (Debe) ──
        if abs(iva8_s) > 0:
            ws_out.append(["M1", int(CTA_IVA_8), ref, 0, iva8_s, int(ID_DIARIO), 0, nombre_catalogo])

        # ── M1 Retención IVA (Haber) ──
        if abs(ret_iva_s) > 0:
            ws_out.append(["M1", int(CTA_RET_IVA), ref, 1, abs(ret_iva_s), int(ID_DIARIO), 0, nombre_catalogo])

        # ── M1 Retención ISR (Haber) ──
        if abs(ret_isr_s) > 0:
            ws_out.append(["M1", int(cta_ret_isr), ref, 1, abs(ret_isr_s), int(ID_DIARIO), 0, nombre_catalogo])

        # ── M1 Proveedor (Haber) ──
        ws_out.append(["M1", int(cta_prov), ref, 1, abs(neto_prov_s), int(ID_DIARIO), 0, nombre_catalogo])

        # ── AD UUID ──
        if uuid:
            ws_out.append(["AD", uuid])

        num_pol += 1

        # Contadores
        cnt["procesadas"] += 1
        if es_nota:
            cnt["notas_credito"] += 1
        else:
            cnt["facturas"] += 1
        if abs(iva8_s) > 0:
            cnt["con_iva8"] += 1
        if abs(ret_iva_s) > 0:
            cnt["con_ret_iva"] += 1
        if abs(ret_isr_s) > 0:
            cnt["con_ret_isr"] += 1
        if arr:
            cnt["arrendamiento"] += 1
        if prov is None:
            cnt["proveedores_nuevos_set"].add(rfc_emisor)

    wb_in.close()

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    polizas_bytes = buf.getvalue()

    stats = {
        "procesadas":        cnt["procesadas"],
        "facturas":          cnt["facturas"],
        "notas_credito":     cnt["notas_credito"],
        "con_iva8":          cnt["con_iva8"],
        "con_ret_iva":       cnt["con_ret_iva"],
        "con_ret_isr":       cnt["con_ret_isr"],
        "arrendamiento":     cnt["arrendamiento"],
        "proveedores_nuevos": len(nuevos_provs_dict),
        "polizas_desde":     num_poliza_inicio,
        "polizas_hasta":     num_pol - 1,
    }

    return polizas_bytes, list(nuevos_provs_dict.values()), stats, ultimo_codigo, \
           plantilla_row, header_rows


# ── Alta de Catálogo de Cuentas ───────────────

def generar_catalogo_cuentas(nuevos_provs):
    fecha_hoy = datetime.combine(date.today(), datetime.min.time())

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catalogo Cuentas"

    hfill = PatternFill("solid", fgColor="1F4E79")
    hfont = Font(bold=True, color="FFFFFF", size=9)
    cfill = PatternFill("solid", fgColor="DDEEFF")
    rfill = PatternFill("solid", fgColor="F2F2F2")
    cfont = Font(size=9)

    for ri, hrow in enumerate(_HEADERS_CUENTAS, 1):
        for ci, val in enumerate(hrow, 1):
            c = ws.cell(ri, ci, val)
            c.font = hfont
            c.fill = hfill

    row_idx = 10
    for prov in nuevos_provs:
        cuenta_vals = [
            "C", str(prov["cta_proveedor"]), prov["nombre"], None,
            _CUENTA_FIJA["cta_sup"], _CUENTA_FIJA["tipo"], _CUENTA_FIJA["es_baja"],
            _CUENTA_FIJA["cta_mayor"], _CUENTA_FIJA["cta_efectivo"],
            fecha_hoy,
            _CUENTA_FIJA["sist_origen"], _CUENTA_FIJA["id_moneda"],
            _CUENTA_FIJA["dig_agrup"], _CUENTA_FIJA["id_seg_neg"],
            _CUENTA_FIJA["seg_neg_movtos"], _CUENTA_FIJA["consume"],
            _CUENTA_FIJA["id_agrup_sat"],
        ]
        for ci, val in enumerate(cuenta_vals, 1):
            c = ws.cell(row_idx, ci, val)
            c.fill = cfill
            c.font = cfont
            if ci == 10:
                c.number_format = "yyyymmdd"
        row_idx += 1

        ws.cell(row_idx, 1, "RF").fill = rfill
        ws.cell(row_idx, 1).font = cfont
        ws.cell(row_idx, 2, "2102.03").fill = rfill
        ws.cell(row_idx, 2).font = cfont
        row_idx += 1

    anchos = [6, 12, 45, 4, 12, 5, 6, 8, 10, 12, 10, 8, 8, 8, 12, 8, 10]
    for ci, ancho in enumerate(anchos, 1):
        ws.column_dimensions[ws.cell(1, ci).column_letter].width = ancho

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Alta de Nuevos Proveedores ────────────────

def generar_altas(nuevos_provs, ultimo_codigo, plantilla_row, header_rows):
    """
    Recibe plantilla_row y header_rows precargados en cargar_catalogo()
    para no volver a leer el catálogo completo desde bytes.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nuevos Proveedores"

    dfill = PatternFill("solid", fgColor="D6E4F0")
    wfill = PatternFill("solid", fgColor="FFF2CC")
    wfont = Font(size=9)
    nfont = Font(bold=True, color="C00000", size=9)

    for row_i, src_row in enumerate(header_rows, 1):
        for col_i, val in enumerate(src_row, 1):
            v = None if (isinstance(val, float) and pd.isna(val)) else val
            c = ws.cell(row_i, col_i, v)
            c.font = Font(bold=True, size=9)
            c.fill = dfill

    for i, prov in enumerate(nuevos_provs):
        nuevo_codigo = ultimo_codigo + i + 1
        data_row = [None] * max(30, len(plantilla_row))

        for ci, v in enumerate(plantilla_row):
            if not (isinstance(v, float) and pd.isna(v)):
                data_row[ci] = v

        data_row[0]  = "P1"
        data_row[1]  = nuevo_codigo
        data_row[2]  = prov["nombre"]
        data_row[3]  = prov["rfc"]
        data_row[4]  = ""
        data_row[6]  = prov["cta_proveedor"]
        data_row[26] = CTA_GASTO_DEFAULT

        excel_row = 7 + i
        for ci, v in enumerate(data_row, 1):
            c = ws.cell(excel_row, ci, v)
            c.fill = wfill
            c.font = wfont

        nota_col = max(30, len(plantilla_row)) + 1
        c = ws.cell(excel_row, nota_col, "⚠ VERIFICAR cuenta de gasto (default 501020000)")
        c.font = nfont

    ws.cell(1, max(30, len(plantilla_row)) + 1, "Nota").font = Font(bold=True, size=9)

    for col in ws.columns:
        max_w = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max_w + 2, 50)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
