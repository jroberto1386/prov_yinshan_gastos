# ─────────────────────────────────────────────
#  config.py  —  Pólizas de Provisión de Gastos
# ─────────────────────────────────────────────

RFC_RECEPTOR = "YIN080808FT6"   # RFC de la empresa receptora (Yinshan)

# ── Cuentas de IVA acreditable ────────────────
CTA_IVA_16   = 119010001
CTA_IVA_8    = 119010002

# ── Cuentas de retenciones ────────────────────
CTA_RET_IVA          = 216100000   # Impuestos retenidos de IVA
CTA_RET_ISR_GENERAL  = 216040000   # Ret ISR servicios profesionales / general
CTA_RET_ISR_ARRENDAM = 216030000   # Ret ISR arrendamiento

# ── Cuenta de IEPS (siempre la misma) ─────────
CTA_IEPS = 602860000

# ── Cuenta de gasto default (nuevos proveedores) ──
CTA_GASTO_DEFAULT = 501020000

# ── Palabras clave para detectar arrendamiento ─
PALABRAS_ARRENDAMIENTO = [
    "arrendamiento", "arrendam", "renta ", "rentas ", "renta de",
    "alquiler", "alquil", "subarrendamiento",
]

# ── Tipo de póliza y diario ───────────────────
TIPO_POL = 3
ID_DIARIO = 0
