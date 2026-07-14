#!/usr/bin/env python3
"""
GRUPO PANABLOCK — Cierre Diario de Caja — Módulo central
Sin dependencias de GUI. Importado por:
  • app.py        (interfaz Streamlit / web)
  • cierre_caja.py (interfaz tkinter / escritorio)
"""

import re
import io
import os
import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ════════════════════════════════════════════════════════════════════
# ESTILOS EXCEL
# ════════════════════════════════════════════════════════════════════
DARK_BLUE  = '1F4E79'
LIGHT_BLUE = 'BDD7EE'
LGRAY      = 'D9D9D9'
MGRAY      = 'BFBFBF'
RED_FILL   = 'FF0000'
WHITE      = 'FFFFFF'
CURRENCY   = '#,##0.00'

_thin = Side(border_style='thin', color='000000')
BRD   = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
CA    = Alignment(horizontal='center', vertical='center')
LA    = Alignment(horizontal='left',   vertical='center', wrap_text=True)
RA    = Alignment(horizontal='right',  vertical='center')

def mk_fill(hex_c):
    return PatternFill(start_color=hex_c, end_color=hex_c, fill_type='solid')

NO_BORDER = Border()

# ════════════════════════════════════════════════════════════════════
# CLASIFICADOR DE GASTOS
# ════════════════════════════════════════════════════════════════════

# Mapa de proveedores cargado desde proveedores.xlsx.
# Se asigna desde fuera: core._PROV_MAP = {...}
# Clave: keyword en MAYÚSCULAS. Valor: categoría.
_PROV_MAP = {}

# Keywords genéricas como segunda línea de defensa
GASTO_KEYWORDS = [
    ('Materia Prima',           ['cemento', 'arena', 'agregado', 'bloque',
                                  'hormigon', 'hormigón', 'material prima']),
    ('Pigmento y Aditivos',     ['pigmento', 'aditivo', 'colorante',
                                  'oxido', 'óxido']),
    ('Mantenimiento de Planta', ['mantenimiento', 'reparacion', 'reparación',
                                  'repuesto', 'herramienta']),
    ('Servicios Básicos',       ['agua', 'luz', 'electricidad', 'energia',
                                  'energía', 'internet', 'telefono', 'teléfono',
                                  'cable', 'gas', 'assa', 'naturgy',
                                  'claro', 'tigo', 'cwp']),
    ('Tarimas',                 ['tarima', 'pallet']),
    ('Transporte',              ['transporte', 'flete', 'envio', 'envío',
                                  'courier', 'acarreo', 'traslado']),
    ('Acreedores',              ['proveedor', 'acreedor', 'pago a proveedor']),
]

ITEMS_PAGOS = [
    'Acreedores', 'Mantenimiento de Planta', 'Materia Prima',
    'Pigmento y Aditivos', 'Servicios Básicos', 'Tarimas',
    'Transporte', 'Otros',
]

def _classify_gasto(desc_lower):
    """
    Clasifica un gasto por descripción.
    Orden de búsqueda:
      1. proveedores.xlsx  (keywords exactas, comparación en MAYÚSCULAS)
      2. GASTO_KEYWORDS    (palabras clave genéricas)
      3. 'Otros'           (fallback)
    """
    desc_upper = desc_lower.upper()
    for keyword, cat in _PROV_MAP.items():
        if keyword in desc_upper:
            return cat
    for cat, kws in GASTO_KEYWORDS:
        if any(k in desc_lower for k in kws):
            return cat
    return 'Otros'

def _merge_gastos(bg_gastos, gb_gastos):
    """Suma los gastos de Banco General y Global Bank por categoría."""
    merged = {}
    for c in set(list(bg_gastos.keys()) + list(gb_gastos.keys())):
        merged[c] = round(bg_gastos.get(c, 0) + gb_gastos.get(c, 0), 2)
    return merged

# ════════════════════════════════════════════════════════════════════
# HELPERS DE LECTURA DE ARCHIVOS
# ════════════════════════════════════════════════════════════════════

def _to_bytes(path_or_file):
    """
    Acepta:
      • ruta de archivo (str / Path)  → abre y lee
      • file-like object (BytesIO, UploadedFile de Streamlit) → lee bytes
    Retorna bytes para poder crear múltiples BytesIO independientes
    (necesario cuando hay que leer el mismo archivo dos veces, ej. Global Bank).
    """
    if isinstance(path_or_file, (str, os.PathLike)):
        with open(path_or_file, 'rb') as f:
            return f.read()
    data = path_or_file.read()
    if hasattr(path_or_file, 'seek'):
        path_or_file.seek(0)   # deja el stream listo para otro uso
    return data

def _float(s):
    """Convierte '$1,234.56'  o  '1234.56'  a float."""
    try:
        return float(str(s).replace('$', '').replace(',', '').strip())
    except Exception:
        return 0.0

# ════════════════════════════════════════════════════════════════════
# PARSERS
# ════════════════════════════════════════════════════════════════════

# ── Facturas PDF ─────────────────────────────────────────────────
def parse_facturas(path, log=print):
    """
    Acepta un solo PDF o una lista de PDFs (ej. una factura por archivo).
    Retorna dict con:
      contado_subtotal, contado_itbms,
      credito_subtotal, credito_itbms,
      notas_credito: [{nc, factura, cliente, motivo, subtotal, itbms, total}]
    """
    # Normalizar: siempre trabajar con lista
    paths = path if isinstance(path, list) else [path]

    facturas   = []
    notas_cred = []

    for p in paths:
        data = _to_bytes(p)
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ''

                is_nc = bool(re.search(r'NOTA\s+CR[EÉ]DITO', text, re.IGNORECASE))

                num_m   = re.search(r'N[uú]mero:\s*0*(\d+)',      text)
                sub_m   = re.search(r'Subtotal:\s*([\d,]+\.?\d*)', text)
                itb_m   = re.search(r'ITBMS:\s*([\d,]+\.?\d*)',    text)
                tot_m   = re.search(r'^Total:\s*([\d,]+\.?\d*)',    text, re.MULTILINE)

                numero   = num_m.group(1)         if num_m  else '?'
                subtotal = _float(sub_m.group(1)) if sub_m  else 0.0
                itbms    = _float(itb_m.group(1)) if itb_m  else 0.0
                total    = _float(tot_m.group(1)) if tot_m  else subtotal + itbms

                if is_nc:
                    fact_ref = '?'
                    app_m = re.search(r'Aplicado\s+A:\s*(\S+)', text, re.IGNORECASE)
                    if app_m:
                        cufe  = app_m.group(1)
                        ref_m = re.search(r'20\d{6}(\d{10})001', cufe)
                        if ref_m:
                            fact_ref = str(int(ref_m.group(1)))

                    com_m  = re.search(r'Comentario:\s*(.+)', text)
                    motivo = com_m.group(1).strip() if com_m else ''

                    cli_m   = re.search(r'DV\s+77\s+(.+)', text)
                    cliente = cli_m.group(1).strip() if cli_m else ''

                    notas_cred.append({
                        'nc':       numero,
                        'factura':  fact_ref,
                        'cliente':  cliente,
                        'motivo':   motivo,
                        'subtotal': subtotal,
                        'itbms':    itbms,
                        'total':    total,
                    })
                    log(f"  NC {numero} → Factura {fact_ref} — {cliente}")
                else:
                    is_contado = bool(re.search(r'Consumidor\s+Final', text, re.IGNORECASE))
                    tipo = 'CONTADO' if is_contado else 'CREDITO'
                    facturas.append({
                        'numero':   numero,
                        'tipo':     tipo,
                        'subtotal': subtotal,
                        'itbms':    itbms,
                        'total':    total,
                    })
                    log(f"  Factura {numero} [{tipo}] Sub={subtotal:.2f}  ITBMS={itbms:.2f}")

    nc_refs = {nc['factura'] for nc in notas_cred}

    contado_sub = sum(f['subtotal'] for f in facturas
                      if f['tipo'] == 'CONTADO' and f['numero'] not in nc_refs)
    contado_itb = sum(f['itbms']    for f in facturas
                      if f['tipo'] == 'CONTADO' and f['numero'] not in nc_refs)
    credito_sub = sum(f['subtotal'] for f in facturas
                      if f['tipo'] == 'CREDITO' and f['numero'] not in nc_refs)
    credito_itb = sum(f['itbms']    for f in facturas
                      if f['tipo'] == 'CREDITO' and f['numero'] not in nc_refs)

    return {
        'contado_subtotal': round(contado_sub, 2),
        'contado_itbms':    round(contado_itb, 2),
        'credito_subtotal': round(credito_sub, 2),
        'credito_itbms':    round(credito_itb, 2),
        'notas_credito':    notas_cred,
    }

# ── Recibos PDF ──────────────────────────────────────────────────
def parse_recibos(path, log=print):
    """
    Acepta un solo PDF o una lista de PDFs.
    Retorna total de cobros a clientes (suma de todos los recibos).
    """
    paths = path if isinstance(path, list) else [path]
    total = 0.0

    for p in paths:
        data = _to_bytes(p)
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = '\n'.join(page.extract_text() or '' for page in pdf.pages)

        tot_m = re.search(r'Totales?:\s*([\d,]+\.\d{2})', text, re.IGNORECASE)
        if tot_m:
            subtotal = _float(tot_m.group(1))
            log(f"  Recibo: total cobros ${subtotal:,.2f}")
            total += subtotal
        else:
            # Fallback: sumar columna de totales individuales
            items = re.findall(r'\b(\d{1,3}(?:,\d{3})*\.\d{2})\s+0\.00', text)
            subtotal = sum(_float(t) for t in items)
            log(f"  Recibo: total cobros (suma individual) ${subtotal:,.2f}")
            total += subtotal

    log(f"  Total cobros acumulado: ${total:,.2f}")
    return round(total, 2)

# ── Banco General PDF ────────────────────────────────────────────
def parse_banco_general(path, log=print):
    """
    Retorna:
      saldo_anterior, depositos,
      gastos: {categoria: monto},
      gastos_detalle: [{desc, monto, cat}]
    """
    if path is None:
        log("  Sin extracto Banco General — se omite.")
        return {'saldo_anterior': 0, 'depositos': 0,
                'gastos': {c: 0.0 for c in ITEMS_PAGOS}, 'gastos_detalle': []}
    data = _to_bytes(path)
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text = '\n'.join(page.extract_text() or '' for page in pdf.pages)

    pairs = re.findall(r'(-?\$[\d,]+\.\d{2})\s+\$([\d,]+\.\d{2})', text)
    if not pairs:
        log("  ADVERTENCIA: no se encontraron transacciones en extracto BG")
        return {'saldo_anterior': 0, 'depositos': 0,
                'gastos': {c: 0.0 for c in ITEMS_PAGOS}, 'gastos_detalle': []}

    amounts = [_float(p[0]) for p in pairs]
    saldos  = [_float(p[1]) for p in pairs]

    saldo_anterior = round(saldos[0] - amounts[0], 2)
    depositos      = round(sum(a for a in amounts if a > 0), 2)

    gastos_por_cat = {c: 0.0 for c in ITEMS_PAGOS}
    gastos_detalle = []
    neg_pattern = re.compile(r'(-\$[\d,]+\.\d{2})\s+\$[\d,]+\.\d{2}')

    for line in text.split('\n'):
        m = neg_pattern.search(line)
        if m:
            monto = abs(_float(m.group(1)))
            desc  = line[:m.start()].strip()
            cat   = _classify_gasto(desc.lower())
            gastos_por_cat[cat] = round(gastos_por_cat.get(cat, 0) + monto, 2)
            gastos_detalle.append({'desc': desc, 'monto': monto, 'cat': cat})
            log(f"  BG gasto [{cat}]: ${monto:.2f} — {desc[:60]}")

    log(f"  BG saldo anterior: ${saldo_anterior:,.2f}  depósitos: ${depositos:,.2f}")
    return {
        'saldo_anterior': saldo_anterior,
        'depositos':      depositos,
        'gastos':         gastos_por_cat,
        'gastos_detalle': gastos_detalle,
    }

# ── Global Bank XLS ──────────────────────────────────────────────
def parse_global_bank(path, log=print):
    """
    Retorna:
      saldo_anterior, depositos,
      gastos: {categoria: monto},
      gastos_detalle: [{desc, monto, cat}]
    """
    if path is None:
        log("  Sin extracto Global Bank — se omite.")
        return {'saldo_anterior': 0, 'depositos': 0,
                'gastos': {c: 0.0 for c in ITEMS_PAGOS}, 'gastos_detalle': []}
    # Leer bytes una vez, crear dos BytesIO independientes para las dos pasadas
    raw = _to_bytes(path)

    df = pd.read_excel(io.BytesIO(raw), engine='xlrd', header=None)

    # Encontrar la fila de encabezado (contiene "Débitos" / "Créditos")
    hdr_row = None
    for i, row in df.iterrows():
        vals = [str(v).strip() for v in row.values]
        if any('bito' in v.lower() for v in vals):
            hdr_row = i
            break

    if hdr_row is None:
        log("  ADVERTENCIA: no se encontró encabezado en Global Bank XLS")
        return {'saldo_anterior': 0, 'depositos': 0,
                'gastos': {c: 0.0 for c in ITEMS_PAGOS}, 'gastos_detalle': []}

    # Segunda pasada con encabezado correcto
    df2 = pd.read_excel(io.BytesIO(raw), engine='xlrd', header=hdr_row)
    df2.columns = [str(c).strip() for c in df2.columns]

    def find_col(keywords):
        for c in df2.columns:
            if any(k in c.lower() for k in keywords):
                return c
        return None

    col_debito  = find_col(['bito', 'débito', 'cargo'])
    col_credito = find_col(['crédito', 'credito', 'abono'])
    col_saldo   = find_col(['saldo'])
    col_desc    = find_col(['concepto', 'descripci', 'detalle'])

    if not all([col_debito, col_credito, col_saldo]):
        log("  ADVERTENCIA: columnas no encontradas en Global Bank XLS")
        return {'saldo_anterior': 0, 'depositos': 0,
                'gastos': {c: 0.0 for c in ITEMS_PAGOS}, 'gastos_detalle': []}

    data = df2.dropna(subset=[col_saldo]).copy()
    data = data[pd.to_numeric(data[col_saldo], errors='coerce').notna()]

    if data.empty:
        return {'saldo_anterior': 0, 'depositos': 0,
                'gastos': {c: 0.0 for c in ITEMS_PAGOS}, 'gastos_detalle': []}

    debitos  = pd.to_numeric(data[col_debito],  errors='coerce').fillna(0)
    creditos = pd.to_numeric(data[col_credito], errors='coerce').fillna(0)
    saldos   = pd.to_numeric(data[col_saldo],   errors='coerce').fillna(0)

    total_debitos  = round(float(debitos.sum()),  2)
    total_creditos = round(float(creditos.sum()), 2)

    # Detectar orden cronológico vs. más-reciente-primero con identidad contable
    def _sfloat(v):
        f = pd.to_numeric(v, errors='coerce')
        return 0.0 if (f != f) else float(f)

    data_r = data.reset_index(drop=True)
    oldest_first_score = newest_first_score = 0
    for i in range(len(data_r) - 1):
        s0  = _sfloat(data_r.iloc[i][col_saldo])
        s1  = _sfloat(data_r.iloc[i+1][col_saldo])
        cr1 = _sfloat(data_r.iloc[i+1][col_credito])
        db1 = _sfloat(data_r.iloc[i+1][col_debito])
        cr0 = _sfloat(data_r.iloc[i][col_credito])
        db0 = _sfloat(data_r.iloc[i][col_debito])
        if abs(s1 - s0 - cr1 + db1) < 0.02:
            oldest_first_score += 1
        if abs(s0 - s1 - cr0 + db0) < 0.02:
            newest_first_score += 1

    if oldest_first_score >= newest_first_score:
        final_saldo = float(list(saldos)[-1])
    else:
        final_saldo = float(list(saldos)[0])

    saldo_ant = round(final_saldo - total_creditos + total_debitos, 2)
    depositos = total_creditos

    gastos_por_cat = {c: 0.0 for c in ITEMS_PAGOS}
    gastos_detalle = []

    for _, row in data.iterrows():
        deb = float(pd.to_numeric(row[col_debito], errors='coerce') or 0)
        if deb > 0:
            raw_desc = row.get(col_desc, '') if col_desc else ''
            desc     = '' if (raw_desc != raw_desc) else str(raw_desc)
            cat      = _classify_gasto(desc.lower())
            gastos_por_cat[cat] = round(gastos_por_cat.get(cat, 0) + deb, 2)
            gastos_detalle.append({'desc': desc, 'monto': deb, 'cat': cat})
            log(f"  GB gasto [{cat}]: ${deb:.2f} — {desc[:60]}")

    log(f"  GB saldo anterior: ${saldo_ant:,.2f}  depósitos: ${depositos:,.2f}")
    return {
        'saldo_anterior': saldo_ant,
        'depositos':      depositos,
        'gastos':         gastos_por_cat,
        'gastos_detalle': gastos_detalle,
    }

# ════════════════════════════════════════════════════════════════════
# GENERADOR DEL REPORTE EXCEL
# ════════════════════════════════════════════════════════════════════

def generate_report(fecha_str, facturas, recibos, bg, gb,
                    output_path=None, log=print):
    """
    Genera el Cierre Diario de Caja.

    Si output_path es None  → retorna un BytesIO listo para descargar (Streamlit).
    Si output_path es str   → guarda en disco y retorna None              (tkinter).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Cierre Diario"

    ws.column_dimensions['A'].width = 52
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 16

    def spacer(row, h=8):
        ws.row_dimensions[row].height = h

    def section_hdr(row, title):
        ws.merge_cells(f'A{row}:D{row}')
        c = ws[f'A{row}']
        c.value     = title
        c.font      = Font(name='Arial', bold=True, size=11, color=WHITE)
        c.fill      = mk_fill(DARK_BLUE)
        c.alignment = CA
        ws.row_dimensions[row].height = 22

    def col_hdr(row, headers):
        for col, hdr in zip(['A', 'B', 'C', 'D'], headers):
            c = ws[f'{col}{row}']
            c.value     = hdr
            c.font      = Font(name='Arial', bold=True, size=11)
            c.fill      = mk_fill(LIGHT_BLUE)
            c.alignment = LA if col == 'A' else CA
            c.border    = BRD
        ws.row_dimensions[row].height = 20

    def row_s1(row, label, b=None, c=None, d=None, bold=False, bg_color=None):
        ws.row_dimensions[row].height = 20
        for col in ['A', 'B', 'C', 'D']:
            cell = ws[f'{col}{row}']
            cell.border    = BRD
            cell.font      = Font(name='Arial', bold=bold, size=11)
            cell.alignment = LA if col == 'A' else RA
            if col != 'A':
                cell.number_format = CURRENCY
            if bg_color:
                cell.fill = mk_fill(bg_color)
        ws[f'A{row}'].value = label
        if b is not None: ws[f'B{row}'].value = b
        if c is not None: ws[f'C{row}'].value = c
        if d is not None: ws[f'D{row}'].value = d

    def row_s2(row, label, d=None, bold=False, bg_color=None):
        ws.row_dimensions[row].height = 20
        for col in ['A', 'B', 'C', 'D']:
            cell = ws[f'{col}{row}']
            cell.border    = BRD
            cell.font      = Font(name='Arial', bold=bold, size=11)
            cell.alignment = LA if col == 'A' else RA
            if col == 'D':
                cell.number_format = CURRENCY
            if bg_color:
                cell.fill = mk_fill(bg_color)
        ws[f'A{row}'].value = label
        if d is not None: ws[f'D{row}'].value = d

    def row_s3(row, label, b=None, c=None, bold=False, bg_color=None):
        ws.row_dimensions[row].height = 20
        for col in ['A', 'B', 'C', 'D']:
            cell = ws[f'{col}{row}']
            cell.border    = BRD
            cell.font      = Font(name='Arial', bold=bold, size=11)
            cell.alignment = LA if col == 'A' else RA
            if col in ['B', 'C']:
                cell.number_format = CURRENCY
            if bg_color:
                cell.fill = mk_fill(bg_color)
        ws[f'A{row}'].value = label
        if b is not None: ws[f'B{row}'].value = b
        if c is not None: ws[f'C{row}'].value = c

    # ── Encabezado ─────────────────────────────────────────────────
    ws.row_dimensions[1].height = 32
    ws.merge_cells('A1:D1')
    ws['A1'].value     = 'GRUPO PANABLOCK'
    ws['A1'].font      = Font(name='Arial', bold=True, size=16)
    ws['A1'].alignment = CA

    ws.row_dimensions[2].height = 24
    ws.merge_cells('A2:D2')
    ws['A2'].value     = 'Informe de Cierre Diario de Caja'
    ws['A2'].font      = Font(name='Arial', bold=True, size=13)
    ws['A2'].alignment = CA

    ws.row_dimensions[3].height = 22
    ws.merge_cells('A3:D3')
    ws['A3'].value     = f'Fecha: {fecha_str}'
    ws['A3'].font      = Font(name='Arial', size=11)
    ws['A3'].alignment = LA

    spacer(4)

    # ── Sección 1 — VENTAS Y COBROS ────────────────────────────────
    section_hdr(5, 'VENTAS Y COBROS')
    col_hdr(6, ['Descripción', 'Subtotal', 'ITBMS', 'Total'])

    row_s1(7,  'Ventas al Contado',
           b=facturas['contado_subtotal'],
           c=facturas['contado_itbms'],
           d='=B7+C7')
    row_s1(8,  'Cobros a Clientes', d=recibos)
    row_s1(9,  'Total Ventas al Contado y Cobros a Clientes',
           d='=D7+D8', bold=True, bg_color=LGRAY)
    row_s1(10, 'Total a Depositar en el Día',
           d='=D9', bold=True, bg_color=MGRAY)
    row_s1(11, 'Ventas al Crédito',
           b=facturas['credito_subtotal'],
           c=facturas['credito_itbms'],
           d='=B11+C11')

    row_s1(12, 'Total ITBMS de Ventas al Contado y Crédito',
           bold=True, bg_color=LGRAY)
    ws['C12'].value         = '=C7+C11'
    ws['C12'].number_format = CURRENCY
    ws['D12'].value         = None

    spacer(13)

    # ── Sección 2 — PAGOS Y GASTOS ─────────────────────────────────
    section_hdr(14, 'PAGOS Y GASTOS')
    col_hdr(15, ['Descripción', '', '', 'Total'])

    gastos = _merge_gastos(bg['gastos'], gb['gastos'])
    for i, item in enumerate(ITEMS_PAGOS):
        val = gastos.get(item, 0)
        row_s2(16 + i, item, d=(val if val else None))

    row_s2(24, 'Total Pagos y Gastos',
           d='=SUM(D16:D23)', bold=True, bg_color=LGRAY)

    spacer(25)

    # ── Sección 3 — MOVIMIENTOS DE BANCOS ──────────────────────────
    section_hdr(26, 'MOVIMIENTOS DE BANCOS')
    col_hdr(27, ['Descripción', 'Banco General', 'Global Bank', ''])

    row_s3(28, 'Saldo Anterior',
           b=bg['saldo_anterior'], c=gb['saldo_anterior'])
    row_s3(29, 'Depósitos según extracto bancario (efectivo, ACH, TC, Yappy)',
           b=bg['depositos'], c=gb['depositos'])
    ws.row_dimensions[29].height = 36

    bg_egresos = round(sum(bg['gastos'].values()), 2)
    gb_egresos = round(sum(gb['gastos'].values()), 2)
    row_s3(30, 'Egresos según extracto bancario',
           b=bg_egresos if bg_egresos else None,
           c=gb_egresos if gb_egresos else None)

    row_s3(31, 'Efectivo en tránsito', b='=D10')

    row_s3(32, 'Cheques en Circulación')
    ws['B32'].fill = mk_fill('FFFF00')
    ws['C32'].fill = mk_fill('FFFF00')

    row_s3(33, 'Saldo Bancario al Final del Día',
           b='=B28+B29-B30-B32', c='=C28+C29-C30-C32',
           bold=True, bg_color=LGRAY)
    ws['D32'].border = NO_BORDER

    spacer(34)

    # ── Firma ───────────────────────────────────────────────────────
    ws.row_dimensions[35].height = 30
    ws.merge_cells('A35:D35')
    ws['A35'].value     = 'Firma: ___________________________'
    ws['A35'].font      = Font(name='Arial', size=11)
    ws['A35'].alignment = LA

    spacer(36)

    # ── Leyenda ─────────────────────────────────────────────────────
    ws.row_dimensions[37].height = 16
    ws.merge_cells('A37:D37')
    ws['A37'].value = ('Leyenda:  Celda en amarillo = ingresar manualmente  '
                       '(solo Cheques en Circulación)')
    ws['A37'].font      = Font(name='Arial', italic=True, size=9, color='595959')
    ws['A37'].alignment = LA
    ws['A37'].fill      = mk_fill('FFFF00')

    # ── Notas de Crédito ────────────────────────────────────────────
    ncs = facturas.get('notas_credito', [])
    if ncs:
        spacer(38)
        ws.row_dimensions[39].height = 22
        ws.merge_cells('A39:F39')
        ws['A39'].value     = 'NOTAS DE CRÉDITO DEL DÍA'
        ws['A39'].font      = Font(name='Arial', bold=True, size=11, color=WHITE)
        ws['A39'].fill      = mk_fill(RED_FILL)
        ws['A39'].alignment = CA

        nc_hdrs = ['N° NC', 'Fact. Anulada', 'Cliente', 'Motivo',
                   'Subtotal', 'ITBMS', 'Total']
        for j, hdr in enumerate(nc_hdrs, 1):
            c = ws.cell(40, j)
            c.value     = hdr
            c.font      = Font(name='Arial', bold=True, size=10)
            c.fill      = mk_fill(RED_FILL)
            c.alignment = CA
            c.border    = BRD
            if c.column_letter not in ['A', 'B', 'C', 'D']:
                ws.column_dimensions[c.column_letter].width = 14
        ws.column_dimensions['C'].width = 28
        ws.column_dimensions['D'].width = 22

        for k, nc in enumerate(ncs):
            r = 41 + k
            ws.row_dimensions[r].height = 18
            for j, val in enumerate([
                nc['nc'], nc['factura'], nc['cliente'], nc['motivo'],
                nc['subtotal'], nc['itbms'], nc['total']
            ], 1):
                c = ws.cell(r, j)
                c.value     = val
                c.font      = Font(name='Arial', size=10)
                c.fill      = mk_fill('FFE0E0')
                c.border    = BRD
                c.alignment = LA if j <= 4 else RA
                if j >= 5:
                    c.number_format = CURRENCY

    ws.freeze_panes = 'A6'

    # ── Guardar ─────────────────────────────────────────────────────
    if output_path is None:
        # Modo Streamlit: retornar BytesIO para descarga directa
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf
    else:
        # Modo escritorio: guardar en disco
        wb.save(output_path)
        log(f"\n✔ Reporte guardado en: {output_path}")
        return None
