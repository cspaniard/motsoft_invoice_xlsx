# -*- coding: utf-8 -*-
import unicodedata
from collections import OrderedDict, defaultdict

from odoo import models
from odoo.exceptions import UserError
from xlsxwriter.utility import xl_rowcol_to_cell

INVOICE_TYPES = ('out_invoice', 'in_invoice')
REFUND_TYPES = ('out_refund', 'in_refund')

# True  -> las rectificativas se muestran en negativo (efecto contable).
# False -> se muestran en positivo, tal como aparecen en el documento.
REFUNDS_NEGATIVE = True

# Formato base de TODAS las celdas (usadas o no)
BASE_FORMAT = {
    'font_name': 'Arial',
    'font_size': 11,
    'valign': 'vcenter',
}
NUM_FORMAT = '#,##0.00'
DATE_FORMAT = 'dd/mm/yyyy'
MAX_COL = 16383  # última columna de Excel (XFD)

# Fila (base 0) de la cabecera de la tabla.
# Filas 0 y 1: Empresa / Fechas. Fila 2: separador en blanco.
HEADER_ROW = 3


def _strip_accents(text):
    """Quita acentos pero conserva ñ/Ñ: 'Ángel Muñoz' -> 'Angel Muñoz'."""
    text = (text or '').replace('ñ', '\x00a').replace('Ñ', '\x00b')
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    return text.replace('\x00a', 'ñ').replace('\x00b', 'Ñ')


class MotsoftInvoiceXlsx(models.AbstractModel):
    _name = 'report.motsoft_invoice_xlsx.report_invoice_xlsx'
    _inherit = 'report.report_xlsx.abstract'
    _description = 'Facturas agrupadas por partner (XLSX)'

    # ------------------------------------------------------------------
    # Opciones del workbook: Arial 11 + centrado vertical por defecto
    # ------------------------------------------------------------------
    def get_workbook_options(self):
        options = super().get_workbook_options() or {}
        options.update({
            'in_memory': True,
            'default_format_properties': dict(BASE_FORMAT),
        })
        return options

    # ------------------------------------------------------------------
    # Preparación de datos
    # ------------------------------------------------------------------
    def _get_sign(self, move):
        """Signo para convertir el balance contable en importe de documento."""
        sign = -1 if move.is_sale_document() else 1
        if not REFUNDS_NEGATIVE and move.move_type in REFUND_TYPES:
            sign = -sign
        return sign

    def _prepare_move(self, move):
        currency = move.company_id.currency_id
        sign = self._get_sign(move)
        product_lines = move.invoice_line_ids.filtered(lambda l: not l.display_type)

        base = sign * sum(product_lines.mapped('balance'))

        taxes = defaultdict(float)
        for line in move.line_ids.filtered('tax_line_id'):
            taxes[line.tax_line_id.id] += sign * line.balance

        # Impuestos aplicados a las líneas (aunque su importe sea 0)
        applied_tax_ids = set(taxes)
        for tax in product_lines.mapped('tax_ids').flatten_taxes_hierarchy():
            applied_tax_ids.add(tax.id)

        pay_lines = move.line_ids.filtered(
            lambda l: l.account_id.user_type_id.type in ('receivable', 'payable'))
        if pay_lines:
            total = -sign * sum(pay_lines.mapped('balance'))
        else:
            total = base + sum(taxes.values())

        return {
            'id': move.id,
            'name': move.name or '/',
            'date': move.invoice_date or move.date,
            'is_refund': move.move_type in REFUND_TYPES,
            'base': currency.round(base),
            'taxes': {tid: currency.round(amt) for tid, amt in taxes.items()},
            'tax_ids': applied_tax_ids,
            'total': currency.round(total),
        }

    def _sorted_taxes(self, tax_ids):
        taxes = self.env['account.tax'].browse(list(tax_ids))
        # Primero impuestos positivos (de mayor a menor), después retenciones
        return taxes.sorted(
            key=lambda t: (t.amount < 0, -abs(t.amount), t.name or '', t.id))

    # ------------------------------------------------------------------
    # Generación del XLSX
    # ------------------------------------------------------------------
    def generate_xlsx_report(self, workbook, data, objs):
        moves = objs.filtered(
            lambda m: m.move_type in INVOICE_TYPES + REFUND_TYPES and m.state != 'cancel')
        if not moves:
            raise UserError(
                'No hay facturas válidas en la selección (se ignoran asientos '
                'y facturas canceladas).')

        # --- datos -------------------------------------------------------
        prepared = {m.id: self._prepare_move(m) for m in moves}
        all_tax_ids = set()
        for p in prepared.values():
            all_tax_ids |= p['tax_ids']
        taxes = self._sorted_taxes(all_tax_ids)
        tax_order = taxes.ids

        for p in prepared.values():
            p['vals'] = ([p['base']]
                         + [p['taxes'].get(tid, 0.0) for tid in tax_order]
                         + [p['total']])

        # --- información de cabecera (empresa y rango de fechas) -----------
        companies = moves.mapped('company_id') or self.env.company
        company_name = ', '.join(companies.mapped('name'))

        dates = [p['date'] for p in prepared.values() if p['date']]
        if dates:
            date_min, date_max = min(dates), max(dates)
            date_range = '%s - %s' % (
                date_min.strftime('%d/%m/%Y'), date_max.strftime('%d/%m/%Y'))
        else:
            date_range = ''

        groups = OrderedDict()
        for m in moves:
            partner = m.commercial_partner_id

            g = groups.setdefault(partner.id, {
                'vat': partner.vat,
                'name': _strip_accents(partner.display_name) or 'Sin Contacto',
                'invoices': [],
                'refunds': [],
            })

            p = prepared[m.id]
            g['refunds' if p['is_refund'] else 'invoices'].append(p)

        ordered_groups = sorted(
            groups.values(),
            key=lambda g: (_strip_accents(g['name']).casefold(), g['name']))

        for g in ordered_groups:
            g['invoices'].sort(key=lambda p: (p['name'], p['id']))
            g['refunds'].sort(key=lambda p: (p['name'], p['id']))

        # --- hoja y formatos ----------------------------------------------
        ws = workbook.add_worksheet('Facturas')

        def fmt(**kw):
            props = dict(BASE_FORMAT)
            props.update(kw)
            return workbook.add_format(props)

        f_text = fmt()
        f_date = fmt(num_format=DATE_FORMAT, align='left')
        f_num = fmt(num_format=NUM_FORMAT)
        f_head = fmt(bold=True, align='center', text_wrap=True,
                     bg_color='#D9D9D9', bottom=1)
        f_partner = fmt(bold=True)
        f_section = fmt(italic=True, bold=True)
        f_tot_text = fmt(bold=True, top=1)
        f_tot_num = fmt(bold=True, top=1, num_format=NUM_FORMAT)
        f_info_label = fmt(bold=True)
        f_info_value = fmt(align='left')

        n_vals = len(tax_order) + 2          # base + impuestos + total
        first_val_col = 2
        last_col = first_val_col + n_vals - 1

        # Anchos + formato de columna (para que las celdas vacías también
        # sean Arial 11 / centradas verticalmente)
        ws.set_column(0, 0, 13, f_text)
        ws.set_column(1, 1, 24, f_text)
        ws.set_column(first_val_col, last_col, 18, f_text)
        ws.set_column(last_col + 1, MAX_COL, None, f_text)
        ws.set_landscape()
        ws.set_paper(9)  # A4
        ws.fit_to_pages(1, 0)  # 1 página de ancho, alto automático
        ws.set_margins(left=0.5, right=0.5, top=0.6, bottom=0.6)  # márgenes en pulgadas (opcional)
        # Inmoviliza info (empresa/fechas) + cabecera de la tabla
        ws.freeze_panes(HEADER_ROW + 1, 0)
        ws.repeat_rows(HEADER_ROW)           # repite la cabecera al imprimir

        # --- información superior ----------------------------------------
        ws.write_string(0, 0, 'Empresa:', f_info_label)
        ws.write_string(0, 1, company_name, f_info_value)
        ws.write_string(1, 0, 'Fechas:', f_info_label)
        ws.write_string(1, 1, date_range, f_info_value)

        # --- cabecera de la tabla ------------------------------------------
        ws.set_row(HEADER_ROW, 34)
        ws.write_string(HEADER_ROW, 0, 'Fecha', f_head)
        ws.write_string(HEADER_ROW, 1, 'Número de Factura', f_head)
        ws.write_string(HEADER_ROW, 2, 'Base Imponible', f_head)
        for i, tax in enumerate(taxes):
            ws.write_string(HEADER_ROW, 3 + i, tax.name or '', f_head)
        ws.write_string(HEADER_ROW, last_col, 'Total Factura', f_head)

        # --- helpers de escritura ------------------------------------------
        def write_move(r, p):
            ws.write_datetime(r, 0, p['date'], f_date)
            ws.write_string(r, 1, p['name'], f_text)
            for i, v in enumerate(p['vals']):
                ws.write_number(r, first_val_col + i, v, f_num)

        def write_total_row(r, label, formula_for_col, vals_list):
            """formula_for_col(col) -> fórmula; vals_list -> lista de listas."""
            ws.write_string(r, 0, label, f_tot_text)
            ws.write_blank(r, 1, None, f_tot_text)
            for i in range(n_vals):
                value = round(sum(v[i] for v in vals_list), 2)
                formula = formula_for_col(first_val_col + i)
                if len(formula) > 8000:      # límite de Excel: 8192 caracteres
                    ws.write_number(r, first_val_col + i, value, f_tot_num)
                else:
                    ws.write_formula(r, first_val_col + i, formula, f_tot_num, value)

        def sum_range(first_row, last_row):
            return lambda col: '=SUM(%s:%s)' % (
                xl_rowcol_to_cell(first_row, col), xl_rowcol_to_cell(last_row, col))

        def sum_rows(rows):
            return lambda col: '=' + '+'.join(xl_rowcol_to_cell(r, col) for r in rows)

        # --- cuerpo --------------------------------------------------------
        r = HEADER_ROW + 1
        inv_total_rows, ref_total_rows = [], []
        all_inv_vals, all_ref_vals = [], []

        for g in ordered_groups:
            invoices, refunds = g['invoices'], g['refunds']

            if g['vat']:
                ws.write_string(r, 0, g['vat'], f_partner)
            else:
                ws.write_string(r, 0, 'Sin CIF', f_partner)

            r += 1
            ws.write_string(r, 0, g['name'], f_partner)
            r += 1

            inv_row = ref_row = None
            if invoices:
                first = r
                for p in invoices:
                    write_move(r, p)
                    r += 1
                label = 'Total facturas' if refunds else 'Total %s' % g['name']
                inv_vals = [p['vals'] for p in invoices]
                write_total_row(r, label, sum_range(first, r - 1), inv_vals)
                inv_row = r
                inv_total_rows.append(r)
                all_inv_vals += inv_vals
                r += 1

            if refunds:
                r += 1
                ws.write_string(r, 0, f'Rectificativas {g["name"] or ""}', f_section)
                r += 1
                first = r
                for p in refunds:
                    write_move(r, p)
                    r += 1
                label = ('Total rectificativas' if invoices
                         else 'Total %s (rectificativas)' % g['name'])
                ref_vals = [p['vals'] for p in refunds]
                write_total_row(r, label, sum_range(first, r - 1), ref_vals)
                ref_row = r
                ref_total_rows.append(r)
                all_ref_vals += ref_vals
                r += 1

            if invoices and refunds:
                write_total_row(r, 'Total %s' % g['name'],
                                sum_rows([inv_row, ref_row]),
                                [p['vals'] for p in invoices + refunds])
                r += 1

            r += 1  # línea en blanco entre partners

        # --- totales generales ------------------------------------------------
        has_inv, has_ref = bool(inv_total_rows), bool(ref_total_rows)
        g_inv_row = g_ref_row = None
        if has_inv:
            write_total_row(r, 'TOTAL FACTURAS' if has_ref else 'TOTAL GENERAL',
                            sum_rows(inv_total_rows), all_inv_vals)
            g_inv_row = r
            r += 1
        if has_ref:
            write_total_row(r, 'TOTAL RECTIFICATIVAS' if has_inv else 'TOTAL GENERAL',
                            sum_rows(ref_total_rows), all_ref_vals)
            g_ref_row = r
            r += 1
        if has_inv and has_ref:
            write_total_row(r, 'TOTAL GENERAL', sum_rows([g_inv_row, g_ref_row]),
                            all_inv_vals + all_ref_vals)
