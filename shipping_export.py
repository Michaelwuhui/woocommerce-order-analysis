"""Spreadsheet builders used by shipping export endpoints."""

import os
from copy import copy
from io import BytesIO

from openpyxl import load_workbook


def _safe_excel_text(value):
    """Normalize customer data before storing it as an explicit string cell."""
    return str(value or '').replace('\r', ' ').replace('\n', ' ').strip()


def build_australia_shipping_workbook(rows, template_path=None):
    """Build an xlsx matching the logistics partner's 地址拆分 workbook."""
    if template_path is None:
        template_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            'assets', 'export_templates', 'australia_shipping_template.xlsx',
        )

    workbook = load_workbook(template_path)
    worksheet = workbook['地址拆分']

    # Row 2 is a sanitized formatting sample retained from the partner's file.
    style_row = []
    for column in range(1, 9):
        source = worksheet.cell(row=2, column=column)
        style_row.append({
            'style': copy(source._style),
            'number_format': source.number_format,
            'alignment': copy(source.alignment),
        })
    if worksheet.max_row > 1:
        worksheet.delete_rows(2, worksheet.max_row - 1)
    if worksheet.max_column > 8:
        worksheet.delete_cols(9, worksheet.max_column - 8)

    for index, item in enumerate(rows, start=1):
        values = [
            index,
            _safe_excel_text(item.get('order_number')),
            _safe_excel_text(item.get('customer_name')),
            _safe_excel_text(item.get('phone')),
            _safe_excel_text(item.get('city')),
            _safe_excel_text(item.get('state')),
            _safe_excel_text(item.get('address')),
            _safe_excel_text(item.get('tracking_number')),
        ]
        row_number = index + 1
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row=row_number, column=column, value=value)
            sample = style_row[column - 1]
            cell._style = copy(sample['style'])
            cell.alignment = copy(sample['alignment'])
            if column in (2, 3, 4, 5, 6, 7, 8):
                # Explicit string type protects leading zeroes / long tracking
                # numbers and prevents formula injection without displaying an
                # apostrophe in the logistics partner's sheet.
                cell.data_type = 's'
                cell.number_format = '@'
            else:
                cell.number_format = sample['number_format']

    worksheet.auto_filter.ref = f'A1:H{len(rows) + 1}'
    worksheet.freeze_panes = 'A2'
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.orientation = 'landscape'
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.print_title_rows = '1:1'
    worksheet.print_area = f'A1:H{len(rows) + 1}'

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output
