import os
import sys
import unittest

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shipping_export import build_australia_shipping_workbook


class AustraliaShippingWorkbookTests(unittest.TestCase):
    def test_template_contains_no_customer_rows(self):
        template = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'assets', 'export_templates', 'australia_shipping_template.xlsx',
        )
        workbook = load_workbook(template, data_only=False)
        worksheet = workbook['地址拆分']
        self.assertEqual(
            ['序号', '订单信息', '客户姓名', '电话', '城市', '州', '收货地址', '正式运输单号'],
            [worksheet.cell(1, column).value for column in range(1, 9)],
        )
        self.assertIn(worksheet['I1'].value, (None, ''))
        for row in worksheet.iter_rows(min_row=2, max_col=8):
            self.assertTrue(all(cell.value in (None, '') for cell in row))

    def test_values_are_text_safe_and_keep_leading_zeroes(self):
        rows = [{
            'order_number': '#TEST-001',
            'customer_name': 'Sample Customer',
            'phone': '0412345678',
            'city': 'Perth',
            'state': 'WA',
            'address': '=HYPERLINK("https://invalid.test")',
            'tracking_number': '00000000000000000002',
        }]
        output = build_australia_shipping_workbook(rows)
        workbook = load_workbook(output, data_only=False)
        worksheet = workbook['地址拆分']

        self.assertEqual('0412345678', worksheet['D2'].value)
        self.assertEqual('00000000000000000002', worksheet['H2'].value)
        self.assertEqual('=HYPERLINK("https://invalid.test")', worksheet['G2'].value)
        self.assertEqual(8, worksheet.max_column)
        for address in ('B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2'):
            self.assertEqual('s', worksheet[address].data_type)
            self.assertEqual('@', worksheet[address].number_format)

        self.assertEqual('A1:H2', worksheet.auto_filter.ref)
        self.assertEqual('A2', worksheet.freeze_panes)


if __name__ == '__main__':
    unittest.main()
