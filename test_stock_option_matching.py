import unittest
from import_stock_update import match_running_singlet


class SingletMatchingTests(unittest.TestCase):
    def setUp(self):
        self.rows = [dict(barcode='LGR-L', outbound_name='ICE LITE 초냉감 러닝 나시_LIGHT GREY_L', stock_qty=44),
                     dict(barcode='LGR-2XL', outbound_name='ICE LITE 초냉감 러닝 나시_LIGHT GREY_2XL', stock_qty=4),
                     dict(barcode='BLK-L', outbound_name='ICE LITE 초냉감 러닝 나시_BLACK_L', stock_qty=89)]

    def test_color_size_and_name_aliases(self):
        for suffix in ('', ' (싱글렛)'):
            for color in ('라이트그레이', 'LIGHT GREY', 'LIGHT GRAY'):
                flag, row = match_running_singlet(dict(name='ICE LITE 초냉감 러닝 나시'+suffix, color=color, size='XXL'), self.rows)
                self.assertTrue(flag)
                self.assertEqual(row['stock_qty'], 4)

    def test_missing_variant_does_not_borrow_other_size(self):
        flag, row = match_running_singlet(dict(name='ICE LITE 초냉감 러닝 나시', color='라이트그레이', size='M'), self.rows)
        self.assertTrue(flag)
        self.assertIsNone(row)

    def test_ambiguous_source_is_not_used(self):
        item = dict(name='ICE LITE 초냉감 러닝 나시', color='라이트그레이', size='L')
        self.assertIsNone(match_running_singlet(item, self.rows + [self.rows[0]])[1])


if __name__ == '__main__':
    unittest.main()
