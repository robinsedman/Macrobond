import unittest
from macrobond import c_macrobond


class MacrobondTest(unittest.TestCase):
    def test_bbg_ticker(self):
        """
        Test static function: f_create_bbg_ticker
        """
        result = c_macrobond.Macrobond.f_create_bbg_ticker(bbg_ticker=['SPX Index'])
        self.assertEqual(result, ['ih:bl:spx index'])

        result = c_macrobond.Macrobond.f_create_bbg_ticker(bbg_ticker=['SPX Index'], **{'BBG_Fields': ['PX_LAST']})
        self.assertEqual(result, ['ih:bl:spx index:px_last'])

    def test_regions(self):
        """
        Check type of output
        """
        result = c_macrobond.Macrobond.f_region_map()
        result_0 = result[0]
        result_1 = result[1]
        self.assertTrue(type(result) is tuple)
        self.assertTrue(type(result_0) is dict)
        self.assertTrue(type(result_1) is dict)
        self.assertEqual(len(result_0), len(result_1))


if __name__ == '__main__':
    unittest.main()
