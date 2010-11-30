from xlrep import Report
from xlwt import Workbook
from os import system
from StringIO import StringIO

import unittest
import xlrd

test_file = 'test.xls'

class TestReport(unittest.TestCase):

    def test_report_general(self):
        # Prepare
        r = Report()
        r.caption = 'test_report_0'

        hs = r.cols
        hs.add_field('col 0')
        hs.add_field('col 1')
        hss = hs.add_section('col section 0')
        hss.add_field('col 2_1')
        hss.add_calc('col total', sum)
        hs.add_field('col 3')
        hs.add_calc('col total 0', func=lambda x: 1.0*sum(x)/len(x))

        rs = r.rows
        rs.add_field('row 0')
        rss = rs.add_section('row section 0')
        rss.add_field('row 1_1')
        rss.add_field('row 1_2')
        rss.add_calc('row total 0', func=sum)
        rs.add_field('row 2')
        rs.add_field('row 3')
        rs.add_calc('row total 1', func=sum)

        book = Workbook()
        ws = book.add_sheet('test worksheet')
        data = (range(4) for i in range(5))
        r.render(ws, data)
        compiled_report = StringIO()
        book.save(compiled_report)

        # Test

        book = xlrd.open_workbook(file_contents=compiled_report.getvalue())
        ws = book.sheet_by_index(0)

        test_data = [
            [0.0, 1.0, 2.0, 2.0, 3.0, 1.5],
            [0.0, 1.0, 2.0, 2.0, 3.0, 1.5],
            [0.0, 1.0, 2.0, 2.0, 3.0, 1.5],
            [0.0, 2.0, 4.0, 4.0, 6.0, 3.0],
            [0.0, 1.0, 2.0, 2.0, 3.0, 1.5],
            [0.0, 1.0, 2.0, 2.0, 3.0, 1.5],
            [0.0, 5.0, 10.0, 10.0, 15.0, 7.5]
        ]

        for i in range(3,10):
            for j in range(2,8):
                self.assertEquals(ws.cell(i,j).value, test_data[i-3][j-2])

    def test_report_calc_0(self):
        """Test fields_ignore feature"""

        # Prepare
        r = Report()
        r.caption = 'test_report_calc_0'

        hs = r.cols
        hs.add_field('col 0')
        hs.add_field('col 1')
        col2 = hs.add_field('col ignore')
        hs.add_calc('col total', sum, fields_ignore=(col2,))

        rs = r.rows
        rs.add_field('row 0')
        rs.add_field('row 1')
        row2 = rs.add_field('row ignore')
        rs.add_calc('row total', sum, fields_ignore=(row2,))

        data = ((1,1,1),(2,2,2),(3,3,3))
        book = Workbook()
        ws = book.add_sheet('test worksheet')

        r.render(ws, data)
        compiled_report = StringIO()
        book.save(compiled_report)

        # Test

        book = xlrd.open_workbook(file_contents=compiled_report.getvalue())
        ws = book.sheet_by_index(0)

        test_data = [
            [1.0, 1.0, 1.0, 2.0],
            [2.0, 2.0, 2.0, 4.0],
            [3.0, 3.0, 3.0, 6.0],
            [3.0, 3.0, 3.0, 6.0],
        ]

        for i in range(4):
            for j in range(4):
                self.assertEquals(ws.cell(i+2,j+1).value, test_data[i][j])

    def test_report_calc_1(self):
        """Test cross_fields_ignore feature"""

        # Prepare
        r = Report()
        r.caption = 'test_report_calc_1'

        hs = r.cols
        hs.add_field('col 0')
        hs.add_field('col 1')
        col2 = hs.add_field('col ignore')
        hs.add_calc('col total', sum)

        rs = r.rows
        rs.add_field('row 0')
        rs.add_field('row 1')
        rs.add_field('row 2')
        rs.add_calc('row total', sum, cross_fields_ignore=(col2,))

        data = ((1,1,1),(2,2,2),(3,3,3))
        book = Workbook()
        ws = book.add_sheet('test worksheet')

        r.render(ws, data)
        compiled_report = StringIO()
        book.save(compiled_report)
        book.save(test_file)

        # Test

        book = xlrd.open_workbook(file_contents=compiled_report.getvalue())
        ws = book.sheet_by_index(0)

        test_data = [
            [1.0, 1.0, 1.0, 3.0],
            [2.0, 2.0, 2.0, 6.0],
            [3.0, 3.0, 3.0, 9.0],
            [6.0, 6.0, '', 12.0],
        ]

        for i in range(4):
            for j in range(4):
                self.assertEquals(ws.cell(i+2,j+1).value, test_data[i][j])


    def test_report_calc_2(self):
        """Test fields feature"""

        # Prepare
        r = Report()

        hs = r.cols
        c0 = hs.add_field('col 0')
        c1 = hs.add_field('col 1')
        hs.add_field('col ignore')
        hs.add_calc('col total', sum, fields=(c0,c1))

        rs = r.rows
        r0 = rs.add_field('row 0')
        r1 = rs.add_field('row 1')
        rs.add_field('row ignore')
        rs.add_calc('row total', sum, fields=(r0,r1))

        data = ((1,1,1),(2,2,2),(3,3,3))
        book = Workbook()
        ws = book.add_sheet('test worksheet')

        r.render(ws, data)
        compiled_report = StringIO()
        book.save(compiled_report)
        #book.save(test_file)

        # Test

        book = xlrd.open_workbook(file_contents=compiled_report.getvalue())
        ws = book.sheet_by_index(0)

        test_data = [
            [1.0, 1.0, 1.0, 2.0],
            [2.0, 2.0, 2.0, 4.0],
            [3.0, 3.0, 3.0, 6.0],
            [3.0, 3.0, 3.0, 6.0],
        ]

        #system("oowriter %s" % test_file)

        for i in range(4):
            for j in range(4):
                self.assertEquals(ws.cell(i+1,j+1).value, test_data[i][j])


def suite():
    suite = unittest.TestSuite()
    suite.addTest(unittest.makeSuite(TestReport))
    return suite

if __name__ == '__main__':
    unittest.main()
