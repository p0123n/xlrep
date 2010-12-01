from xlrep import Report
from xlwt import Workbook, easyxf
from random import randrange

calc_style = easyxf('pattern: pattern solid, fore-colour gray25;')
header_style = easyxf('pattern: pattern solid, fore-colour rose;')
total_style = easyxf('font: bold on;')

report_name = 'quadric.xls'
rep = Report("Quadric report")      # Create report object

cs = rep.cols                       # Create alias for column and row sections
rs = rep.rows

for i in range(3):                  # Create row sections
    ccs = cs.add_section('col %d' % i)
    for j in range(3):
        ccs.add_field('%d' % j, header_style=header_style)
    ccs.add_calc('', func=sum, style=calc_style)
cs.add_calc('grand total', func=sum, style=total_style)

for i in range(3):                  # Create column sections
    rrs = rs.add_section('row %d' % i)
    for j in range(3):
        rrs.add_field('%d' % j, header_style=header_style)
    rrs.add_calc('', func=sum, style=calc_style)
rs.add_calc('grand total', func=sum, style=total_style)

rep.cols_width = 1200

data = [[randrange(10) for j in range(9)] for i in range(9)]  # Generate data

book = Workbook()                   # xlwt part: create Workbook and Worksheet
ws = book.add_sheet('Worksheet')
rep.render(ws, data)                # Render report
book.save(report_name)              # Save report

print 'Report has been constructed:', report_name
