from xlrep import Report
from xlwt import Workbook

report_name = 'simple.xls'

rep = Report("A simple report")     # Create report object

cs = rep.cols                       # Add columns
cs.add_field('col 0')
cs.add_field('col 1')
css = cs.add_section('section 0')   # Add a subsection
css.add_field('col 3')              # Add fields to the subsection
css.add_field('col 4')

rs = rep.rows                       # Add rows
rs.add_calc('total', func=sum)      # Add just a calc field

data = (range(4) for i in range(10))    # Generate some data
                                        # Note: dimentions of data matrix should
                                        # be compatible with number of data
                                        # fields 

book = Workbook()                   # xlwt part: create Workbook and Worksheet
ws = book.add_sheet('Worksheet')
rep.render(ws, data)                # Render report
book.save(report_name)              # Save report

print 'Report has been constructed:', report_name
