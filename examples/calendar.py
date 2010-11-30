from xlrep import Report
from xlwt import Workbook, easyxf
from datetime import datetime, timedelta
from random import randrange

report_name = 'calendar.xls'

cell_style = easyxf('''
            font: height 200, name Arial;
            align: horiz center;
            border: right thin, left thin, top thin, bottom thin;
            ''')

cell_style_weekend = easyxf('''
            font: height 200, name Arial;
            pattern: pattern solid, fore-colour rose;
            border: left thin, right thin, top thin, bottom thin;
            align: horiz center;
            ''')

total_style = easyxf('''
            font: bold on, height 200, name Arial;
            pattern: pattern solid, fore-colour light_yellow;
            border: top thin, bottom thin;
            ''')

row_header_style = easyxf('''
            font: bold on, height 200, name Arial;
            pattern: pattern solid, fore-colour gray25;
            border: left thin, right thin, top thin, bottom thin;
            ''')


# Following functions generate list of dates

def dategen(from_date=datetime.now(), to_date=None):
    while to_date is None or from_date <= to_date:
        yield from_date
        from_date = from_date + timedelta(days=1)

def weekgen(date_start, date_end):
    week = []
    for day in dategen(date_start, date_end):
        week.append(day)
        if day.weekday() == 6 or day == date_end:
            yield week
            week = []

def weeknum(date):
    n = int(date.strftime('%W'))
    if n == 0:
        return int(datetime(date.year - 1, 12, 31).strftime('%W'))
    return n


# Construct report 
def draw():
    report = Report("A calendar-like report")     # Create report object
    rs = report.rows
    cs = report.cols

    report.cols_width = 1600

    # Create row headers (calendar)
    for week in weekgen(datetime.now() - timedelta(days=30), datetime.now()):
        wes = rs.add_section()
        for day in week:
            if day.weekday() in (5,6):
                style = cell_style_weekend
            else:
                style = cell_style
            wes.add_field(day.strftime("%b %d"), style=style)
        wes.add_calc('Week %s' % weeknum(week[0]), func=sum, style=row_header_style)
    rs.add_calc('Total', func=sum, style=total_style)


    # Create column headers (just one calc field that computes average)
    cs.add_calc('Average', func=lambda x: 1.0*sum(x)/len(x), style=total_style)

    # Generate random data
    data = [[randrange(100) for i in range(5)] for j in range(30)]

    book = Workbook()
    ws = book.add_sheet('Worksheet')
    report.render(ws, data)             # Render report
    book.save(report_name)              # Save report

draw()
print 'Report has been constructed:', report_name
