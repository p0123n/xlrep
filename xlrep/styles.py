# -*- coding: utf-8 -*-
from xlwt import easyxf

caption_style = easyxf('''
            font: bold on, height 280, name Arial;
            ''')

description_style = easyxf('''
            font: height 200, name Arial;
            ''')

col_header_style = easyxf('''
            font: bold on, height 200, name Arial; 
            align: wrap on, horiz center, vert center;
            border: left thin, right thin, top thin, bottom thin;
            pattern: pattern solid, fore-colour tan;
            ''')

row_header_style = easyxf('''
            font: bold on, height 200, name Arial;
            border: left thin, right thin, top thin, bottom thin;
            align: wrap on;
            pattern: pattern solid, fore-colour tan;
            ''')

cell_style = easyxf('''
            font: height 200, name Arial;
            border: left thin, right thin, top thin, bottom thin;
            align: wrap on;
            ''')
