# -*- coding: utf-8 -*-
from xlwt import easyxf, XFStyle, Font
from StringIO import StringIO
from styles import caption_style, description_style, col_header_style, row_header_style, cell_style

"""
Konstantin Selivanov, 2010

MS Excel and OO Cacl reports generator.
Can be used as a high-level framework upon low-level xlwt.
For usage examples see ./examples

"""

class ReportException(Exception):
    """Raise on internal errors"""
    pass

class _truth_container(object):
    """Special collection that always returns True
    on the __contains__ check""" 
    def __contains__(self, item):
        return True

class Field(object):
    """Abstract field class"""
    def __init__(self, name, style=None, header_style=None, width=None, height=None, num_format=None):
        """Base field calls. Should not be created directly.

        Keyword arguments:
        name    --      field name
        style   --      XFStyle that applies to header cells and regular cells
        header_style -- XFStyle that only applies to regular cells
        width   --      field width
        height  --      field height
        num_format  --  excel-like data format

        """
        self.name = name
        self.style = style
        self.header_style = header_style
        self.width = width
        self.height = height
        self.num_format = num_format
        self._level = 0

class DataField(Field):
    """Field that contains data value"""
    def __init__(self, name, style=None, header_style=None, width=None, height=None, num_format=None):
        """Data field class. Should not be created directly but only through Section.add_field method.

        """
        Field.__init__(self, name, style, header_style, width, height, num_format)

class CalcField(Field):
    def __init__(self, section, name, func, fields, fields_ignore, cross_fields, cross_fields_ignore, style=None, header_style=None, width=None, height=None, num_format=None):
        """
        Field that contains calculated values. Should not be created directly
        but only through Section.add_calc method.

        """
        Field.__init__(self, name, style, header_style, width, height, num_format)
        self.sec = section
        self.func = func
        self._fields = fields
        self._fields_ignore = fields_ignore
        self._cross_fields = cross_fields
        self._cross_fields_ignore = cross_fields_ignore

    @property
    def fields(self):
        """If self._fields is empty return all data fields in section"""
        if not self._fields:
            #return self.sec.get_data_fields()
            return _truth_container()
        return self._fields

    @property
    def fields_ignore(self):
        """Just return fields_ignore"""
        return self._fields_ignore

    @property
    def cross_fields(self):
        """If self._cross_fields is empty return all cross fields.
        For a column section cross-fields is rows.
        Far a row section cross-fields is columns."""
        if not self._cross_fields:
            return _truth_container()
        return self._cross_fields

    @property
    def cross_fields_ignore(self):
        """Just return fields_ignore"""
        return self._cross_fields_ignore

class Section(object):
    def __init__(self, name='', style=None, header_style=None, collapse=False):
        """Section is a logical fields container. It should be used
        - to group fields
        - when calc fields are used 

        Keyword arguments:
        name --             section name
        style --            XFStyle that is applied to all cells in section
        header_style --     XFStyle that is only applied to header cells
        collapse --         if True creates section with ability to collapse

        """
        self.name = name
        self.style = style
        self.header_style = header_style
        self.collapse = collapse
        self._items = []
        self.visible = False
        self._level = 0
        if self.name:
            self.visible = True         

    def add_field(self, name='', style=None, header_style=None, width=None, height=None, num_format=None):
        """Creates data field in the section.
        
        Keyword arguments:
        name --             field name (default: empty string)
        style --            XFStyle that is applied to all cells in section
        header_style --     XFStyle that is only applied to header cells
        width --            field width
        height --           field height
        num_format --       excel-like cell format

        """
        style = style or self.style
        header_style = header_style or self.header_style
        field = DataField(name, style, header_style, width, height, num_format)
        field._level = self._level
        self._items.append(field)
        return field

    def _add_fake_field(self):
        """Create fake field"""
        field = DataField(name='', style=self.style, header_style=self.header_style)
        self._items.insert(0, field)
        return field

    def add_section(self, name='', style=None, header_style=None, collapse=True):
        """Add subsection to the section.
        Section is a logical fields container. It should be used
        - to group fields
        - when calc fields are used 

        When section name is specified it creates a folded section.
        If not the section is invisible.

        In the Report class there are two base sections
            report.cols - base column section 
            report.rows - base row section
        
        Keyword arguments:
        name --             section name. If is set to empty string 
        style --            XFStyle that is applied to all cells in section
        header_style --     XFStyle that is only applied to header cells
        collapse --         if True creates section with ability to collapse

        """
        style = style or self.style
        header_style = header_style or self.header_style
        section = Section(name, style, header_style, collapse)
        section._level = self._level + (1 if collapse else 0)
        self._items.append(section)
        return section
    
    def add_calc(self, name, func, fields=[], fields_ignore=[], cross_fields=[], cross_fields_ignore=[], style=None, header_style=None, width=None, height=None, num_format=None):
        """Add calculated field. Calculated field (or calc field) is 
        special field with attached function.

        The function could be any function that gets sequence as input
        and calculates single value. The input sequence is values
        of the fields in current section.

        Example: one could create total sum as follow
            section.add_calc('Total', sum)

        Keyword arguments:
        name --             field name
        func --             aggregation function
        fields --           computes function only for specified fields
                            if empty (by default) computes function for
                            all data(!) fields in the section 
        fields_ignore --    computes function for all data fields in the
                            section except these fields
        cross_fields --     computes function only for specified fields
                            which has opposite layout 
                            e.g. if current section is column section 
                            than cross-fields are rows
                            and if current section is row section then
                            cross-fields are columns
                            if empty (by default) computes function for
                            all fields with opposite layout
        cross_fields_ignore --  computes function for all fields with
                            opposite layout except these fields

        style --            XFStyle that is applied to all cells in section
        header_style --     XFStyle that is only applied to header cells
        width --            field width
        height --           field height
        num_format --       excel-like cell format

        """
        style = style or self.style
        header_style = header_style or self.header_style

        field = CalcField(self, name, func, fields, fields_ignore, cross_fields, cross_fields_ignore, style, header_style, width, height, num_format)
        field._level = self._level - (1 if self.collapse else 0)
        self._items.append(field)
        return field

    def get_fields(self, field_cls=None):
        """Generates fields of the section (including subfields),
        which type is equal to field_cls.

        Keyword arguments:
        field_cls -- object type to select 
        
        """
        for item in self._items:
            if issubclass(type(item), Field):
                if not field_cls:
                    yield item
                elif type(item) == field_cls:
                    yield item
            elif type(item) == Section:
                for subitem in item.get_fields(field_cls):
                    yield subitem
            else:
                raise ReportException('Unknown item type: %s' % type(item))

    def get_data_fields(self):
        """Alias for get_fields(DataField)"""
        return self.get_fields(field_cls=DataField)

    def get_calc_fields(self):
        """Alias for get_fields(CalcField)"""
        return self.get_fields(field_cls=CalcField)

class Report(object):
    OFFSET = 3

    __slots__ = ['rows', 'cols', 'caption', 'desc', 'cols_width', 'rows_height', 'row_header_style', 'col_header_style', 'cell_style', 'cell_filter', 'merge_styles', 'ignore_none', '_top', '_left', '__fake_cols', '__fake_rows']

    def __init__(self, caption='', desc='', cell_filter=None, merge_styles=True):
        """Creates a report.
        Report is created with two default sections:
        self.rows   -- row section
        self.cols   -- cols section
        
        Fields and section should be added to these default sections.

        Keyword arguments:
        caption --          Report header. Can be splitted by '\n' symbol.
        desc --             Text below the caption.
        cell_filter --      Function that can be used to transform
                            cell values (e.g. to hides zero values)
                            It gets an singular value as input.
        merege_style --     Don't use it. It's not implemented yet.

        """
        self.rows = Section()
        self.cols = Section()
        self._top, self._left = 0, 0
        self.caption = caption
        self.desc = desc
        self.cols_width = 0x00ff
        self.rows_height = 0x00ff
        self.row_header_style = XFStyle()
        self.col_header_style = XFStyle()
        self.cell_style = XFStyle()
        self.__fake_cols = False
        self.__fake_rows = False
        self.cell_filter = cell_filter
        self.merge_styles = merge_styles
        self.ignore_none = True

    def render(self, ws, data, transpose=False):
        """Render the report. Constructs report and writes
        it to worksheet.

        Keyword arguments:
        ws --           xlwt worksheet where the report is drawn
        data --         report data. Shold be the list
                        List of rows is expected. If you have
                        list of columns insted just use transpose feature.
        transpose --    transpose matrix 

        """
        if ws.last_used_row:
            self._top, self._left = ws.last_used_row + self.OFFSET, 0
        self.__draw_caption(ws)
        self.__draw_description(ws)
        if transpose:
            data = transposed(data)
        self.__draw(ws, self.__render(data))
            
    def __draw_caption(self, ws):
        """Draws a report caption"""
        if not self.caption:
            return
        for line in self.caption.splitlines():
            ws.write(self._top, self._left, line, caption_style)
            self._top += 1

    def __draw_description(self, ws):
        """Draws some text after caption"""
        if not self.desc:
            return
        for line in self.desc.splitlines():
            ws.write(self._top, self._left, line, description_style)
            self._top += 1

    def __render(self, data_matrix):
        """Main rendering routine"""
        data = list(data_matrix)
        _data = []
        # Get data rows
        rfields = list(self.rows.get_data_fields())
        cfields = list(self.cols.get_data_fields())

        # Check if at least one field exists
        if not rfields and not cfields:
            raise ReportException('At least one field should be added')

        # Check dimensions compatibility
        # If rfields == 0 then it seem we should build fake fields
        if len(data) != len(rfields) and len(rfields) !=0:
            raise ReportException('Row fields count does not match input data rows count. Expected %s but got %s.' % (len(rfields), len(data)))
        for i, row in enumerate(data):
            # If cfields == 0 then it seem we should build fake fields
            if len(cfields) != len(row) and len(cfields) !=0:
                raise ReportException("Cells count in %sth row do not match input data. Expected %s bot got %s." % (i+1, len(cfields), len(row)))

        # Append fake rows if required
        if not rfields:
            for i in range(len(data)):
                self.rows._add_fake_field()
            self.__fake_rows = True

        # Append fake columns if required
        if not cfields:
            for i in range(len(data[0])):
                self.cols._add_fake_field()
            self.__fake_cols = True

        # Enumerate rows
        for i, row in enumerate(self.rows.get_fields()):
            row.index = i
        for i, col in enumerate(self.cols.get_fields()):
            col.index = i

        # Making result Matrix and fill it with initial data
        key = lambda item: type(item) == DataField
        for i, row in enumerate_if(self.rows.get_fields(), key):
            _row = []
            for j, col in enumerate_if(self.cols.get_fields(), key):
                if type(col) == type(row) == DataField:
                    _row.append(data[i][j])
                else:
                    _row.append(None)
            _data.append(_row)

        # Set calc items for rows
        for row in self.rows.get_calc_fields():
            for col in self.cols.get_fields():
                if col in row.cross_fields and col not in row.cross_fields_ignore:
                    index = []
                    for f in row.sec.get_data_fields():
                        if f in row.fields and f not in row.fields_ignore:
                            index.append((f.index, col.index))
                    _data[row.index][col.index] = _calc(row.func, index, _data, self.ignore_none)

        # Set calc items for columns
        for col in self.cols.get_calc_fields():
             for row in self.rows.get_fields():
                if row in col.cross_fields and row not in col.cross_fields_ignore:
                    index = []
                    for f in col.sec.get_data_fields():
                        if f in col.fields and f not in col.fields_ignore:
                            index.append((row.index, f.index))
                    _data[row.index][col.index] = _calc(col.func, index, _data, self.ignore_none)

        # Second pass: evaluate callable item
        for i, row in enumerate(_data):
            for j, item in enumerate(row):
                if callable(item):
                    _data[i][j] = item()
        return _data

    def __draw(self, ws, data):
        """Main drawing routine"""
        top, left = self._top, self._left

        # Get styles
        column_styles, row_styles = [], []
        for col in self.cols.get_fields():
            column_styles.append((col.style, col.num_format))
        for row in self.rows.get_fields():
            row_styles.append((row.style, row.num_format))

        # Rendering columns
        if not self.__fake_cols:
            size, top_offset, col_headers = self.__render_headers(self.cols, top, left)
        else:
            top_offset, col_headers = 0, []

        # Rendering rows
        if not self.__fake_rows:
            size, left_offset, row_headers = self.__render_headers(self.rows, left, top)
        else:
            left_offset, row_headers = 0, []

        # Drawing columns headers
        for item, r, c, r_size, c_size in col_headers:
            header_style = item.header_style or item.style \
                    or self.col_header_style
            ws.write_merge(r, r + r_size - 1, left_offset + c, left_offset + c + c_size - 1, item.name, header_style)
            if isinstance(item, Field):
                ws.col(left_offset + c).width = item.width or self.cols_width
                if item.height:
                    ws.row(r).height = item.height
                    ws.row(r).height_mismatch = True
                ws.col(left_offset + c).level = item._level
        top += top_offset

        # Drawing rows headers
        for item, c, r, c_size, r_size in row_headers:
            header_style = item.header_style or item.style \
                    or self.row_header_style
            ws.write_merge(top_offset + r, top_offset + r + r_size - 1, c, c + c_size - 1, item.name, header_style)
            if isinstance(item, Field):
                ws.row(top_offset + r).height = item.width or self.cols_width
                ws.row(top_offset + r).level = item._level
        left += left_offset

        # Drawing data
        for i, items in enumerate(data):
            row_style, row_num_format = row_styles[i]
            for j, item in enumerate(items):
                # Here should be styles merging
                col_style, col_num_format = column_styles[j]

                # Apply cell filter
                if self.cell_filter:
                    item = self.cell_filter(item)

                if row_style and col_style and self.merge_styles:
                    cell_style = merge_styles(row_style, col_style)
                else:
                    cell_style = row_style or col_style or self.cell_style

                # TODO: move following to styles merge block
                if row_num_format:
                    cell_style.num_format_str = row_num_format
                elif col_num_format:
                    cell_style.num_format_str = col_num_format

                ws.write(top + i, left + j, item, cell_style)
    
    def __render_headers(self, items, top, left):
        """Render folded headers"""
        size, cells = _head_render(items)
        level_count = max(i[1] for i in cells) + 1

        items = []
        for item, r, c, r_size, c_size in cells:
            # If deep level is zero than it means 
            # all available space should be filled
            if r_size == 0:
                r_size = level_count - r
            items.append((item, r + top, c + left, r_size, c_size))
        return size, level_count, items

styles_cache = {}   # Workout to avoid XFStyle limit

def merge_styles(row_style, col_style, default_style=easyxf('')):
    """Merges row and column style.

    Method tries to get "strongest" style feauters from col (row) style
    and replicate it to row (col) style.

    Alas, it doesn't stable yet.
        
    """

    if (row_style, col_style) in styles_cache:
        new_style = styles_cache[row_style, col_style]
    else:
        new_style = XFStyle()

        # Merge borders
        new_style.borders.top = row_style.borders.top if row_style.borders.top > col_style.borders.top \
            else col_style.borders.top
        new_style.borders.left = row_style.borders.left if row_style.borders.left > col_style.borders.left \
            else col_style.borders.left
        new_style.borders.right = row_style.borders.right if row_style.borders.right > col_style.borders.right \
            else col_style.borders.right
        new_style.borders.bottom = row_style.borders.bottom \
            if row_style.borders.bottom > col_style.borders.bottom else col_style.borders.bottom

        # Merge pattern
        if default_style.pattern.pattern == row_style.pattern.pattern:
            new_style.pattern.pattern = col_style.pattern.pattern
        else:
            new_style.pattern.pattern = row_style.pattern.pattern

        if default_style.pattern.pattern_fore_colour == row_style.pattern.pattern_fore_colour:
            new_style.pattern.pattern_fore_colour = col_style.pattern.pattern_fore_colour
        else:
            new_style.pattern.pattern_fore_colour = row_style.pattern.pattern_fore_colour

        if default_style.pattern.pattern_back_colour == row_style.pattern.pattern_back_colour:
            new_style.pattern.pattern_back_colour = col_style.pattern.pattern_back_colour
        else:
            new_style.pattern.pattern_back_colour = row_style.pattern.pattern_back_colour

        # Merge font
        new_style.font = merge_fonts(row_style.font, col_style.font, default_style.font)
        
        styles_cache[row_style, col_style] = new_style
    return new_style

def merge_fonts(row_font, col_font, default_font):
    fields = [
        'bold', 'charset', 'colour_index', 'escapement', 'family',
        'get_biff_record', 'height', 'italic', 'name',
        'outline', 'shadow', 'struck_out', 'underline',
    ]

    new_font = Font()
    for field in fields:
        if getattr(row_font, field) == getattr(default_font, field):
            setattr(new_font, field, getattr(col_font, field))
        else:
            setattr(new_font, field, getattr(row_font, field))
    return new_font
    
def _head_render(item, level=0, pos=0):
    '''
    Recursive function to render header cells according to sections
    hierarchy. Section won't be drawn if its name is empty.
    Return value: size, [cell(name, x,y,x_size, y_size), cell(...), ...]
    '''
    if isinstance(item, Field):
        return 1, [(item, level, pos, 0, 1)]
    items = []
    size = 0
    # Skip section level if section name is not defined
    if item.name:
        level = level + 1
    for child in item._items:
        child_size, children = _head_render(child, level, pos + size)
        items.extend(children)
        size += child_size
    # Do not add section if name is not defined
    if not item.name:
        return size, items
    return size, [(item, level-1, pos, 1, size)] + items

class _calc(object):
    """Class that represents calculation.
    Do not use directly.
    
    """
    def __init__(self, func, index, data, ignore_none=True):
        self.func = func
        self.index = index
        self.data = data
        self.ignore_none = ignore_none
    def __call__(self, visited=set()):
        # TODO: raise exception on cyclic references
        _data = []
        for i, j in self.index:
            cell = self.data[i][j]
            if callable(cell):
                v = cell(visited)
            else:
                v = cell
            _data.append(v)
        try:
            if self.ignore_none:     # Filter items with None value
                _data = filter(lambda item: item != None, _data)
            result = self.func(_data)
        except Exception, e:
            raise ReportException('Data should be compatible with aggregation function:  %s, %s: %s' % (str(_data), str(self.func), str(e)))
        return result

def enumerate_if(seq, key=lambda: True):
    """Enumerates sequence. If condition doesn't hold
    yields the same number.
    
    """
    count = 0
    for item in seq:
        yield count, item
        if key(item):
            count += 1
            
def transposed(lists):
    """Transpose data matrix
    
    Keyword args:
    lists   --  2d array (list of lists)

    """
    if not lists: return []
    return map(lambda *row: list(row), *lists)
