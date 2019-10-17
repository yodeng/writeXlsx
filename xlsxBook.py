#!/usr/bin/env python
# coding:utf-8

import cvs
from xlsxwriter.workbook import Workbook


class xlsxBook(object):
    PROPERTY = [
        'font_name',
        'font_size',
        'font_color',
        'bold',
        'italic',
        'underline',
        'font_strikeout',
        'font_script',
        'num_format',
        'locked',
        # 'hidden',
        'align',
        'valign',
        'rotation',
        'text_wrap',
        'reading_order',
        'text_justlast',
        'center_across',
        'indent',
        'shrink',
        'pattern',
        'bg_color',
        'fg_color',
        'border',
        'bottom',
        'top',
        'left',
        'right',
        'border_color',
        'bottom_color',
        'top_color',
        'left_color',
        'right_color',
    ]

    OPTIONS = [
        'width',
        'hidden',
        'level'	,
        'collapsed',
    ]

    def __init__(self, filename, **kwargs):
        self._path = filename
        if os.path.isfile:
            print "File %s exists, Exit" % self._path
        initformat = {'strings_to_numbers': True, 'strings_to_urls': True}
        ## init workbook
        self.workbook = Workbook(self._path, initformat.update(kwargs))
        self.sheets = 0

    def addsheet(self, sheetfile, name=None, fmt=None, bg_color=None, even=1, boldheader=True, autofilter=True):

        if name is None:
            name = "Sheet%d" % (self.sheets)
        data = []
        sep = xlsxBook.detectDelimiter(sheetfile)
        with open(sheetfile) as fi:
            for line in fi:
                if not line.strip():
                    continue
                data.append(line.split(sep))

        worksheet = self.workbook.add_worksheet(name=name)
        header = data[0]
        fmtda = {}
        if fmt:
            for c, col in enumerate(header):
                wds = fmt.get(col, {}).get("width")
                ops = {i: fmt.get(col, {}).get(i)
                       for i in ["hidden", "level", "collapsed"]}
                cfmt = {}
                for i in xlsxBook.PROPERTY:
                    v = fmt.get(col, {}).get(i)
                    if v is not None:
                        cfmt[i] = v
                fmtda[c] = cfmt
                if len(cfmt):
                    worksheet.set_column(c, c, width=wds, options=ops,
                                         cell_format=self.workbook.add_format(dict(cfmt)))
                else:
                    worksheet.set_column(c, c, width=wds, options=ops)
        if boldheader:
            worksheet.write_row(
                0, 0, header, cell_format=self.workbook.add_format({'bold': True}))
        else:
            worksheet.write_row(0, 0, header)
        for r, row_data in enumerate(data[1:]):
            r += 1
            for c, col in enumerate(row_data):
                fmtd = fmtda.get(c, {}).copy()
                if (r % 2 == even) and (bg_color is not None):
                    fmtd["bg_color"] = bg_color
                if len(fmtd):
                    worksheet.write(r, c, col, self.workbook.add_format(fmtd))
                else:
                    worksheet.write(r, c, col)

        ncol, nrow = worksheet.dim_colmax, worksheet.dim_rowmax
        if autofilter:
            worksheet.autofilter(0, 0, nrow, ncol)
        self.sheets += 1

    @staticmethod
    def detectDelimiter(filepath, bitysize=32):
        with open(filepath, "rb") as fi:
            ctx = fi.read(1024*int(bitysize))
        dialect = csv.Sniffer().sniff(ctx)
        return dialect.delimiter

    def close(self):
        self.workbook.close()
