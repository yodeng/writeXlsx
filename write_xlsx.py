#!/usr/bin/env python
# coding:utf-8

import argparse
import itertools
import sys
import csv
import re
import os

from xlsxwriter.workbook import Workbook
from collections import defaultdict, OrderedDict

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


def getExpressBool(var, expression):
    if not expression:
        return False
    if isinstance(expression, list):
        if var in expression:
            return True
        else:
            return False
    token_re = re.compile(r'"(?:[^"]|"")*"|\S+')
    tokens = token_re.findall(expression)
    if not (len(tokens) == 7 or len(tokens) == 3):
        warn("Incorrect number of tokens in criteria '%s'" % expression)
        sys.exit(1)
    new_tokens = []
    for i, token in enumerate(tokens):
        token = token.strip('\'"')
        if i == 3:
            if re.match('(and|&&)', token):
                token = "and"
            elif re.match(r'(or|\|\|)', token):
                token = "or"
        elif i == 2 or i == 6:
            try:
                f = float(token)
            except:
                token = "'" + token + "'"
        new_tokens.append(token)
    try:
        x = float(var)
    except:
        x = var.strip()
    return eval(" ".join(new_tokens))


def parseFile(infile, indir):
    d = OrderedDict()
    with open(infile, "rb") as fi:
        for line in fi:
            if not line.strip() or line.strip().startswith("#"):
                continue
            if line.strip().startswith("[") and line.strip().endswith("]"):
                sheetname = line.strip().strip('[]').strip()
                continue
            if sheetname:
                if sheetname in d:
                    d[sheetname]["sep"] = line.strip().strip('"').strip("'")
                else:
                    d[sheetname] = {}
                    d[sheetname]["file"] = os.path.join(indir, line.strip())
            else:
                continue
    return d


def writexlsx(data, sheetname=None, fmt={}, bg_color=None, filterule={}, even=1, boldheader=True, autofilter=True):
    worksheet = workbook.add_worksheet(name=sheetname)
    header = data[0]
    fmtda = {}
    if len(fmt):
        for c, col in enumerate(header):
            wds = fmt.get(col, {}).get("width")
            ops = {i: fmt.get(col, {}).get(i)
                   for i in ["hidden", "level", "collapsed"]}
            cfmt = {}
            for i in PROPERTY:
                v = fmt.get(col, {}).get(i)
                if v is not None:
                    cfmt[i] = v
            fmtda[c] = cfmt
            if len(cfmt):
                worksheet.set_column(c, c, width=wds, options=ops,
                                     cell_format=workbook.add_format(dict(cfmt)))
            else:
                worksheet.set_column(c, c, width=wds, options=ops)
    if boldheader:
        worksheet.write_row(
            0, 0, header, cell_format=workbook.add_format({'bold': True}))
    else:
        worksheet.write_row(0, 0, header)
    for r, row_data in enumerate(data[1:]):
        r += 1
        for c, col in enumerate(row_data):
            fmtd = fmtda.get(c, {}).copy()
            if c == 0 and header[0] == "inReport" and (sampleId is not None):
                fmtd.update({"font_color": "blue", "underline": 1})
                filterlinedict = dict(zip(header, row_data))
                if (r % 2 == even) and (bg_color is not None):
                    fmtd["bg_color"] = bg_color
                url = interp_local + sampleId + "&p=" + \
                    filterlinedict["Chr"] + ':' + \
                    filterlinedict["Start"] + '-' + filterlinedict["Stop"]
                worksheet.write_url(
                    r, 0, url, string=col, tip="ReadsPlot Info", cell_format=workbook.add_format(fmtd))
                continue
            if col.startswith("http://"):
                fmtd.update({"font_color": "black", "underline": 0})
                if (r % 2 == even) and (bg_color is not None):
                    fmtd["bg_color"] = bg_color
                worksheet.write(r, c, col, workbook.add_format(fmtd))
            else:
                if (r % 2 == even) and (bg_color is not None):
                    fmtd["bg_color"] = bg_color
                if len(fmtd):
                    worksheet.write(r, c, col, workbook.add_format(fmtd))
                else:
                    worksheet.write(r, c, col)

    ncol, nrow = worksheet.dim_colmax, worksheet.dim_rowmax
    if autofilter:
        worksheet.autofilter(0, 0, nrow, ncol)
    if len(filterule):
        for colnames, express in filterule.items():
            try:
                colindex = header.index(colnames)
            except ValueError:
                continue
            if isinstance(express, list):
                worksheet.filter_column_list(colindex, express)
            else:
                worksheet.filter_column(colindex, express)
            colvalue = [i[colindex] for i in data]
            row = 1
            for var in colvalue[1:]:
                if not getExpressBool(var, express):
                    worksheet.set_row(row, options={'hidden': True})
                row += 1


def parseRule(rulefile):
    rule = defaultdict(lambda: defaultdict(None))
    if rulefile is None:
        return rule
    with open(rulefile) as rf:
        for line in rf:
            if not line.strip() or line.strip().startswith("#"):
                continue
            if line.strip().startswith("[") and line.strip().endswith("]"):
                sheetname = line.strip().strip('[]').strip()
                continue
            h,op,expres = re.findall('{(.+?)}',line)            
            h = h.strip()            
            op = op.strip()
            if h in rule:
                sys.exit(1)  # dup filter header error
            if op == "@@":
                rule[sheetname][h] = expres.strip()
            elif op == "##":
                rule[sheetname][h] = eval(
                    "[" + expres.strip().strip("()") + "]")
            else:
                print 'illegal operations, only "##" or "@@" allowed'
                sys.exit(1)
    return rule


def getVfmort(vf):
    vformat = {}
    rformat = {}
    if vf is None:
        return vformat, rformat
    ele = False
    with open(vf) as fi:
        for line in fi:
            if not line.strip() or line.strip().startswith("#"):
                continue
            if line.strip().startswith("[[") and line.strip().endswith("]]"):
                ele = line.strip().strip("[]").strip()
                continue
            if line.strip().startswith("[") and line.strip().endswith("]"):
                sheetname = line.strip().strip('[]').strip()
                h = False
                continue
            if (not h) and ele == "column":
                h = line.strip("\n").strip("#").split("\t")
                continue
            if ele == "row":
                k, v = line.split("=")
                k = k.strip()
                v = v.strip()
                if len(v):
                    if (v.isdigit() and int(v) == 0) or v.capitalize() == "False":
                        rformat.setdefault(sheetname, {})[k] = 0
                    elif (v.isdigit() and int(v) == 1) or v.capitalize() == "True":
                        rformat.setdefault(sheetname, {})[k] = 1
                    else:
                        assert k not in ["bold_header","auto_filter"], "Only 0/False or 1/True allowed for 'bold_header' and 'auto_filter'"                            
                        rformat.setdefault(sheetname, {})[k] = v
            elif ele == "column":
                line = line.strip("\n").split("\t")
                for i in range(len(line[1:])):
                    i += 1
                    v = line[i]
                    try:
                        v = int(v)
                    except ValueError:
                        pass
                    line[i] = v
                    vformat.setdefault(sheetname, {}).setdefault(
                        line[0].strip('"'), {})[h[i]] = line[i]
            else:
                continue
    return vformat, rformat


def parseArg():
    parser = argparse.ArgumentParser(
        description="For output formatted and filtered excel.")
    parser.add_argument("-f", "--format", help='format of sheet',
                        metavar="<file>")
    parser.add_argument(
        "-o", "--out", help='excel output filename, ".xlsx" will be added to your name if not specified.', required=True, metavar="<str>")
    parser.add_argument(
        "-r", "--rule", help="filter rule list to sheet table.", metavar="<file>")
    parser.add_argument("-b", "--inputdir", help='input sheet file directory',
                        required=True, metavar="<dir>")
    parser.add_argument("-i", "--infile", help='input sheet name and file',
                        required=True, metavar="<file>")
    parser.add_argument("-s", "--samplename",
                        help="sample name",required = True, metavar="<str>")
    return parser.parse_args()


def main():
    args = parseArg()
    outxlsx = args.out if args.out.endswith(".xlsx") else args.out + ".xlsx"
    global workbook, sampleId, interp_local
    workbook = workbook = Workbook(
        outxlsx, {'strings_to_numbers': True, 'strings_to_urls': True})
    sampleId = args.samplename
    interp_local = "http://10.0.1.7:4384/cgi-bin/plot/plotReads.cgi?f=1&c=1&b="
    sepDict = {"tsv": "\t", "csv": ","}
    rule = parseRule(args.rule)
    fmt, stylefile = getVfmort(args.format)
    sheetfiles = parseFile(args.infile, args.inputdir)
    for sheetName, sheetfile in sheetfiles.items():
        sheet_file = re.sub("{SAMP}",sampleId,sheetfile["file"])
        assert os.path.isfile(sheet_file), 'File {} is not found'.format(sheet_file)
        data = []
        sep = sheetfile.get("sep", "csv")
        with open(sheet_file, "rb") as fi:
            for row in csv.reader(fi, delimiter=sepDict.get(sep, ",")):
                data.append([c.decode("utf-8") for c in row])
        sheeFmt = fmt.get(sheetName, {})
        sheetRule = rule.get(sheetName, {})
        l1 = stylefile.get(sheetName, {}).get("single_line_color")
        l2 = stylefile.get(sheetName, {}).get("even_line_color")
        if l1 and (l2 is None):
            c = l1
            e = 0
        elif l2 and (l1 is None):
            c = l2
            e = 1
        else:
            c = None
            e = 1
        b = stylefile.get(sheetName, {}).get("bold_header", 1)
        f = stylefile.get(sheetName, {}).get("auto_filter", 1)
        writexlsx(data, sheetname=sheetName, fmt=sheeFmt,
                  bg_color=c, filterule=sheetRule, even=e, boldheader=b, autofilter=f)
    workbook.close()


if __name__ == "__main__":
    main()

