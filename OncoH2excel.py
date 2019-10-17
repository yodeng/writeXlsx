#!/usr/bin/env python
# coding:utf-8

import argparse
import sys
import csv
import re
import os

from xlsxwriter.workbook import Workbook
from warnings import warn


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


def txt2xlsx(txtfile, sheetname=None, sep=",", header=True, bg_color=None):
    if sheetname is None:
        sheetname = workbook.sheet_name + str(workbook.sheetname_count + 1)
    infile = open(txtfile, "rb")
    worksheet = workbook.add_worksheet(name=sheetname)
    fbg = [None, workbook.add_format({'bg_color': bg_color}) if (
        bg_color is not None) else None]
    for r, row in enumerate(csv.reader(infile, delimiter=sep)):
        if row[-1] == "N" and sheetname == "norm_sv":
            continue
        for c, col in enumerate(row):
            if col.startswith("http://"):
                if r % 2:
                    hpft = hyperlinkFmt(font_color="black",
                                        underline=0, bg_color=bg_color)
                else:
                    hpft = hyperlinkFmt(font_color="black", underline=0)
                worksheet.write_url(r, c, col.decode(
                    "utf-8"), hpft)
                continue
            worksheet.write(r, c, col.decode("utf-8"), fbg[r % 2])
    infile.close()
    if header:
        worksheet.set_row(0, None, workbook.add_format({'bold': True}))
        ncol, nrow = worksheet.dim_colmax, worksheet.dim_rowmax
        worksheet.autofilter(0, 0, nrow, ncol)
    if sheetname == "version" and os.path.isfile(os.path.join(workdir, "run.sh")):
        with open(os.path.join(os.getcwd(), "run.sh")) as cmdfile:
            cmd = cmdfile.readline().strip()
        worksheet.write_row(r+1, 0, ["CMD", cmd])


def var2xlsx(annobed, sheetname=None, bg_color=None):
    vardata = []
    ab = open(annobed, "rb")
    header = False
    addInreport = False
    hasInexcel = False
    nm2np = {}
    for line in ab:
        if not line.strip() or line.startswith("##") or len(line.split("\t")) < 6:
            continue
        else:
            line = line.strip().split("\t")
            if not header:
                if line[0].startswith("#"):
                    line[0] = line[0].strip("#")
                fistline = line[:]
                line.pop(5)
                if line[0] != "inReport":
                    line.insert(0, "inReport")
                if "InExcel" in line:
                    line.pop(line.index("InExcel"))
                    hasInexcel = True
                header = line
                if "Transcript" not in header or "Protein" not in header:
                    print "Error: no found 'Transcript' and 'Protein' columns in it."
                    sys.exit(1)
                vardata.append(header)
                continue
            linedict = dict(zip(fistline, line))
            sample_id = line.pop(5)
            if sample_id == "nullSample" or (hasInexcel and not re.search("1|Y", linedict["InExcel"])) or (linedict.get("A.Ratio", 1) == "." or linedict.get("A.Depth", 1) == "."):
                continue
            if re.match("^[NX]M_", linedict["Transcript"]):
                linedict["Transcript"] = re.sub(
                    "-\d+$", "", linedict["Transcript"])
                nm2np[linedict["Transcript"]] = linedict["Protein"]
            data = [linedict.get(i, "inReport") for i in header]
            vardata.append(data)
    ab.close()
    worksheet = workbook.add_worksheet(name=sheetname)
    fbg = [workbook.add_format({'bg_color': bg_color}) if (
        bg_color is not None) else None, None]
    worksheet.write_row(
        0, 0, vardata[0], cell_format=workbook.add_format({'bold': True}))
    for r, row in enumerate(vardata[1:]):
        r += 1
        for c, col in enumerate(row):
            title = vardata[0][c]
            wds = VarFormat.get(title, {}).get("width")
            ops = {i: VarFormat.get(title, {}).get(i)
                   for i in ["hidden", "level", "collapsed"]}
            worksheet.set_column(c, c, width=wds, options=ops)
            if c == 0:
                filterlinedict = dict(zip(header, vardata[r]))
                filter_tag = getInReportTag(filterlinedict)
                url = interp_local + sampleId + "&p=" + \
                    filterlinedict["Chr"] + ':' + \
                    filterlinedict["Start"] + '-' + filterlinedict["Stop"]
                ufmt = {"font_color": "blue", "underline": 1}
                if r % 2 == 0:
                    ufmt["bg_color"] = bg_color
                worksheet.write_url(
                    r, 0, url, string=filter_tag, tip="ReadsPlot Info", cell_format=workbook.add_format(ufmt))
                continue
            worksheet.write(r, c, col.decode("utf-8"), fbg[r % 2])
    ncol, nrow = worksheet.dim_colmax, worksheet.dim_rowmax
    worksheet.autofilter(0, 0, nrow, ncol)
    for colnames, express in RULE.items():
        colindex = header.index(colnames)
        if isinstance(express, list):
            worksheet.filter_column_list(colindex, express)
        else:
            worksheet.filter_column(colindex, express)
        colvalue = [i[colindex] for i in vardata]
        row = 1
        for var in colvalue[1:]:
            if not getExpressBool(var, express):
                worksheet.set_row(row, options={'hidden': True})
            row += 1


def getVfmort(vf):
    vformat = {}
    with open(vf) as fi:
        h = fi.next().strip().strip("#").split("\t")
        for line in fi:
            line = line.strip().split("\t")
            vformat[line[0].strip('"')] = dict(zip(h[1:], map(int, line[1:])))
    return vformat


def getCoreTrs(corelist):
    d = {}
    with open(corelist) as fi:
        for line in fi:
            if not line.strip():
                continue
            items = line.split()
            items[0] = re.sub("\..+$", "", items[0])
            d[items[0]] = 1
    return d


def getInReportTag(anno):
     rAnnoItms = {}
     for i, v in anno.items():
        try:
            v = float(v)
        except:
            pass
        rAnnoItms[i] = v
    ftag = "InExcel"
    if not all([rAnnoItms.has_key(i) for i in ["A.Depth", "RepeatTag", "A.Ratio", "ExAC AF", "1000G AF", "ExAC EAS AF", "1000G EAS AF", "Panel AlleleFreq", "Transcript"]]):
        print "Error: Header String Error for InReportTag assignment, please check"
        sys.exit(1)
    coreTr = rAnnoItms["Transcript"]
    coreTr = re.sub("\..+$", "", coreTr)
    adOpt = 1 if rAnnoItms["A.Depth"] > 20 else 0
    repeatOpt = 0 if rAnnoItms["RepeatTag"] == '.' else 1
    arOpt = 0
    if repeatOpt:
        if rAnnoItms["A.Ratio"] > 0.3:
            arOpt = 1
    else:
        if rAnnoItms["A.Ratio"] > 0.2:
            arOpt = 1
    popOpt = (rAnnoItms["ExAC AF"] == '.' or rAnnoItms["ExAC AF"] < 0.01) and (rAnnoItms["ExAC EAS AF"] == '.' or rAnnoItms["ExAC EAS AF"] < 0.01) and (
        rAnnoItms["1000G AF"] == '.' or rAnnoItms["1000G AF"] < 0.01) and (rAnnoItms["1000G EAS AF"] == '.' or rAnnoItms["1000G EAS AF"] < 0.01)
    panelOpt = rAnnoItms["Panel AlleleFreq"] == '.' or rAnnoItms["Panel AlleleFreq"] < 0.03
    trOpt = coreTr in coreReportedTrs_noVer
    if adOpt and arOpt and not(popOpt and not(panelOpt)):
        ftag = 'Pass'
        if trOpt:
            ftag = 'HHReport'
            if popOpt and panelOpt:
                ftag = 'DHReport'
    return ftag


def parseArg():
    parser = argparse.ArgumentParser(description="For oncoH to excel.")
    parser.add_argument("-f", "--vformat", help='format of variation sheet, default no extra formatting, will only export rows with value 1 or "Y" in "InExcel" column. this column will be ommitted in Excel.', metavar="<file>", required=True)
    parser.add_argument(
        "-o", "--out", help='excel output filename, ".xlsx" will be added to your name if not specified.', required=True, metavar="<str>")
    parser.add_argument(
        "-g", "--geneinfo", help="add a sheet named 'geneInfo' to show gene coverage information.", metavar="<file>")
    parser.add_argument(
        "--version", help="add a sheet named 'version' to show software information.", metavar="<file>")
    parser.add_argument(
        "-a", "--qcbam", help="QC result of sample bam (csv)", metavar="<file>")
    parser.add_argument(
        "-d", "--ctdrug", help="Drug hits information (csv)", metavar="<file>")
    parser.add_argument(
        "-w", "--cnvhf", help="Cnv result of sample bam (csv)", metavar="<file>")
    parser.add_argument(
        "-S", "--svhf", help="SV result of sample bam (csv)", metavar="<file>")
    parser.add_argument(
        "-c", "--corelist", help="Core reported transcript list.", required=True, metavar="<file>")
    parser.add_argument(
        "-r", "--ruleset", help="filter rule list to variation table.", required=True, metavar="<file>")
    parser.add_argument("-n", "--samplename",
                        help="sample name", required=True, metavar="<str>")
    parser.add_argument(
        "--hla", help="add a sheet named 'hla' to show hla typing result information.", metavar="<file>")
    parser.add_argument("-ab", "--anno_bed",
                        help="anno bed file", required=True, metavar="<file>")
    parser.add_argument("-wd", "--workdir", help="work directory, default: $PWD",
                        default=os.getcwd(), metavar="<dir>")
    return parser.parse_args()


def hyperlinkFmt(**kwargs):
    hfmt = workbook.add_format(kwargs)
    return hfmt


def parseRule(rulefile):
    rule = {}
    with open(rulefile) as rf:
        for line in rf:
            if not line.strip() or line.startswith("#"):
                continue
            h, expres = re.split('##|@@', line.strip())
            h = h.strip()
            expres = expres.strip()
            if h in rule:
                sys.exit(1)  # dup filter header error
            if "@@" in line:
                rule[h] = expres.strip()
            elif "##" in line:
                rule[h] = eval("[" + expres.strip().strip("()") + "]")
    return rule


def main():
    args = parseArg()
    outxlsx = args.out if args.out.endswith(".xlsx") else args.out + ".xlsx"

    global workbook, RULE, coreReportedTrs_noVer, interp_local, sampleId, workdir, VarFormat
    coreReportedTrs_noVer = getCoreTrs(args.corelist)
    workbook = Workbook(
        outxlsx, {'strings_to_numbers': False, 'strings_to_urls': False})
    interp_local = "http://10.0.1.7:4384/cgi-bin/plot/plotReads.cgi?f=1&c=1&b="
    sampleId = args.samplename
    RULE = parseRule(args.ruleset)
    workdir = args.workdir
    VarFormat = getVfmort(args.vformat)

    var2xlsx(args.anno_bed, "SmallVariations", bg_color="90EE90")

    qc_bam = args.qcbam
    if qc_bam:
        txt2xlsx(qc_bam, sheetname="QC")

    ct_drug = args.ctdrug
    if ct_drug:
        txt2xlsx(ct_drug, sheetname="Drug", bg_color="#87CEEB")

    hcnv = args.cnvhf
    if hcnv:
        txt2xlsx(hcnv, sheetname="norm_cnv", bg_color="#ADD8E6")

    hsv = args.svhf
    if hsv:
        txt2xlsx(hsv, sheetname="norm_sv", bg_color="#98FB98")

    gene_info = args.geneinfo
    if gene_info:
        txt2xlsx(gene_info, sheetname="geneInfo", sep="\t", bg_color="#F0E68C")

    hla_rst = args.hla
    if hla_rst:
        txt2xlsx(hla_rst, sheetname="hla", sep="\t", bg_color="#7FFFD4")

    version_file = args.version
    if version_file:
        txt2xlsx(version_file, sheetname="version", sep="\t", header=False)

    workbook.close()


if __name__ == "__main__":
    main()
