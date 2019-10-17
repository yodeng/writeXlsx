#!/usr/bin/env python
# coding:utf-8
# oncoh_filter.py

import argparse
import sys
import re


def parseArg():
    parser = argparse.ArgumentParser(description="For oncoH filter.")
    parser.add_argument(
        "-o", "--out", help='output filter file name', required=True, metavar="<str>")
    parser.add_argument(
        "-c", "--corelist", help="Core reported transcript list.", required=True, metavar="<file>")
    parser.add_argument("-i", "--infile",
                        help="anno bed file", required=True, metavar="<file>")
    return parser.parse_args()


def getData(annobed):
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
    return vardata


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


def main():
    args = parseArg()
    vardata = getData(args.infile)
    global coreReportedTrs_noVer
    coreReportedTrs_noVer = getCoreTrs(args.corelist)
    with open(args.out, "w") as fo:
        for i, v in enumerate(vardata):
            if i == 0:
                header = v
                fo.write("\t".join(v) + "\n")
            else:
                data = dict(zip(header, v))
                filter_tag = getInReportTag(data)
                vardata[i][0] = filter_tag
                fo.write("\t".join(vardata[i]) + "\n")


if __name__ == "__main__":
    main()
