#!/usr/bin/env python

"""Requirements Analysis"""

__author__ = 'Michael Meisinger'

import csv
import datetime
import os
import xlwt
from xlsparser import XLSParser
import json

REQ_FILE = "Req_Export_CI_2013-09-10_ver_0-14.xlsx"

TAB_L2 = "L2_CU"
TAB_L3 = "L3_CI"
TAB_L4 = "L4"

class ReqAnalysis(object):
    def __init__(self):
        self.req = {}

    def parse(self, filename):
        # Load from load xlsx file
        if os.path.exists(filename):
            with open(filename, "rb") as f:
                print "Opened", filename
                doc_str = f.read()
                print "Read req file, size=%s", len(doc_str)
                xls_parser = XLSParser()
                self.csv_files = xls_parser.extract_csvs(doc_str)
                print "Parsed req file OK. Tabs=%s" % len(self.csv_files)

        PARSE_TABS = [
            (TAB_L2, "L2"),
            (TAB_L3, "L3"),
            (TAB_L4, "L4"),
        ]
        for tab, name in PARSE_TABS:
            tab_rows = self.csv_files[tab]
            parse_func_name = "_parse_%s" % name
            if hasattr(self, parse_func_name):
                parse_func = getattr(self, parse_func_name)
                reader = csv.DictReader(tab_rows, delimiter=',')
                print "Parsing tab", tab
                self._lnum = 0
                for row in reader:
                    parse_func(row)
                    self._lnum += 1

    def dump_analysis(self, filename=None):
        self._wb = xlwt.Workbook()
        self._worksheets = {}

        l2_req_dict = self.req[TAB_L2]
        l3_req_dict = self.req[TAB_L3]
        l4_req_dict = self.req[TAB_L4]

        ws = self._wb.add_sheet("L3_L4")
        [ws.write(0, col, hdr) for (col, hdr) in enumerate(["SRC_ID", "SRC_Text", "TARG_ID", "TARG_Text", "Parents", "Group"])]
        self._row = 1

        l3_ws = self._wb.add_sheet("L3")
        [l3_ws.write(0, col, hdr) for (col, hdr) in enumerate(["L3_ID", "SRC_Text", "Num L4", "Num Ver", "Num STC", "Num R4", "Num OUT", "Num Other", "Status1", "Status2"])]
        self._row_l3 = 1

        for l3_req in sorted(l3_req_dict.values(), key=lambda x: x["order"]):
            req_id = l3_req["req_id"]
            out_links = l3_req.get("l4_out_links", [])
            l4_count_by_group = {}
            l4_count = 0
            if out_links:
                for i, link in enumerate(out_links):
                    l4_count += 1
                    ws.write(self._row, 0, req_id)
                    value = unicode(l3_req["req_txt"], "latin1")
                    ws.write(self._row, 1, value.encode("ascii", "replace"))
                    ws.write(self._row, 2, link)
                    l4_req = l4_req_dict.get(link, None)
                    if l4_req:
                        value = unicode(l4_req["req_txt"], "latin1")
                        ws.write(self._row, 3, value.encode("ascii", "replace"))
                        ws.write(self._row, 4, l4_req["l3_link_parents"])
                        group = str(l4_req["group"])
                        ws.write(self._row, 5, group)

                        if group not in l4_count_by_group:
                            l4_count_by_group[group] = 0
                        l4_count_by_group[group] += 1
                    else:
                        ws.write(self._row, 3, "ERROR: NOT FOUND")
                    self._row += 1
            else:
                ws.write(self._row, 0, req_id)
                ws.write(self._row, 1, l3_req["req_txt"])
                self._row += 1

            l3_ws.write(self._row_l3, 0, req_id)
            value = unicode(l3_req["req_txt"], "latin1")
            l3_ws.write(self._row_l3, 1, value.encode("ascii", "replace"))
            l3_ws.write(self._row_l3, 2, l4_count)
            l4_cnt_ver = l4_count_by_group.get("0", 0)
            l3_ws.write(self._row_l3, 3, l4_cnt_ver or "")
            l4_cnt_r3 = l4_count_by_group.get("1", 0)
            l3_ws.write(self._row_l3, 4, l4_cnt_r3 or "")
            l4_cnt_r4 = l4_count_by_group.get("2", 0)
            l3_ws.write(self._row_l3, 5, l4_cnt_r4 or "")
            l4_cnt_out = l4_count_by_group.get("5", 0)
            l3_ws.write(self._row_l3, 6, l4_cnt_out or "")
            l4_cnt_off = l4_count - l4_cnt_ver - l4_cnt_r3 - l4_cnt_r4 - l4_cnt_out
            l3_ws.write(self._row_l3, 7, l4_cnt_off or "")

            l3_status = ""
            if not l4_count:
                l3_status = ""
            elif l4_count == l4_cnt_ver:
                l3_status = "VERIFIED"
            elif l4_count == l4_cnt_ver + l4_cnt_r3:
                l3_status = "EXPECTED STC"
            elif l4_count == l4_cnt_ver + l4_cnt_r3 + l4_cnt_r4:
                l3_status = "EXPECTED R4"
            elif l4_count == l4_cnt_out + l4_cnt_off:
                l3_status = "OUT"
            elif l4_cnt_ver:
                l3_status = "PARTIAL"
            else:
                l3_status = "OTHER"
            l3_ws.write(self._row_l3, 8, l3_status)
            l3_req["l3_status"] = l3_status

            if l3_status in ("VERIFIED", "EXPECTED STC", "EXPECTED R4", "PARTIAL"):
                l3_status2 = "ADDRESSED"
            elif not l3_status:
                l3_status2 = ""
            else:
                l3_status2 = "NOT ADDRESSED"
            l3_ws.write(self._row_l3, 9, l3_status2)
            l3_req["l3_status2"] = l3_status2
            self._row_l3 += 1

        ws = self._wb.add_sheet("L2_L3")
        [ws.write(0, col, hdr) for (col, hdr) in enumerate(["SRC_ID", "SRC_Text", "TARG_ID", "TARG_Text", "Parents", "Status1", "Status2"])]
        self._row = 1

        l2_ws = self._wb.add_sheet("L2")
        [l2_ws.write(0, col, hdr) for (col, hdr) in enumerate(["L2_ID", "SRC_Text", "Num L3", "Num Ver", "Num R3", "Num R4", "Num Out", "Num Other", "Num Addressed", "Num Not", "Status1", "Status2"])]
        self._row_l2 = 1

        for l2_req in sorted(l2_req_dict.values(), key=lambda x: x["order"]):
            req_id = l2_req["req_id"]
            out_links = l2_req.get("l3_out_links", [])
            l3_count_by_group = {}
            l3_count = 0
            if out_links:
                for i, link in enumerate(out_links):
                    l3_count += 1
                    ws.write(self._row, 0, req_id)
                    value = unicode(l2_req["req_txt"], "latin1")
                    ws.write(self._row, 1, value.encode("ascii", "replace"))
                    ws.write(self._row, 2, link)
                    l3_req = l3_req_dict.get(link, None)
                    if l3_req:
                        value = unicode(l3_req["req_txt"], "latin1")
                        ws.write(self._row, 3, value.encode("ascii", "replace"))
                        ws.write(self._row, 4, l3_req["l2_link_parents"])
                        group = str(l3_req["l3_status"])
                        ws.write(self._row, 5, group)
                        if group not in l3_count_by_group:
                            l3_count_by_group[group] = 0
                        l3_count_by_group[group] += 1
                        group2 = str(l3_req["l3_status2"])
                        ws.write(self._row, 6, group2)
                        if group2 not in l3_count_by_group:
                            l3_count_by_group[group2] = 0
                        l3_count_by_group[group2] += 1
                    else:
                        ws.write(self._row, 3, "ERROR: NOT FOUND")
                    self._row += 1
            else:
                ws.write(self._row, 0, req_id)
                ws.write(self._row, 1, l2_req["req_txt"])
                self._row += 1

            l2_ws.write(self._row_l2, 0, req_id)
            value = unicode(l2_req["req_txt"], "latin1")
            l2_ws.write(self._row_l2, 1, value.encode("ascii", "replace"))
            l2_ws.write(self._row_l2, 2, l3_count)
            l4_cnt_ver = l3_count_by_group.get("VERIFIED", 0)
            l2_ws.write(self._row_l2, 3, l4_cnt_ver or "")
            l4_cnt_r3 = l3_count_by_group.get("EXPECTED STC", 0)
            l2_ws.write(self._row_l2, 4, l4_cnt_r3 or "")
            l4_cnt_r4 = l3_count_by_group.get("EXPECTED R4", 0)
            l2_ws.write(self._row_l2, 5, l4_cnt_r4 or "")
            l4_cnt_out = l3_count_by_group.get("OUT", 0)
            l2_ws.write(self._row_l2, 6, l4_cnt_out or "")
            l4_cnt_off = l3_count - l4_cnt_ver - l4_cnt_r3 - l4_cnt_r4 - l4_cnt_out
            l2_ws.write(self._row_l2, 7, l4_cnt_off or "")
            l3_cnt_add = l3_count_by_group.get("ADDRESSED", 0)
            l2_ws.write(self._row_l2, 8, l3_cnt_add or "")
            l3_cnt_nadd = l3_count_by_group.get("NOT ADDRESSED", 0)
            l2_ws.write(self._row_l2, 9, l3_cnt_nadd or "")

            l2_status = ""
            if not l3_count:
                l2_status = ""
            elif l3_count == l4_cnt_ver:
                l2_status = "VERIFIED"
            elif l3_count == l4_cnt_ver + l4_cnt_r3:
                l2_status = "EXPECTED STC"
            elif l3_count == l4_cnt_ver + l4_cnt_r3 + l4_cnt_r4:
                l2_status = "EXPECTED R4"
            elif l3_count == l4_cnt_out + l4_cnt_off:
                l2_status = "OUT"
            elif l4_cnt_ver:
                l2_status = "PARTIAL"
            else:
                l2_status = "OTHER"
            l2_ws.write(self._row_l2, 10, l2_status)
            l2_req["l2_status"] = l2_status

            if not l3_count:
                l2_status2 = ""
            elif l3_cnt_add:
                l2_status2 = "ADDRESSED"
            else:
                l2_status2 = "NOT ADDRESSED"
            l2_ws.write(self._row_l2, 11, l2_status2)
            l2_req["l2_status"] = l2_status2

            self._row_l2 += 1

        dtstr = datetime.datetime.today().strftime('%Y%m%d_%H%M%S')
        path = filename or "output/reqanalysis_%s.xls" % dtstr
        self._wb.save(path)

    def do_all(self):
        self.parse(REQ_FILE)
        self.dump_analysis()

    # -------------------------------------------------------------------------

    def _add_req(self, level, req_id, req):
        if not level in self.req:
            self.req[level] = {}
        req_dict = self.req[level]
        if req_id in req_dict:
            print "WARNING: Duplicate %s" % req_id
        req_dict[req_id] = req

    def _build_req_links(self, field, prefix):
        result = []
        if field:
            #print "_build_req_links", json.dumps(field)
            field = str(field)
            links = field.splitlines()
            result.extend(prefix + link.strip() for link in links)
        return result

    def _add_req_links(self, req_id, links, targ, targ_attr):
        """Adds given req_id to target requirement for each link in links"""
        targ_req_dict = self.req[targ]
        for link in links:
            targ_req = targ_req_dict.get(link, None)
            if not targ_req:
                print " ERROR: Link %s target does not exist: %s/%s " % (req_id, targ, link)
                continue

            if targ_attr not in targ_req:
                targ_req[targ_attr] = []
            targ_link_list = targ_req[targ_attr]
            if req_id in targ_link_list:
                print " WARNING: Link to %s/%s already present: %s" % (targ, link, req_id)
            targ_link_list.append(req_id)

    # -------------------------------------------------------------------------

    def _parse_L4(self, row):
        item_class = row["Item Class"]
        if item_class not in ["Approved Req", "Approved Int"]:
            return

        req_id = row["ID"]
        l3_links = self._build_req_links(row["L3 Link"], "L3-CI-RQ-")
        req_dict = dict(
            req_id=req_id,
            req_txt=row["Requirement Statement"],
            item_type=row["Item Type"],
            l3_links=l3_links,
            l3_link_parents=len(l3_links),
            group=row["Group"],
            order=self._lnum
        )
        self._add_req(TAB_L4, req_id, req_dict)
        self._add_req_links(req_id, l3_links, TAB_L3, "l4_out_links")

    def _parse_L3(self, row):
        req_id = row["ID"]
        l2_links = self._build_req_links(row["L2_CU"], "L2-CU-RQ-")
        req_dict = dict(
            req_id=req_id,
            req_txt=row["Requirement Statement"],
            l2_links=l2_links,
            l2_link_parents=len(l2_links),
            order=self._lnum
        )
        self._add_req(TAB_L3, req_id, req_dict)
        self._add_req_links(req_id, l2_links, TAB_L2, "l3_out_links")

    def _parse_L2(self, row):
        req_id = row["ID"]
        req_dict = dict(
            req_id=req_id,
            req_txt=row["Requirement Statement"],
            order=self._lnum
        )
        self._add_req(TAB_L2, req_id, req_dict)

if __name__ == '__main__':
    ra = ReqAnalysis()
    ra.do_all()
