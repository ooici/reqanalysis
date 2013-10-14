#!/usr/bin/env python

"""Requirements Tracing Generator.

USAGE: python reqgen.py <in_filename> <out_filename>

Prerequisites: xtwt, xlrd in virtualenv
"""

__author__ = 'Michael Meisinger'

import csv
import datetime
import os
import sys
import xlwt
from xlsparser import XLSParser

REQ_FILE = "Deliverabe-Milestone-Requirement_Mapping_V02.xlsx"
OUT_FILE_PREFIX = "output/reqanalysis"
OUT_TRACE_PREFIX = "output/tracing"

TAB_TRACING = "Example v2"

HTABLE_START = """
<div class="panel" style="border-width: 1px;">
  <div class="panelContent">
    <h3>%%TITLE%%</h3>
    <div class='table-wrap'>
      <table class='confluenceTable'>"""

HTABLE_END = """
      </table>
    </div>
  </div>
</div>"""

HTABLE_ROW_START = """
        <tr>"""

HTABLE_ROW_END = """
        </tr>"""

HTABLE_HEAD_ROW = """
          <th class='confluenceTh'>%%TEXT%%</th>"""

HTABLE_ROW = """
          <td class='confluenceTd'>%%TEXT%%</td>"""

HTABLE_SEP = """
<p>
  <br class="atl-forced-newline" />
</p>"""


class ReqAnalysis(object):
    def __init__(self):
        self.req = {}

    def parse(self, filename):
        # Load from load xlsx file
        if os.path.exists(filename):
            with open(filename, "rb") as f:
                print "Opened", filename
                doc_str = f.read()
                print "Read req file, size=%s" % len(doc_str)
                xls_parser = XLSParser()
                self.csv_files = xls_parser.extract_csvs(doc_str)
                print "Parsed req file OK. Found %s tabs." % len(self.csv_files)
        else:
            print "ERROR: Requirements file %s does not exist" % filename
            sys.exit(1)

        PARSE_TABS = [
            (TAB_TRACING, "tracing"),
        ]
        for tab, name in PARSE_TABS:
            tab_rows = self.csv_files[tab]
            parse_func_name = "_parse_%s" % name
            if hasattr(self, parse_func_name):
                parse_func = getattr(self, parse_func_name)
                reader = csv.DictReader(tab_rows, delimiter=',')
                print "Parsing tab", tab
                self._lnum = 0
                self._add_cnt = 0
                for row in reader:
                    res = parse_func(row)
                    self._lnum += 1
                    if res:
                        self._add_cnt += 1
                print " ...using %s of %s rows" % (self._add_cnt, self._lnum)

    def dump_trace_files(self):
        for ms_id in sorted(self.req):
            ms_list = self.req[ms_id]

            if not os.path.exists(OUT_TRACE_PREFIX):
                os.makedirs(OUT_TRACE_PREFIX)

            ms_filename = "%s/%s.html" % (OUT_TRACE_PREFIX, ms_id)
            with open(ms_filename, "w") as f:
                ms_item = ms_list[0]
                f.write(HTABLE_START.replace("%%TITLE%%", "Requirements for %s (%s)" % (ms_id, ms_item["subject"])))
                f.write(HTABLE_ROW_START)
                f.write(HTABLE_HEAD_ROW.replace("%%TEXT%%", "Relationship"))
                f.write(HTABLE_HEAD_ROW.replace("%%TEXT%%", "Object"))
                f.write(HTABLE_HEAD_ROW.replace("%%TEXT%%", "Sub-Relationship"))
                f.write(HTABLE_HEAD_ROW.replace("%%TEXT%%", "Sub-Object"))
                f.write(HTABLE_HEAD_ROW.replace("%%TEXT%%", "Description"))
                f.write(HTABLE_ROW_END)

                for item in ms_list[1:]:
                    f.write(HTABLE_ROW_START)
                    f.write(HTABLE_ROW.replace("%%TEXT%%", item["rel1"]))
                    f.write(HTABLE_ROW.replace("%%TEXT%%", item["object"]))
                    f.write(HTABLE_ROW.replace("%%TEXT%%", item["rel2"]))
                    f.write(HTABLE_ROW.replace("%%TEXT%%", item["subobj"]))
                    f.write(HTABLE_ROW.replace("%%TEXT%%", item["desc"]))
                    f.write(HTABLE_ROW_END)

                f.write(HTABLE_END)

    def do_all(self, in_filename=None, out_filename=None):
        self.parse(in_filename or REQ_FILE)

        self.dump_trace_files()

    # -------------------------------------------------------------------------

    def _add_req(self, level, req_id, req):
        if not level in self.req:
            self.req[level] = []
        req_dict = self.req[level]
        req_dict.append(req)


    # -------------------------------------------------------------------------

    def _parse_tracing(self, row):
        activated = row["Activated"]
        sub_dom = row["Subject Domain"]

        if str(activated) != "1":
            return
        if sub_dom != "Milestone":
            return

        ms_id = row["Subject ID"]
        item_id = row["Sort ID"]
        ms_dict = dict(
            ms_id=ms_id,
            item_id=item_id,
            subject=row["Subject Title"],
            rel1=row["Relationship"],
            object=row["Object Title"],
            rel2=row["Sub-Relationship"],
            subobj=row["Subobject Title"],
            desc=row["Description"],
            order=self._lnum
        )
        self._add_req(ms_id, item_id, ms_dict)
        return True


if __name__ == '__main__':
    in_filename = sys.argv[1] if len(sys.argv) >= 2 else None
    out_filename = sys.argv[2] if len(sys.argv) >= 3 else None

    ra = ReqAnalysis()
    ra.do_all(in_filename, out_filename)
