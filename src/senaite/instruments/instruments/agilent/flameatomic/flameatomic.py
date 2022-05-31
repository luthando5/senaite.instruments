# -*- coding: utf-8 -*-
#
# This file is part of SENAITE.INSTRUMENTS.
#
# SENAITE.CORE is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation, version 2.
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details.
#
# You should have received a copy of the GNU General Public License along with
# this program; if not, write to the Free Software Foundation, Inc., 51
# Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
#
# Copyright 2018-2019 by it's authors.
# Some rights reserved, see README and LICENSE.
import csv
import json
import types
import traceback
from cStringIO import StringIO
from mimetypes import guess_type
from openpyxl import load_workbook
from os.path import abspath
from os.path import basename
from os.path import splitext
from re import subn
from xlrd import open_workbook
# from senaite.instruments.instrument import xls_to_csv
# from senaite.instruments.instrument import xlsx_to_csv

from senaite.core.exportimport.instruments import (
    IInstrumentAutoImportInterface, IInstrumentImportInterface
)
from senaite.core.exportimport.instruments.resultsimport import (
    AnalysisResultsImporter)
from senaite.core.exportimport.instruments.resultsimport import (
    InstrumentResultsFileParser)

from bika.lims import api
from bika.lims import bikaMessageFactory as _
from bika.lims.catalog import CATALOG_ANALYSIS_REQUEST_LISTING
from senaite.instruments.instrument import FileStub
from senaite.instruments.instrument import SheetNotFound
from zope.interface import implements
from zope.publisher.browser import FileUpload
from Products.CMFPlone.utils import safe_unicode

field_interim_map = {
    "Formula": "formula",
    "Concentration": "concentration",
    "Z": "z",
    "Status": "status",
    "Line 1": "line_1",
    "Net int.": "net_int",
    "LLD": "lld",
    "Stat. error": "stat_error",
    "Analyzed layer": "analyzed_layer",
    "Bound %": "bound_pct",
}


class SampleNotFound(Exception):
    pass


class MultipleAnalysesFound(Exception):
    pass


class AnalysisNotFound(Exception):
    pass


class FlameAtomicParser(InstrumentResultsFileParser):
    ar = None

    def __init__(self, infile, worksheet=None, encoding=None, delimiter=None):
        self.delimiter = delimiter if delimiter else ","
        self.encoding = encoding
        self.ar = None
        self.analyses = None
        self.worksheet = worksheet if worksheet else 0
        self.infile = infile
        self.csv_data = None
        self.sample_id = None
        mimetype, encoding = guess_type(self.infile.filename)
        InstrumentResultsFileParser.__init__(self, infile, mimetype)

    def xls_to_csv(self, infile, worksheet=0, delimiter=","):
        """
        Convert xlsx to easier format first, since we want to use the
        convenience of the CSV library
        """

        def find_sheet(wb, worksheet):
            for sheet in wb.sheets():
                if sheet.name == worksheet:
                    return sheet

        wb = open_workbook(file_contents=infile.read())
        sheet = wb.sheets()[worksheet]

        buffer = StringIO()

        # extract all rows
        for row in sheet.get_rows():
            line = []
            for cell in row:
                value = cell.value
                if type(value) in types.StringTypes:
                    value = value.encode("utf8")
                if value is None:
                    value = ""
                line.append(str(value))
            print >> buffer, delimiter.join(line)
        buffer.seek(0)
        return buffer

    def xlsx_to_csv(self, infile, worksheet=None, delimiter=","):
        worksheet = worksheet if worksheet else 0
        wb = load_workbook(filename=infile)
        if worksheet in wb.sheetnames:
            sheet = wb[worksheet]
        else:
            try:
                index = int(worksheet)
                sheet = wb.worksheets[index]
            except (ValueError, TypeError, IndexError):
                raise SheetNotFound

        buffer = StringIO()
        # import pdb; pdb.set_trace()
        for row in sheet.rows:
            line = []
            for cell in row:
                new_val = ''
                if cell.number_format == "0.00%":
                    new_val = '{}%'.format(cell.value * 100)
                # if ('\x') in cell.value.encode('utf-8'):
                #     cellval = cell.value
                    # cell.value = cell.value.replace('u\xe9','u\e')
                    # cell.value = cell.value.replace('u\xe9','u\e')
                
                cellval = new_val if new_val else cell.value
                if (isinstance(cellval, (int, long, float))):
                    value = "" if cellval is None else str(cellval).encode("utf8")
                else:
                    # value = "" if cellval is None else str(cellval).encode("utf8")
                    value = "" if cellval is None else cellval.encode("utf8")
                if "\n" in value:  # fixme multi-line cell gives only first line
                    value = value.split("\n")[0]
                line.append(value.strip())
            if not any(line):
                continue
            buffer.write(delimiter.join(line) + "\n")
        buffer.seek(0)
        # import pdb; pdb.set_trace()
        return buffer


    def parse(self):
        # import pdb; pdb.set_trace()
        order = []
        ext = splitext(self.infile.filename.lower())[-1]
        if ext == ".xlsx":
            order = (self.xlsx_to_csv, self.xls_to_csv)
        elif ext == ".xls":
            order = (self.xls_to_csv, self.xlsx_to_csv)
        elif ext == ".csv":
            self.csv_data = self.infile
        if order:
            for importer in order:
                try:
                    self.csv_data = importer(
                        infile=self.infile,
                        worksheet=self.worksheet,
                        delimiter=self.delimiter,
                    )
                    break
                except SheetNotFound:
                    self.err("Sheet not found in workbook: %s" % self.worksheet)
                    return -1
                except Exception as e:  # noqa
                    pass
            else:
                self.warn("Can't parse input file as XLS, XLSX, or CSV.")
                return -1
        stub = FileStub(file=self.csv_data, name=str(self.infile.filename))
        self.csv_data = FileUpload(stub)

        # try:
        #     #MESS AROUND HERE
        #     import pdb; pdb.set_trace()
        #     sample_id, ext = splitext(basename(self.infile.filename))
        #     # maybe the filename is a sample ID, just the way it is
        #     ar = self.get_ar(sample_id)
        #     if not ar:
        #         # maybe we need to chop of it's -9digit suffix
        #         sample_id = "-".join(sample_id.split("-")[:-1])
        #         ar = self.get_ar(sample_id)
        #         if not ar:
        #             # or we are out of luck
        #             msg = "Can't find sample for " + self.infile.filename
        #             self.warn(msg)
        #             return -1
        #     self.ar = ar
        #     self.sample_id = sample_id
        #     self.analyses = self.get_analyses(ar)
        # except Exception as e:
        #     self.err(repr(e))
        #     return False
        lines = self.csv_data.readlines()
        #  POSSIBKY ITERATE THROUGH LINES INSTEAD OF READER
        # reader = csv.DictReader(lines)
        sample_service = 'Gold'#lines[4][1]
        round = 0
        # split_row = lines.split(";")
        for row_nr, row in enumerate(lines):
            # import pdb; pdb.set_trace()
            if 'M\xc3\xa9thode: Au Aqua Regia Echelle' in row.split(",")[0]:
                round = round + 1
            self.parse_row(row_nr, row.split(","),sample_service,round)
        # import pdb; pdb.set_trace()
        return 0
    # def get_sample(row,)

    def parse_row(self, row_nr, row,sample_service,round):#(self, ar, row_nr, row)
        # import pdb; pdb.set_trace()
        parsed = {}
        if row[0] == 'Analyste\xc3\xa9':
            return
        # if row_nr == 4:
        if row[0] == 'M\xc3\xa9thodes':
            if row[1] == 'Au':
                sample_service = 'Gold' #This is not stored for all indexes of the list.
        if 'M\xc3\xa9thode: Au Aqua Regia Echelle' in row[0]:#Introducing a new round - no entries on this line
            return
        if row_nr > 5:
            sample = self.get_ar(row[0])
            analyses = sample.getAnalyses()
            for analysis in analyses:
                if sample_service == analysis.Title:
                    keyword = analysis.getKeyword
                    if row[1] == 'Over':#Will be catered for in the next sample batch
                        return
                    parsed["Reading"] = row[1] #row[2] and row[3] for second and third rounds
                    parsed["Factor"] = row[8] #Factor part of equation
                    parsed["round"] = round #In which round these were handled
                    parsed.update({"DefaultResult": "Reading"})
                    self._addRawResult(sample.id, {keyword: parsed})



            #then anlayis
            #Then iterate
            # do stuff

        # import pdb; pdb.set_trace()
        # convert row to use interim field names
        #MORE LOGIC HERE
        # import pdb; pdb.set_trace()
        # parsed = {field_interim_map.get(k, ""): v for k, v in row.items()}

        #AR TO BE FOUND FROM ROW
        return 
        formula = parsed.get("formula")
        kw = subn(r'[^\w\d\-_]*', '', formula)[0]
        try:
            analysis = self.get_analysis(ar, kw)#"ar remove"
            if not analysis:
                return 0
            keyword = analysis.getKeyword
        except Exception as e:
            self.warn(
                msg="Error getting analysis for '${kw}': ${e}",
                mapping={"kw": kw, "e": repr(e)},
                numline=row_nr,
                line=str(row),
            )
            return

        # Concentration can be PPM or PCT as it likes, I'll save both.
        concentration = parsed["concentration"]
        try:
            val = float(subn(r'[^.\d]', '', str(concentration))[0])
        except (TypeError, ValueError, IndexError):
            self.warn(
                msg="Can't extract numerical value from `concentration`",
                numline=row_nr,
                line=str(row),
            )
            parsed["reading"] = ""
            return 0
        else:
            if "ppm" in concentration.lower():
                parsed["reading"] = val * 0.0001
            elif "%" in concentration:
                parsed["reading"] = val
            else:
                self.warn(
                    "Can't decide if reading units are PPM or %",
                    numline=row_nr,
                    line=str(row),
                )
                return 0

        parsed.update({"DefaultResult": "reading"})

        self._addRawResult(self.sample_id, {keyword: parsed})
        return 0

    @staticmethod
    def get_ar(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, CATALOG_ANALYSIS_REQUEST_LISTING)
        try:
            return api.get_object(brains[0])
        except IndexError:
            pass

    @staticmethod
    def get_analyses(ar):
        analyses = ar.getAnalyses()
        return dict((a.getKeyword, a) for a in analyses)

    def get_analysis(self, ar, kw):
        analyses = self.get_analyses(ar)
        analyses = [v for k, v in analyses.items() if k.startswith(kw)]
        if len(analyses) < 1:
            self.log('No analysis found matching keyword "${kw}"', mapping=dict(kw=kw))
            return None
        if len(analyses) > 1:
            self.warn(
                'Multiple analyses found matching Keyword "${kw}"', mapping=dict(kw=kw)
            )
            return None
        return analyses[0]


class flameatomicimport(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "Agilent Flame Atomic Absorption"
    __file__ = abspath(__file__)  # noqa

    def __init__(self, context):
        self.context = context
        self.request = None

    @staticmethod
    def Import(context, request):
        errors = []
        logs = []
        warns = []

        infile = request.form["instrument_results_file"]
        if not hasattr(infile, "filename"):
            errors.append(_("No file selected"))

        artoapply = request.form["artoapply"]
        override = request.form["results_override"]
        instrument = request.form.get("instrument", None)
        worksheet = request.form.get("worksheet", 0)
        # import pdb; pdb.set_trace()
        parser = FlameAtomicParser(infile, worksheet=worksheet)
        if parser:

            status = ["sample_received", "attachment_due", "to_be_verified"]
            if artoapply == "received":
                status = ["sample_received"]
            elif artoapply == "received_tobeverified":
                status = ["sample_received", "attachment_due", "to_be_verified"]

            over = [False, False]
            if override == "nooverride":
                over = [False, False]
            elif override == "override":
                over = [True, False]
            elif override == "overrideempty":
                over = [True, True]

            importer = AnalysisResultsImporter(
                parser=parser,
                context=context,
                allowed_ar_states=status,
                allowed_analysis_states=None,
                override=over,
                instrument_uid=instrument,
            )

            try:
                importer.process()
                errors = importer.errors
                logs = importer.logs
                warns = importer.warns
            except Exception as e:
                errors.extend([repr(e), traceback.format_exc()])

        results = {"errors": errors, "log": logs, "warns": warns}

        return json.dumps(results)