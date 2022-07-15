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
from DateTime import DateTime
from mimetypes import guess_type
from openpyxl import load_workbook
from os.path import abspath
from os.path import splitext
from xlrd import open_workbook
import xml.etree.cElementTree as ET
from bika.lims.browser import BrowserView

from senaite.core.exportimport.instruments import (
    IInstrumentAutoImportInterface, IInstrumentImportInterface
)
from senaite.core.exportimport.instruments.resultsimport import (
    AnalysisResultsImporter)
from senaite.core.exportimport.instruments.resultsimport import (
    InstrumentResultsFileParser)
from senaite.core.exportimport.instruments import IInstrumentExportInterface

from bika.lims import api
from bika.lims import bikaMessageFactory as _
from bika.lims.catalog import CATALOG_ANALYSIS_REQUEST_LISTING
from senaite.core.catalog import ANALYSIS_CATALOG
from senaite.instruments.instrument import FileStub
from senaite.instruments.instrument import SheetNotFound
from zope.interface import implements
from zope.publisher.browser import FileUpload
from zope.component import getUtility
from plone.i18n.normalizer.interfaces import IIDNormalizer
from senaite.app.supermodel.interfaces import ISuperModel
from zope.component import getAdapter

# Unused Code
# field_interim_map = {
#     "Formula": "formula",
#     "Concentration": "concentration",
#     "Z": "z",
#     "Status": "status",
#     "Line 1": "line_1",
#     "Net int.": "net_int",
#     "LLD": "lld",
#     "Stat. error": "stat_error",
#     "Analyzed layer": "analyzed_layer",
#     "Bound %": "bound_pct",
# }


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
        self.processed_samples_class = []
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

        for row in sheet.rows:
            line = []
            for cell in row:
                new_val = ''
                if cell.number_format == "0.00%":
                    new_val = '{}%'.format(cell.value * 100)
                cellval = new_val if new_val else cell.value
                if (isinstance(cellval, (int, long, float))):
                    value = "" if cellval is None else str(cellval).encode("utf8")
                else:
                    value = "" if cellval is None else cellval.encode("utf8")
                if "\n" in value:
                    value = value.split("\n")[0]
                line.append(value.strip())
            if not any(line):
                continue
            buffer.write(delimiter.join(line) + "\n")
        buffer.seek(0)
        return buffer


    def parse(self):
        order = []
        ext = splitext(self.infile.filename.lower())[-1]
        if ext == ".xlsx":
            order = (self.xlsx_to_csv, self.xls_to_csv)
        elif ext == ".xls":
            order = (self.xls_to_csv, self.xlsx_to_csv)
        elif ext == ".csv" or ".prn":
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
                except Exception as e:
                    pass
            else:
                self.warn("Can't parse input file as XLS, XLSX, or CSV.")
                return -1
        stub = FileStub(file=self.csv_data, name=str(self.infile.filename))
        self.csv_data = FileUpload(stub)

        lines_with_parentheses = self.csv_data.readlines()
        lines = [i.replace('"','').replace('\r\n','') for i in lines_with_parentheses]

        analysis_round = 0
        sample_service = []
        for row_nr, row in enumerate(lines): #This whole part can be a function
            split_row = row.split(",")
            if 'M\xc3\xa9thode:' in split_row[0] or 'M\xe9thode:' in split_row[0]:
                analysis_round = analysis_round + 1
            if 'M\xc3\xa9thodes' in split_row[0] or 'M\xe9thodes' in split_row[0]:
                #Here we determine how many rounds there are in the sheet (Max = 3)
                if split_row[1]:
                    sample_service.append(split_row[1])
                if len(split_row) > 2 and split_row[2]:
                    sample_service.append(split_row[2])
                if len(split_row) > 3 and split_row[3]:
                    sample_service.append(split_row[3])
            if analysis_round > 0 and split_row[0] and len(split_row)>2 and split_row[1]: #How long is split_row of theres an empty cell inbetween?
                #If we are past the headerlines and the first and second columns entries (of that row) are non empty
                self.parse_row(row_nr, split_row,sample_service[analysis_round-1],analysis_round)
        return 0


    def parse_row(self, row_nr, row,sample_service,analysis_round):
        #Try to restructure parse row suc
        parsed = {}
        #Here we check whether this sample ID has been processed already
        if {row[0]:sample_service} in self.processed_samples_class:
            msg = ("Multiple results for Sample '{}' with sample service '{}' found. Not imported".format(row[0],sample_service))
            raise MultipleAnalysesFound(msg)
        if self.is_sample(row[0]):
            sample = self.get_ar(row[0])
        else:
            #Updating the Reference analyses
            sample = self.get_duplicate_or_qc(row[0],sample_service)# change to qc or reference
            if sample:# Don't have to check for sample as an error will be thrown already
                keyword = sample.getKeyword
                self.processed_samples_class.append({row[0]:sample_service})
                parsed["Reading"] = float(row[1])
                parsed["Factor"] = float(row[8])
                parsed.update({"DefaultResult": "Reading"})
                self._addRawResult(row[0], {keyword: parsed})
                return 0
            else:
                return 0
        # Updating the analysis requests
        analyses = sample.getAnalyses()
        for analysis in analyses: #Use getAnalysis instead of getAnalyses using the keyword is the distinguisher
            if sample_service == analysis.getKeyword:
                keyword = analysis.getKeyword
                if row[1] == 'OVER':
                    if analysis_round == 3:
                        #If in the third analysis_round [Reading] = OVER then the value 999999 is assigned.
                        self.processed_samples_class.append({row[0]:sample_service})
                        parsed["Reading"] = float(999999)
                        parsed["Factor"] = float(1)
                        parsed.update({"DefaultResult": "Reading"})
                        self._addRawResult(row[0], {keyword: parsed})
                    #If not in the 3rd analysis_round and Reading = OVER, we don't update the Reading
                    return
                self.processed_samples_class.append({row[0]:sample_service})
                parsed["Reading"] = float(row[1])
                parsed["Factor"] = float(row[8])
                parsed.update({"DefaultResult": "Reading"})
                self._addRawResult(row[0], {keyword: parsed})
                #Avoid repetition
        return 


    @staticmethod
    def get_ar(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, CATALOG_ANALYSIS_REQUEST_LISTING)
        try:
            return api.get_object(brains[0])
        except IndexError:
            pass
    
    @staticmethod
    def is_sample(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, CATALOG_ANALYSIS_REQUEST_LISTING)
        return True if brains else False


    @staticmethod
    def get_duplicate_or_qc(analysis_id,sample_service,):
        portal_types = ["DuplicateAnalysis", "ReferenceAnalysis"]
        query = dict(
            portal_type=portal_types, getReferenceAnalysesGroupID=analysis_id
        )
        brains = api.search(query, ANALYSIS_CATALOG)
        analyses = dict((a.getKeyword, a) for a in brains)
        brains = [v for k, v in analyses.items() if k.startswith(sample_service)]
        if len(brains) < 1:
            msg = ("No analysis found matching Keyword {}".format(sample_service))
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword {}".format(sample_service))
            raise MultipleAnalysesFound(msg)
        return brains[0]


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


class Export(BrowserView):
    # implements(IInstrumentExportInterface)
    title = "Agilent Flame Atomic Exporter"
    __file__ = abspath(__file__)  # noqa

    def __init__(self, context,request):
        self.context = context
        self.request = request

    def __call__(self, analyses):
        # import pdb;pdb.set_trace()
        # tray = 1
        now = DateTime().strftime('%Y%m%d-%H%M')
        uc = api.get_tool('uid_catalog')
        instrument = self.context.getInstrument()
        norm = getUtility(IIDNormalizer).normalize
        filename = '{}-{}.csv'.format(
            self.context.getId(), norm(self.title))
            # self.context.getId(), norm(instrument.getDataInterface()))
        # listname = '{}_{}_{}'.format(
        #     self.context.getId(),  norm(instrument.getDataInterface()))

            
        options = {
            'dilute_factor': 1,
            'method': 'Dilution',
            'notneeded1': 1,
            'notneeded2': 1,
            'notneeded3': 10
        }
        
        sample_cases = {'a':'SAMP','b':'BLANK','c':'CRM','d':'DUP'}

        # for k, v in instrument.getDataInterfaceOptions():
        #     options[k] = v

        # for looking up "cup" number (= slot) of ARs
        # parent_to_slot = {}
        
        # for x in range(len(layout)):
        #     a_uid = layout[x]['analysis_uid']
        #     p_uid = uc(UID=a_uid)[0].getObject().aq_parent.UID()
        #     layout[x]['parent_uid'] = p_uid
        #     if p_uid not in parent_to_slot.keys():
        #         parent_to_slot[p_uid] = int(layout[x]['position'])

        # write rows, one per PARENT
        # import pdb;pdb.set_trace()
        layout = self.context.getLayout()
        # header = [listname, options['method']]
        # rows = []
        # rows.append(header)
        tmprows = []
        Used_IDs = []
        test_list = []


        for item in layout: #Could use enumerate
            # create batch header row
            c_uid = item['container_uid']
            # p_uid = item['parent_uid']
            a_uid = item['analysis_uid']
            analysis = uc(UID=a_uid)[0].getObject() if a_uid else None
            keyword = str(analysis.Keyword)
            container = uc(UID=c_uid)[0].getObject() if c_uid else None
            sample_type = sample_cases[item['type']]
            
            # sample = getAdapter(item['container_uid'], ISuperModel).id


            if item['type'] == 'a':
                analysis_id = container.id
            elif (item['type'] in 'bcd'):
                analysis_id = analysis.getReferenceAnalysesGroupID()
            test_list.append([[analysis_id,sample_type,keyword]])
            if keyword == "PeseePourFusion":
                if analysis_id in Used_IDs:
                    continue
                weight = analysis.getResult()
                if not weight:
                    weight = 50
                tmprows.append(['',
                                analysis_id,
                                sample_type,
                                weight,
                                # keyword,
                                options['dilute_factor'],
                                options["notneeded1"],
                                options["notneeded2"],
                                options["notneeded3"]])
                Used_IDs.append(analysis_id)
        # import pdb;pdb.set_trace()
        # tmprows.sort(lambda a, b: cmp(a[0], b[0]))
        # rows += tmprows


        # ramdisk = StringIO()
        # writer = csv.writer(ramdisk, delimiter=',')

        # assert(writer)
        # writer.writerows(rows)
        # result = ramdisk.getvalue()
        # ramdisk.close()
        import pdb;pdb.set_trace()
        # headers = {'Content-Length': 21321,'Content-Type': 'text/comma-separated-values','Content-Disposition': 'inline; filename=%s' % filename}
        # request.get(request.getURL(),headers)
        # stream file to browser

        rows = self.row_sorter(tmprows)
        result2 = self.list_to_string(rows)
        setheader = self.request.RESPONSE.setHeader
        setheader('Content-Length', len(result2))
        setheader('Content-Disposition', 'inline; filename=%s' % filename)
        setheader('Content-Type', 'text/csv')
        # self.self.request.RESPONSE.write(result2)
        self.request.RESPONSE.write(result2)
    

    @staticmethod
    def utf8len(s):
        return len(s.encode('utf-8'))
    

    @staticmethod
    def list_to_string(rows): #maybe dict to string
        final_rows = ''
        interim_rows = []
        
        for row in rows:
            row = ','.join(str(item) for item in row)
            interim_rows.append(row)
        final_rows = '\r\n'.join(interim_rows)
        # import pdb;pdb.set_trace()
        return final_rows
            

    
    @staticmethod
    def row_sorter(rows):
        sample_cases = {'SAMP': 2,'BLANK':0,'CRM':1,'DUP':3}
        reversed_dict = {v: k for k, v in sample_cases.items()}
        for row in rows:
            row[2] = sample_cases[row[2]]
        rows.sort(lambda a, b: cmp(a[2], b[2]))
        for indx,row in enumerate(rows):
            row[0] = indx+1
            row[2] = reversed_dict[row[2]]
        return rows
        
        
            

                # self.warn("No Pesee Pour Fusion result for {}. Default of 50g allocated".format(analysis_id))
            # cup = parent_to_slot[p_uid]

    # def Export(self, context, request):
    #     tray = 1
    #     norm = getUtility(IIDNormalizer).normalize
    #     filename = '{}-{}.xml'.format(
    #         context.getId(), norm(self.title))
    #     now = str(DateTime())[:16]

    #     root = ET.Element('SequenceTableDataSet')
    #     root.set('SchemaVersion', "1.0")
    #     root.set('SequenceComment', "")
    #     root.set('SequenceOperator', "")
    #     root.set('SequenceSeqPathFileName', "")
    #     root.set('SequencePreSeqAcqCommand', "")
    #     root.set('SequencePostSeqAcqCommand', "")
    #     root.set('SequencePreSeqDACommand', "")
    #     root.set('SequencePostSeqDACommand', "")
    #     root.set('SequenceReProcessing', "False")
    #     root.set('SequenceInjectBarCodeMismatch', "OnBarcodeMismatchInjectAnyway")
    #     root.set('SequenceOverwriteExistingData', "False")
    #     root.set('SequenceModifiedTimeStamp', now)
    #     root.set('SequenceFileECMPath', "")

    #     # for looking up "cup" number (= slot) of ARs
    #     parent_to_slot = {}
    #     layout = context.getLayout()
    #     for item in layout:
    #         p_uid = item.get('parent_uid')
    #         if not p_uid:
    #             p_uid = item.get('container_uid')
    #         if p_uid not in parent_to_slot.keys():
    #             parent_to_slot[p_uid] = int(item['position'])

    #     rows = []
    #     sequences = []
    #     for item in layout:
    #         # create batch header row
    #         p_uid = item.get('parent_uid')
    #         if not p_uid:
    #             p_uid = item.get('container_uid')
    #         if not p_uid:
    #             continue
    #         if p_uid in sequences:
    #             continue
    #         sequences.append(p_uid)
    #         cup = parent_to_slot[p_uid]
    #         rows.append({
    #             'tray': tray,
    #             'cup': cup,
    #             'analysis_uid': getAdapter(item['analysis_uid'], ISuperModel),
    #             'sample': getAdapter(item['container_uid'], ISuperModel)
    #         })
    #     rows.sort(lambda a, b: cmp(a['cup'], b['cup']))

    #     cnt = 0
    #     for row in rows:
    #         seq = ET.SubElement(root, 'Sequence')
    #         ET.SubElement(seq, 'SequenceID').text = str(row['tray'])
    #         ET.SubElement(seq, 'SampleID').text = str(cnt)
    #         ET.SubElement(seq, 'AcqMethodFileName').text = 'Dunno'
    #         ET.SubElement(seq, 'AcqMethodPathName').text = 'Dunno'
    #         ET.SubElement(seq, 'DataFileName').text = row['sample'].Title()
    #         ET.SubElement(seq, 'DataPathName').text = 'Dunno'
    #         ET.SubElement(seq, 'SampleName').text = row['sample'].Title()
    #         import pdb;pdb.set_trace()
    #         if row['sample'].title == 'Test Reference Sample':
    #             ET.SubElement(seq, 'SampleType').text = "Test Reference Sample"
    #         else:
    #             ET.SubElement(seq, 'SampleType').text = row['sample'].SampleType.Title()
    #         ET.SubElement(seq, 'Vial').text = str(row['cup'])
    #         cnt += 1
    #     import pdb;pdb.set_trace()
    #     xml = ET.tostring(root, method='xml')
    #     # stream file to browser
    #     setheader = request.RESPONSE.setHeader
    #     setheader('Content-Length', len(xml.encode('utf-16')))
    #     # setheader('Content-Length', len(xml))
    #     setheader('Content-Disposition',
    #               'attachment; filename="%s"' % filename)
    #     setheader('Content-Type', 'text/xml')
    #     request.RESPONSE.write(xml)
