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
import traceback
from mimetypes import guess_type
from os.path import abspath
from os.path import splitext
from re import subn

from senaite.core.exportimport.instruments import IInstrumentAutoImportInterface
from senaite.core.exportimport.instruments import IInstrumentImportInterface
from senaite.core.exportimport.instruments.resultsimport import \
    AnalysisResultsImporter
from senaite.core.exportimport.instruments.resultsimport import \
    InstrumentResultsFileParser

from bika.lims import api
from bika.lims import bikaMessageFactory as _
from bika.lims.catalog import CATALOG_ANALYSIS_REQUEST_LISTING
from senaite.core.catalog import ANALYSIS_CATALOG, SENAITE_CATALOG
from senaite.instruments.instrument import FileStub
from senaite.instruments.instrument import SheetNotFound
from senaite.instruments.instrument import xls_to_csv
from senaite.instruments.instrument import xlsx_to_csv
from zope.interface import implements
from zope.publisher.browser import FileUpload


class MultipleAnalysesFound(Exception):
    pass


class AnalysisNotFound(Exception):
    pass


class Winlab32(InstrumentResultsFileParser):
    ar = None

    def __init__(self, infile, encoding=None, delimiter=None):
        self.delimiter = delimiter if delimiter else ','
        self.encoding = encoding
        self.infile = infile
        self.csv_data = None
        self.worksheet = 'Concentrations'
        self.sample_id = None
        self.processed_samples = []
        mimetype, encoding = guess_type(self.infile.filename)
        InstrumentResultsFileParser.__init__(self, infile, mimetype)

    def parse(self):
        order = []
        ext = splitext(self.infile.filename.lower())[-1]
        if ext == '.xlsx':
            order = (xlsx_to_csv, xls_to_csv)
        elif ext == '.xls':
            order = (xls_to_csv, xlsx_to_csv)
        elif ext == '.csv':
            self.csv_data = self.infile
        if order:
            for importer in order:
                try:
                    self.csv_data = importer(
                        infile=self.infile,
                        worksheet=self.worksheet,
                        delimiter=self.delimiter)
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

        lines = self.csv_data.readlines()
        reader = csv.DictReader(lines)
        for row in reader:
            self.parse_row(reader.line_num, row)
        return 0

    def parse_row(self, row_nr, row):
        # convert row to use interim field names
        try:
            value = float(row['Conc (Samp)'])
        except (TypeError, ValueError, KeyError):
            value = row.get('Conc (Samp)')
        # reading and Reading - found out users can have Reading or reading
        # when entering interim fields so we cater for both cases
        parsed = {'Reading': value, 'DefaultResult': 'Reading', 'reading': value}
        parsed.update(row)

        sample_id = subn(r'[^\w\d\-_]*', '', row.get('Sample ID', ""))[0]
        kw = subn(r"[^\w\d]*", "", row.get('Analyte Name', ""))[0]
        kw = kw
        if not sample_id or not kw:
            return 0

        try:
            if self.is_sample(sample_id):
                if {sample_id: kw} in self.processed_samples:
                    analysis = self.get_ar_duplicates(sample_id, kw)
                    new_kw = analysis.getKeyword
                else:
                    self.processed_samples.append({sample_id: kw})
                    ar = self.get_ar(sample_id)
                    analysis = self.get_analysis(ar, kw)
                    new_kw = analysis.getKeyword
            elif self.is_analysis_group_id(sample_id):
                analysis = self.get_duplicate_or_qc_analysis(sample_id, kw)
                new_kw = analysis.getKeyword
            else:
                sample_reference = self.get_reference_sample(sample_id, kw)
                analysis = self.get_reference_sample_analysis(sample_reference, kw)
                new_kw = analysis.getKeyword()
        except Exception as e:
            self.warn(msg="Error getting analysis for '${s}/${kw}': ${e}",
                      mapping={'s': sample_id, 'kw': kw, 'e': repr(e)},
                      numline=row_nr, line=str(row))
            return

        self._addRawResult(sample_id, {new_kw: parsed})
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
        brains = ar.getAnalyses()
        return dict((a.getKeyword, a) for a in brains)

    def get_analysis(self, ar, kw):
        kw = kw
        brains = self.get_analyses(ar)
        brains = [v for k, v in brains.items() if k.startswith(kw[:2])]
        if len(brains) < 1:
            msg = "No analysis found matching Keyword '${kw}'",
            raise AnalysisNotFound(msg, kw=kw)
        if len(brains) > 1:
            msg = "Multiple brains found matching Keyword '${kw}'",
            raise MultipleAnalysesFound(msg, kw=kw)
        return brains[0]

    @staticmethod
    def get_reference_sample_analyses(reference_sample):
        brains = reference_sample.getObject().getReferenceAnalyses()
        return dict((a.getKeyword(), a) for a in brains)

    def get_reference_sample_analysis(self, reference_sample, kw):
        kw = kw
        brains = self.get_reference_sample_analyses(reference_sample)
        brains = [v for k, v in brains.items() if k.startswith(kw[:2])]
        if len(brains) < 1:
            msg = "No analysis found matching Keyword '${kw}'",
            raise AnalysisNotFound(msg, kw=kw)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword '{}'".format(kw))
            raise MultipleAnalysesFound(msg)
        return brains[0]

    @staticmethod
    def is_sample(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, CATALOG_ANALYSIS_REQUEST_LISTING)
        return True if brains else False

    @staticmethod
    def is_analysis_group_id(analysis_group_id):
        portal_types = ["DuplicateAnalysis", "ReferenceAnalysis"]
        query = dict(
            portal_type=portal_types, getReferenceAnalysesGroupID=analysis_group_id
        )
        brains = api.search(query, ANALYSIS_CATALOG)
        return True if brains else False

    @staticmethod
    def get_duplicate_or_qc_analysis(analysis_id, kw):
        portal_types = ["DuplicateAnalysis", "ReferenceAnalysis"]
        query = dict(
            portal_type=portal_types, getReferenceAnalysesGroupID=analysis_id
        )
        brains = api.search(query, ANALYSIS_CATALOG)
        analyses = dict((a.getKeyword, a) for a in brains)
        brains = [v for k, v in analyses.items() if k.startswith(kw[:2])]
        if len(brains) < 1:
            msg = ("No analysis found matching Keyword '${kw}'",)
            raise AnalysisNotFound(msg, kw=kw)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword '${kw}'",)
            raise MultipleAnalysesFound(msg, kw=kw)
        return brains[0]

    @staticmethod
    def get_ar_duplicates(analysis_id, kw):
        query = dict(portal_type="DuplicateAnalysis")
        duplicates = api.search(query, ANALYSIS_CATALOG)
        analyses = dict((a.getKeyword, [a, a.getReferenceAnalysesGroupID]) for a in duplicates)
        brains = []
        for k, v in analyses.items():
            if k.startswith(kw[:2]) and v[1].startswith(analysis_id):
                brains.append(v[0])
        if len(brains) < 1:
            msg = ("No analysis found matching Keyword '${kw}'",)
            raise AnalysisNotFound(msg, kw=kw)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword '${kw}'",)
            raise MultipleAnalysesFound(msg, kw=kw)
        return brains[0]

    @staticmethod
    def get_reference_sample(reference_sample_id, kw):
        query = dict(
            portal_type="ReferenceSample", getId=reference_sample_id
        )
        brains = api.search(query, SENAITE_CATALOG)
        if len(brains) < 1:
            msg = ("No reference sample found matching Keyword '${kw}'",)
            raise AnalysisNotFound(msg, kw=kw)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword '{kw}'".format(kw))
            raise MultipleAnalysesFound(msg)
        return brains[0]


class importer(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "Perkin Elmer Winlab32"
    __file__ = abspath(__file__)  # noqa

    def __init__(self, context):
        self.context = context
        self.request = None

    @staticmethod
    def Import(context, request):
        errors = []
        logs = []
        warns = []

        infile = request.form['instrument_results_file']
        if not hasattr(infile, 'filename'):
            errors.append(_("No file selected"))

        artoapply = request.form['artoapply']
        override = request.form['results_override']
        instrument = request.form.get('instrument', None)

        parser = Winlab32(infile)
        if parser:

            status = ['sample_received', 'attachment_due', 'to_be_verified']
            if artoapply == 'received':
                status = ['sample_received']
            elif artoapply == 'received_tobeverified':
                status = ['sample_received', 'attachment_due', 'to_be_verified']

            over = [False, False]
            if override == 'nooverride':
                over = [False, False]
            elif override == 'override':
                over = [True, False]
            elif override == 'overrideempty':
                over = [True, True]

            importer = AnalysisResultsImporter(
                parser=parser,
                context=context,
                allowed_ar_states=status,
                allowed_analysis_states=None,
                override=over,
                instrument_uid=instrument)

            try:
                importer.process()
                errors = importer.errors
                logs = importer.logs
                warns = importer.warns
            except Exception as e:
                errors.extend([repr(e), traceback.format_exc()])

        results = {'errors': errors, 'log': logs, 'warns': warns}

        return json.dumps(results)
