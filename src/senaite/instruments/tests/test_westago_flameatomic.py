# -*- coding: utf-8 -*-
#
# This file is part of SENAITE.INSTRUMENTS
#
# Copyright 2018 by it's authors.





#From winlab32

import cStringIO
from datetime import datetime
from os.path import abspath
from os.path import dirname
from os.path import join

import unittest2 as unittest
from plone.app.testing import TEST_USER_ID
from plone.app.testing import TEST_USER_NAME
from plone.app.testing import login
from plone.app.testing import setRoles

from bika.lims import api
from senaite.instruments.instruments.agilent.flameatomic.flameatomic import flameatomicimport
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest

#Imports for QC samples

import re
import transaction
from bika.lims import api
from bika.lims.utils.analysisrequest import create_analysisrequest
from bika.lims.workflow import doActionFor
from DateTime import DateTime
# from plone.app.testing import TEST_USER_ID
# from plone.app.testing import TEST_USER_PASSWORD
# from plone.app.testing import setRoles




TITLE = 'Agilent Flame Atomic Absorption'
IFACE = 'senaite.instruments.instruments' \
        '.agilent.flameatomic.flameatomic.flameatomicimport'

here = abspath(dirname(__file__))
path = join(here, 'files', 'instruments', 'agilent', 'flameatomic')
fn_many_analyses = join(path, 'flameatomic_many_analyses.xlsx')
fn_second = join(path, 'flameatomic_second.xlsx')
fn_result_over = join(path,'flameatomic_OVER.xlsx' )

service_interims = [
    dict(keyword='reading', title='Reading', hidden=False)
]

calculation_interims = [
    dict(keyword='reading', title='Reading', hidden=False),
    dict(keyword='factor', title='Factor', hidden=False)
]


class TestFlameAtomic(BaseTestCase):

    def setUp(self):
        super(TestFlameAtomic, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ['Member', 'LabManager'])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(title='Happy Hills', ClientID='HH')

        self.contact = self.add_contact(
            self.client, Firstname='Rita', Surname='Mohale')

        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title='Flame Atomic Absorption'),
            Manufacturer=self.add_manufacturer(title='Agilent'),
            Supplier=self.add_supplier(title='Instruments Inc'),
            ImportDataInterface=IFACE)

        self.calculation = self.add_calculation(
            title='Dilution', Formula='[reading] * [factor]',
            InterimFields=calculation_interims)

        self.services = [
            self.add_analysisservice(
                title='Gold',
                Keyword='Au',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(title='Organic'),
                Calculation=self.calculation,
                InterimFields=service_interims),
            self.add_analysisservice(
                title='Magnesium',
                Keyword='Mg',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(title='Organic'),
                Calculation= self.calculation,
                InterimFields=service_interims)
        ]
        self.sampletype = self.add_sampletype(
            title='Dust', RetentionPeriod=dict(days=1),
            MinimumVolume='1 kg', Prefix='DU')

    def test_single_ar_with_multiple_analysis_services(self):
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        import pdb;pdb.set_trace()
        data = open(fn_many_analyses, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_many_analyses))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = flameatomicimport.Import(self.portal, request)
        a_gold = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_mag = ar.getAnalyses(full_objects=True, getKeyword='Mg')[0]
        test_results = eval(results)  # noqa
        self.assertEqual(a_gold.getResult(), '1.5')
        self.assertEqual(a_mag.getResult(), '3.0')


    def test_OVER_result(self):
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        ar_2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar_3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        import pdb;pdb.set_trace()
        data = open(fn_result_over, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_result_over))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = flameatomicimport.Import(self.portal, request)
        a_gold_first = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_second = ar_2.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_third = ar_3.getAnalyses(full_objects=True, getKeyword='Au')[0]

        test_results = eval(results)  # noqa
        self.assertEqual(a_gold_first.getResult(), '2.0')
        self.assertEqual(a_gold_second.getResult(), '1.5')
        self.assertEqual(a_gold_third.getResult(), '999999.0')

    def test_general_requirements(self):
        #1. Instrument should only accept requests in received state
        #2. Dilution factor multiplication should work fine.
        #3. 
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        ar_2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar_3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        import pdb;pdb.set_trace()
        data = open(fn_result_over, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_result_over))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = flameatomicimport.Import(self.portal, request)
        a_gold_first = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_second = ar_2.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_third = ar_3.getAnalyses(full_objects=True, getKeyword='Au')[0]

        test_results = eval(results)  # noqa
        self.assertEqual(a_gold_first.getResult(), '2.0')
        self.assertEqual(a_gold_second.getResult(), '1.5')
        self.assertEqual(a_gold_third.getResult(), '999999.0')

    def test_qc_samples(self):
        #1. The QC analyses can sometimes come in with only the sample ID QC22-019 and other times with the postfix QC22-019-005
        #2. The Sample ID read on file does not (necessarily) contain the postfix and the importer must search for QC analyses on Sample ID only for these
        #3. If more than one QC Analysis is found for the same Sample ID, the results value is not imported and the system reports “More than one Analysis found for Sample ID <...>. Not imported”
    #     pass

        #Variables
        portal = self.portal
        request = self.request
        bika_setup = portal.bika_setup
        bikasetup = portal.bika_setup
        bika_analysisservices = bika_setup.bika_analysisservices
        bika_calculations = bika_setup.bika_calculations

        # Create some basic objects for the test

        setRoles(portal, TEST_USER_ID, ['LabManager', 'Analyst'])
        date_now = DateTime().strftime("%Y-%m-%d")
        date_future = (DateTime() + 5).strftime("%Y-%m-%d")
        client = api.create(portal.clients, "Client", Name="Happy Hills", ClientID="HH", MemberDiscountApplies=True)
        contact = api.create(client, "Contact", Firstname="Rita", Lastname="Mohale")
        sampletype = api.create(bikasetup.bika_sampletypes, "SampleType", title="Water", Prefix="W")
        labcontact = api.create(bikasetup.bika_labcontacts, "LabContact", Firstname="Lab", Lastname="Manager")
        department = api.create(bikasetup.bika_departments, "Department", title="Chemistry", Manager=labcontact)
        category = api.create(bikasetup.bika_analysiscategories, "AnalysisCategory", title="Metals", Department=department)
        supplier = api.create(bikasetup.bika_suppliers, "Supplier", Name="Naralabs")

        interim_calc = api.create(bika_calculations, 'Calculation', title='Test-Total-Pest')
        pest1 = {'keyword': 'pest1', 'title': 'Pesticide 1', 'value': 12.3, 'type': 'int', 'hidden': False, 'unit': ''}
        pest2 = {'keyword': 'pest2', 'title': 'Pesticide 2', 'value': 14.89, 'type': 'int', 'hidden': False, 'unit': ''}
        pest3 = {'keyword': 'pest3', 'title': 'Pesticide 3', 'value': 16.82, 'type': 'int', 'hidden': False, 'unit': ''}
        interims = [pest1, pest2, pest3]
        interim_calc.setInterimFields(interims)
        self.assertEqual(interim_calc.getInterimFields(), interims)
        interim_calc.setFormula('((([pest1] > 0.0) or ([pest2] > .05) or ([pest3] > 10.0) ) and "FAIL" or "PASS" )')
        total_terpenes = api.create(bika_analysisservices, 'AnalysisService', title='Total Terpenes', Keyword="TotalTerpenes")
        total_terpenes.setUseDefaultCalculation(False)
        total_terpenes.setCalculation(interim_calc)
        total_terpenes.setInterimFields(interims)
        service_uids = [total_terpenes.UID()]

        # Create a Reference Definition for blank:

        blankdef = api.create(bikasetup.bika_referencedefinitions, "ReferenceDefinition", title="Blank definition", Blank=True)
        blank_refs = [{'uid': total_terpenes.UID(), 'result': '0', 'min': '0', 'max': '0'},]
        blankdef.setReferenceResults(blank_refs)

        # And for control:

        controldef = api.create(bikasetup.bika_referencedefinitions, "ReferenceDefinition", title="Control definition")
        control_refs = [{'uid': total_terpenes.UID(), 'result': '10', 'min': '9.99', 'max': '10.01'},]
        controldef.setReferenceResults(control_refs)

        blank = api.create(supplier, "ReferenceSample", title="Blank",
            ReferenceDefinition=blankdef,
            Blank=True, ExpiryDate=date_future,
            ReferenceResults=blank_refs)
        control = api.create(supplier, "ReferenceSample", title="Control",
            ReferenceDefinition=controldef,
            Blank=False, ExpiryDate=date_future,
            ReferenceResults=control_refs)
        
        # Create an Analysis Request:

        sampletype_uid = api.get_uid(sampletype)
        values = {
            'Client': api.get_uid(client),
            'Contact': api.get_uid(contact),
            'DateSampled': date_now,
            'SampleType': sampletype_uid,
            'Priority': '1',
        }

        ar = create_analysisrequest(client, request, values, service_uids)
        success = doActionFor(ar, 'receive')
        

    # Create a new Worksheet and add the analyses:

        worksheet = api.create(portal.worksheets, "Worksheet", Analyst='test_user_1_')

        analyses = map(api.get_object, ar.getAnalyses())
        analysis = analyses[0]

        worksheet.addAnalysis(analysis)

    # analysis.getWorksheet().UID() == worksheet.UID()

        # Add a blank and a control:

        blanks = worksheet.addReferenceAnalyses(blank, service_uids)
        transaction.commit()
        blanks.sort(key=lambda analysis: analysis.getKeyword(), reverse=False)
        controls = worksheet.addReferenceAnalyses(control, service_uids)
        transaction.commit()
        controls.sort(key=lambda analysis: analysis.getKeyword(), reverse=False)
        transaction.commit()
        for analysis in worksheet.getAnalyses():
            if analysis.portal_type == 'ReferenceAnalysis':
                if analysis.getReferenceType() == 'b' or analysis.getReferenceType() == 'c':
                    # 3 is the number of interim fields on the analysis/calculation
                    if len(analysis.getInterimFields()) != 3:
                        self.fail("Blank or Control Analyses interim field are not correct")



def test_suite():
    suite = unittest.TestSuite()
    suite.addTest(unittest.makeSuite(TestFlameAtomic))
    return suite
