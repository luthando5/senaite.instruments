# -*- coding: utf-8 -*-
#
# This file is part of SENAITE.INSTRUMENTS
#
# Copyright 2018 by it's authors.


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




TITLE = 'Agilent Flame Atomic Absorption'
IFACE = 'senaite.instruments.instruments' \
        '.agilent.flameatomic.flameatomic.flameatomicimport'

here = abspath(dirname(__file__))
path = join(here, 'files', 'instruments', 'agilent', 'flameatomic')
fn_many_analyses = join(path, 'flameatomic_many_analyses.xlsx')
fn_sameID = join(path, 'flameatomic_sameID_error_test.xlsx')
fn_result_over = join(path,'flameatomic_OVER.xlsx' )
fn_QC_and_blank = join(path,'flameatomic_QC_and_blank.xlsx')

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
        self.assertEqual(a_gold.getResult(), '0.75')
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

        test_results = eval(results)
        self.assertEqual(a_gold_first.getResult(), '0.5')
        self.assertEqual(a_gold_second.getResult(), '0.375')
        self.assertEqual(a_gold_third.getResult(), '999999.0')


    def test_same_ID_appears_more_than_once(self):
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
        data = open(fn_sameID, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_sameID))
        
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

        test_results = eval(results)
        #No results will be updated if one of the ID's have not been entered correctly
        self.assertEqual(a_gold_first.getResult(), '')
        self.assertEqual(a_gold_second.getResult(), '')
        self.assertEqual(a_gold_third.getResult(), '')


    def test_of_legends(self):

        import re
        import transaction
        from bika.lims import api
        from bika.lims.utils.analysisrequest import create_analysisrequest
        from bika.lims.workflow import doActionFor
        from DateTime import DateTime
        from plone.app.testing import TEST_USER_ID
        from plone.app.testing import TEST_USER_PASSWORD
        from plone.app.testing import setRoles

    # Variables:

        portal = self.portal
        request = self.request
        bika_setup = portal.bika_setup
        bikasetup = portal.bika_setup
        bika_analysisservices = bika_setup.bika_analysisservices
        bika_calculations = bika_setup.bika_calculations

    # We need to create some basic objects for the test:

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

        interim_calc = api.create(bika_calculations, 'Calculation', title='Test-Total-Dust')
        result = {'keyword': 'result', 'title': 'Results', 'value': 12.3, 'type': 'int', 'hidden': False, 'unit': ''}
        factor = {'keyword': 'factor', 'title': 'Factor', 'value': 14.89, 'type': 'int', 'hidden': False, 'unit': ''}
        interims = [result, factor]
        interim_calc.setInterimFields(interims)
        self.assertEqual(interim_calc.getInterimFields(), interims)
        interim_calc.setFormula('[reading] * [factor]')
        total_Dust = api.create(bika_analysisservices, 'AnalysisService', title='Gold', Keyword="Au")
        total_Dust.setUseDefaultCalculation(False)
        total_Dust.setCalculation(interim_calc)
        total_Dust.setInterimFields(interims)
        service_uids = [total_Dust.UID()]

    # Create a Reference Definition for blank:

        blankdef = api.create(bikasetup.bika_referencedefinitions, "ReferenceDefinition", title="Blank definition", Blank=True)
        blank_refs = [{'uid': total_Dust.UID(), 'result': '0', 'min': '0', 'max': '0'},]
        blankdef.setReferenceResults(blank_refs)

    # And for control:

        controldef = api.create(bikasetup.bika_referencedefinitions, "ReferenceDefinition", title="Control definition")
        control_refs = [{'uid': total_Dust.UID(), 'result': '10', 'min': '9.99', 'max': '10.01'},]
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


        worksheet = api.create(portal.worksheets, "Worksheet", Analyst='test_user_1_')
        analyses = map(api.get_object, ar.getAnalyses())
        analysis = analyses[0]
        worksheet.addAnalysis(analysis)
        analysis.getWorksheet().UID() == worksheet.UID()

        data = open(fn_QC_and_blank, 'r').read()

        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_QC_and_blank))
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = flameatomicimport.Import(self.portal, request)

    # Add a blank and a control:
        import pdb;pdb.set_trace()
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
