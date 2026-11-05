import copy
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import unittest

import openpyxl
from openpyxl.styles import PatternFill

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_MAIN = '(auto) SEO Checklist'
sys.path.insert(0, str(ROOT / 'scripts'))


class ConsolidationTests(unittest.TestCase):
    def setUp(self):
        (ROOT / '.runs').mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=ROOT / '.runs')
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.template = self.folder / 'template.xlsx'
        w = openpyxl.Workbook()
        s = w.active
        s.title = 'SEO Implementation Checklist'
        s.append(['', '#', 'ITEM', 'Initial Check', 'Findings', 'Instruction', 'Links'])
        for ident in ['7.1', '9.1', '9.2', '9.3'] + [f'11.{i}' for i in range(1, 10)] + [f'17.{i}' for i in range(1, 5)]:
            s.append(['', ident, ident + ' question'])
        s['D2'] = 'original GSC'
        s['J2'] = '=1+2'
        s['J2'].fill = PatternFill('solid', fgColor='123456')
        for title, headers in [('11. Sitemap', ['Sitemaps', 'URLs Example', 'Instructions', 'Screenshots (If Applicable)']), ('9. Implement HTTPS Protocol', ['Source','Destination','Anchor Text','Issues','Screenshots (If have)','Specific Instructions (If have)'])] + [(title, ['Address','Titles','New Title','Instructions']) for title in ['17.1 Page titles - Missing','17.2 Page titles - Duplicate','17.3 Page Titles - Too Long','17.4 Page Titles - Multiple']]:
            s = w.create_sheet(title)
            s['A1'] = 'Implementation Checklist'
            s['A1'].hyperlink = "#'SEO Implementation Checklist'!A1"
            s.append(['Priorities & Status'])
            s.append(headers)
        w['11. Sitemap']['C4'] = 'preserve template note'
        w.save(self.template)
        self.manifest = {'run_id':'test-run', 'run_folder':'.', 'company':'Example Co', 'site_hosts':['www.example.com'], 'output_date':'2026-10-15', 'template':'template.xlsx', 'sources':[]}
        self.path = self.folder / 'run.json'

    def report(self, topic='sitemap_audit', ids=None, detail=None, statuses=None, split=False, legacy=False):
        ids = ids or [f'11.{i}' for i in range(1,10)]
        name = ('Example Co_' if legacy else f'{ids[0]}-{ids[-1]}_Example Co_') + topic + '_2026-10-13.xlsx'
        p = self.folder / name
        w = openpyxl.Workbook()
        s = w.active
        s.title = 'Checklist'
        s.append(['Item No.','Item Name','Initial Check','Findings','Instruction'])
        for ident in ids:
            s.append([ident,ident+' question',(statuses or {}).get(ident,'√'),'Source observation','Source instruction'])
        if topic == 'sitemap_audit':
            s = w.create_sheet('11. Sitemap')
            s.append(['Sitemaps','URLs Example','Instructions','Screenshots (If Applicable)'])
        else:
            if split:
                for ident in ids:
                    s=w.create_sheet(ident+' detail')
                    s.append(['Address','Title 1','New Title','Instruction'])
                    for row in detail or []:
                        if row[0]==ident:s.append(row[1:])
            else:
                s=w.create_sheet('17. Page Title')
                s.append(['Item No.','Address','Title 1','New Title','Instruction'])
                for row in detail or []:s.append(row)
        w.save(p)
        self.manifest['sources'].append({'file':name,'topic':topic,'date':'2026-10-13','legacy_name':legacy})
        return p

    def save(self):
        self.path.write_text(json.dumps(self.manifest),encoding='utf-8')
        return self.path

    def read(self):
        from inputs import read_run
        return read_run(self.save())

    def build(self):
        from workbook import build
        return build(self.save())

    def test_human_status_and_original_findings_survive(self):
        self.report(statuses={'11.3':'Human check'})
        r=self.read()
        self.assertEqual(r['checks']['11.3']['status'],'Human check')
        self.assertEqual(r['checks']['11.3']['raw_status'],'Human check')
        self.assertEqual(r['checks']['11.3']['findings'],'Source observation')

    def test_unknown_status_stops(self):
        self.report(statuses={'11.3':'maybe'})
        with self.assertRaisesRegex(ValueError,'status'):self.read()

    def test_legacy_filename_requires_explicit_permission(self):
        self.report(legacy=True)
        self.manifest['sources'][0]['legacy_name']=False
        with self.assertRaisesRegex(ValueError,'filename'):self.read()

    def test_wrong_filename_company_stops(self):
        self.report()
        self.manifest['company']='Other Co'
        with self.assertRaisesRegex(ValueError,'company'):self.read()

    def test_date_mismatch_stops(self):
        self.report()
        self.manifest['sources'][0]['date']='2026-10-14'
        with self.assertRaisesRegex(ValueError,'date'):self.read()

    def test_duplicate_candidates_stop(self):
        self.report()
        self.manifest['sources'].append(copy.deepcopy(self.manifest['sources'][0]))
        with self.assertRaisesRegex(ValueError,'duplicate'):self.read()

    def test_missing_check_in_full_report_stops(self):
        self.report(ids=['11.1'])
        with self.assertRaisesRegex(ValueError,'coverage|missing'):self.read()

    def test_explicit_subset_report_is_supported(self):
        self.report(ids=['11.1'])
        self.manifest['sources'][0]['checks']=['11.1']
        self.assertEqual(list(self.read()['checks']),['11.1'])

    def test_unknown_headers_stop(self):
        p=self.report()
        w=openpyxl.load_workbook(p);w['Checklist']['C1']='Unrecognized';w.save(p)
        with self.assertRaisesRegex(ValueError,'column|header'):self.read()

    def test_wrong_site_in_detail_stops(self):
        self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://other.example/title','old','new','fix']], {'17.2':'X'})
        with self.assertRaisesRegex(ValueError,'host'):self.read()

    def test_gsc_explicitly_selected_is_rejected(self):
        self.manifest['sources']=[{'file':'gsc.xlsx','topic':'gsc','date':'2026-10-13'}]
        with self.assertRaisesRegex(ValueError,'excluded|GSC|topic'):self.read()

    def test_combined_details_route_to_exact_subitem(self):
        self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        r=self.read()
        self.assertEqual(len(r['checks']['17.2']['details']),1)
        self.assertEqual(len(r['checks']['17.1']['details']),0)

    def test_split_details_work_with_reordered_columns(self):
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'},split=True)
        w=openpyxl.load_workbook(p);s=w['17.2 detail'];s['E1']='Extra evidence';s['E2']='keep me';w.save(p)
        self.assertEqual(self.read()['checks']['17.2']['details'][0]['values']['Extra evidence'],'keep me')

    def test_source_count_disagreement_stops(self):
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        w=openpyxl.load_workbook(p);w['Checklist']['D3']='2 issue row(s) found.';w.save(p)
        with self.assertRaisesRegex(ValueError,'count'):self.read()

    def test_pass_with_problem_details_stops(self):
        self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']])
        with self.assertRaisesRegex(ValueError,'result|status'):self.read()

    def test_failed_result_without_any_detail_stops(self):
        self.report(statuses={'11.3':'X'})
        with self.assertRaisesRegex(ValueError,'detail'):self.read()

    def test_unlisted_files_are_not_read(self):
        self.report()
        (self.folder/'unrelated.xlsx').write_bytes(b'invalid and not selected')
        self.assertEqual(len(self.read()['checks']),9)

    def test_path_escape_stops(self):
        self.report();self.manifest['sources'][0]['file']='../outside.xlsx'
        with self.assertRaisesRegex(ValueError,'outside|relative'):self.read()

    def test_build_preserves_template_and_gsc(self):
        self.report(statuses={'11.3':'Human check'})
        p=self.build();w=openpyxl.load_workbook(p)
        self.assertEqual(p.name,'Autocheck_onsite_audit_Example Co_2026-10-15.xlsx')
        s=w[OUTPUT_MAIN]
        self.assertEqual(s['D2'].value,'original GSC')
        self.assertEqual(s['J2'].value,'=1+2')
        self.assertEqual(s['J2'].fill.fgColor.rgb,'00123456')
        self.assertEqual(s['D3'].value,'-')
        self.assertIn('Not checked',s['E3'].value)
        self.assertEqual(w['11. Sitemap']['C4'].value,'preserve template note')
        self.assertEqual(openpyxl.load_workbook(self.template)['SEO Implementation Checklist']['D6'].value,None)

    def test_count_and_hyperlink_and_source_traceability(self):
        self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        p=self.build();w=openpyxl.load_workbook(p)
        s=w['17.2 Page titles - Duplicate']
        self.assertEqual(s['A4'].value,'https://www.example.com/a')
        self.assertEqual(s['A4'].hyperlink.target,'https://www.example.com/a')
        report=json.loads(p.with_suffix('.validation.json').read_text(encoding='utf-8'))
        self.assertEqual(report['checks']['17.2']['details'][0]['source_check'],'17.2')
        self.assertIn('page_titles_audit',report['checks']['17.2']['source_file'])
        main=w[OUTPUT_MAIN]
        r=next(row[0].row for row in main if row[1].value=='17.2')
        self.assertEqual(main.cell(r,4).value,'✖')
        self.assertIn('1',main.cell(r,5).value)
        self.assertIn('17.2 Page titles',main.cell(r,7).hyperlink.location)

    def test_no_output_overwrite(self):
        self.report();p=self.build();original=p.read_bytes()
        with self.assertRaisesRegex(ValueError,'exists'):self.build()
        self.assertEqual(original,p.read_bytes())

    def test_rebuild_stale_return_and_formula_navigation(self):
        self.report()
        w=openpyxl.load_workbook(self.template)
        w['11. Sitemap']['A1'].hyperlink='https://old.example/checklist'
        w['SEO Implementation Checklist']['G6']='=HYPERLINK("#old-anchor","Details")'
        w.save(self.template)
        p=self.build();w=openpyxl.load_workbook(p)
        main=w[OUTPUT_MAIN];detail=w['11. Sitemap']
        self.assertIsNone(detail['A1'].hyperlink.target)
        self.assertEqual(detail['A1'].hyperlink.location,f"'{OUTPUT_MAIN}'!B6")
        self.assertEqual(main['G6'].data_type,'s')
        self.assertEqual(main['G6'].value,'11. Sitemap')
        self.assertIn("'11. Sitemap'!A",main['G6'].hyperlink.location)

    def test_navigation_uses_actual_detail_or_existing_header_without_new_blocks(self):
        self.report()
        p=self.build();w=openpyxl.load_workbook(p)
        main=w[OUTPUT_MAIN];detail=w['11. Sitemap']
        from workbook import internal_target
        for row in main:
            if str(row[1].value).startswith('11.'):
                sheet,cell=internal_target(row[6].hyperlink.location)
                self.assertEqual(sheet,detail.title)
                self.assertEqual(cell,'A3')
                self.assertEqual(detail[cell].value,'Sitemaps')

    def test_detail_output_only_fills_existing_schema(self):
        self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        p=self.build();w=openpyxl.load_workbook(p);s=w['17.2 Page titles - Duplicate']
        self.assertEqual([c.value for c in s[3]],['Address','Titles','New Title','Instructions'])
        self.assertEqual(s.max_column,4)
        self.assertEqual(s.max_row,4)
        r=next(row[1].row for row in w[OUTPUT_MAIN] if row[1].value=='17.2')
        self.assertEqual(w[OUTPUT_MAIN].cell(r,7).hyperlink.location,"'17.2 Page titles - Duplicate'!A4")

    def test_navigation_validation_rejects_wrong_existing_target(self):
        from workbook import validate_navigation,navigation
        w=openpyxl.Workbook();s=w.active;s['A1']='link';s['B1']='correct';s['C1']='wrong'
        navigation(s['A1'],s.title,'C1')
        with self.assertRaisesRegex(ValueError,'target mismatch'):
            validate_navigation(w,[[s.title,'A1',s.title,'B1']])

    def test_unmappable_detail_row_stops_instead_of_adding_columns(self):
        self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        w=openpyxl.load_workbook(self.template);s=w['17.2 Page titles - Duplicate']
        for c in s[3]:c.value='Unmapped '+c.coordinate
        w.save(self.template)
        with self.assertRaisesRegex(ValueError,'no source detail fields map'):
            self.build()

    def test_output_tampering_detected(self):
        from workbook import verify
        self.report();p=self.build()
        w=openpyxl.load_workbook(p);w[OUTPUT_MAIN]['D6']='✖';w.save(p)
        with self.assertRaisesRegex(ValueError,'hash|changed'):verify(p)

    def test_asset_host_does_not_authorize_foreign_page_urls(self):
        self.manifest['asset_hosts']=['other.example']
        self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://other.example/title','old','new','fix']], {'17.2':'X'})
        with self.assertRaisesRegex(ValueError,'host'):self.read()

    def test_split_sheet_and_internal_id_conflict_stops(self):
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'},split=True)
        w=openpyxl.load_workbook(p);s=w['17.2 detail'];s['F1']='Item No.';s['F2']='17.1';w.save(p)
        with self.assertRaisesRegex(ValueError,'mismatch'):self.read()

    def test_template_formula_in_data_area_is_not_overwritten(self):
        self.report()
        w=openpyxl.load_workbook(self.template);w['SEO Implementation Checklist']['E6']='=1+1';w.save(self.template)
        with self.assertRaisesRegex(ValueError,'formula'):self.build()

    def test_template_hyperlink_in_data_area_is_not_overwritten(self):
        self.report()
        w=openpyxl.load_workbook(self.template);s=w['SEO Implementation Checklist'];s['E6']='Keep link';s['E6'].hyperlink='https://www.example.com/';w.save(self.template)
        with self.assertRaisesRegex(ValueError,'hyperlink'):self.build()

    def test_source_formula_is_not_executed(self):
        p=self.report()
        w=openpyxl.load_workbook(p);w['Checklist']['D2']='=HYPERLINK("https://other.example", "run")';w.save(p)
        with self.assertRaisesRegex(ValueError,'formula'):self.read()

    def test_longer_topic_alias_is_not_misread_as_company(self):
        p=self.report()
        renamed=p.with_name(p.name.replace('_sitemap_audit_', '_onsite_sitemap_audit_'))
        p.rename(renamed)
        self.manifest['sources'][0]['file']=renamed.name
        profile=json.loads((ROOT/'scripts/profiles.json').read_text())
        profile['topics']['sitemap_audit']['aliases'].append('onsite_sitemap_audit')
        (self.folder/'profile.json').write_text(json.dumps(profile))
        self.manifest['profile']='profile.json'
        self.assertEqual(len(self.read()['checks']),9)

    def test_populated_detail_cell_without_header_stops(self):
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        w=openpyxl.load_workbook(p);w['17. Page Title']['G2']='material evidence';w.save(p)
        with self.assertRaisesRegex(ValueError,'header'):self.read()

    def test_plain_plural_source_count_disagreement_stops(self):
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        w=openpyxl.load_workbook(p);w['Checklist']['D3']='2 issue rows found.';w.save(p)
        with self.assertRaisesRegex(ValueError,'count'):self.read()

    def test_merged_source_detail_requires_adapter(self):
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        w=openpyxl.load_workbook(p);w['17. Page Title'].merge_cells('C2:D2');w.save(p)
        with self.assertRaisesRegex(ValueError,'merged'):self.read()

    def test_merged_source_checklist_requires_adapter(self):
        p=self.report()
        w=openpyxl.load_workbook(p);w['Checklist'].merge_cells('D2:E2');w.save(p)
        with self.assertRaisesRegex(ValueError,'merged'):self.read()

    def test_incomplete_evidence_keeps_explicit_failed_result(self):
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        w=openpyxl.load_workbook(p);w['Checklist']['D3']='1 issue row(s) found. Evidence incomplete.';w.save(p)
        r=self.read()
        self.assertEqual(r['checks']['17.2']['status'],'✖')
        self.assertIn('Evidence incomplete',r['checks']['17.2']['findings'])

    def test_declared_asset_destination_is_clickable(self):
        self.manifest['asset_hosts']=['cdn.example.com']
        w=openpyxl.load_workbook(self.template);w['17.2 Page titles - Duplicate']['E3']='Destination';w.save(self.template)
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        w=openpyxl.load_workbook(p);s=w['17. Page Title'];s['F1']='Destination';s['F2']='https://cdn.example.com/file.png';w.save(p)
        out=self.build();w=openpyxl.load_workbook(out)
        cells=[c for row in w['17.2 Page titles - Duplicate'] for c in row if c.value=='https://cdn.example.com/file.png']
        self.assertEqual(len(cells),1)
        self.assertEqual(cells[0].hyperlink.target,'https://cdn.example.com/file.png')

    def test_unmapped_evidence_is_retained_in_sidecar_without_extra_columns(self):
        p=self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a','old','new','fix']], {'17.2':'X'})
        w=openpyxl.load_workbook(p);s=w['17. Page Title'];s['F1']='Evidence URL';s['F2']='https://evidence.example/screenshot';w.save(p)
        out=self.build();w=openpyxl.load_workbook(out)
        self.assertEqual(w['17.2 Page titles - Duplicate'].max_column,4)
        report=json.loads(out.with_suffix('.validation.json').read_text(encoding='utf-8'))
        self.assertEqual(report['checks']['17.2']['details'][0]['source_values']['Evidence URL'],'https://evidence.example/screenshot')

    def test_missing_other_template_check_marked_unchecked(self):
        from openpyxl.worksheet.datavalidation import DataValidation
        w=openpyxl.load_workbook(self.template);s=w['SEO Implementation Checklist'];s.append(['','21.1','Other onsite check'])
        row=s.max_row;dv=DataValidation(type='list',formula1='"✔,✖,NA,-"');s.add_data_validation(dv);dv.add(s.cell(row,4));w.save(self.template)
        self.report();out=self.build();s=openpyxl.load_workbook(out)[OUTPUT_MAIN]
        self.assertEqual(s.cell(row,4).value,'-')
        self.assertIn('Not checked',s.cell(row,5).value)

    def test_named_extra_checklist_field_survives(self):
        p=self.report()
        w=openpyxl.load_workbook(p);s=w['Checklist'];s['F1']='Evidence limit';s['F2']='Sample only';w.save(p)
        out=self.build();report=json.loads(out.with_suffix('.validation.json').read_text(encoding='utf-8'))
        self.assertEqual(report['checks']['11.1']['extra_checklist']['Evidence limit'],'Sample only')

    def test_missing_image_destination_column_stops(self):
        w=openpyxl.Workbook();s=w.active;s.title='Checklist'
        s.append(['Item No.','Item Name','Initial Check','Findings','Instruction'])
        s.append(['20.1','Missing alt','X','1 issue row(s) found.','fix'])
        s.append(['20.2','Oversize','√','No issue found.',''])
        s=w.create_sheet('20. Image');s.append(['Item No.','Source']);s.append(['20.1','https://www.example.com/'])
        name='20.1-20.2_Example Co_images_audit_2026-10-13.xlsx';w.save(self.folder/name)
        self.manifest['sources']=[{'file':name,'topic':'images_audit','date':'2026-10-13'}]
        with self.assertRaisesRegex(ValueError,'column|field'):self.read()

    def test_unencoded_whitespace_in_website_url_stops(self):
        self.report('page_titles_audit',[f'17.{i}' for i in range(1,5)], [['17.2','https://www.example.com/a b','old','new','fix']], {'17.2':'X'})
        with self.assertRaisesRegex(ValueError,'URL'):self.read()

    def test_human_check_validation_and_missing_status_are_distinct(self):
        from openpyxl.worksheet.datavalidation import DataValidation
        w=openpyxl.load_workbook(self.template);s=w['SEO Implementation Checklist']
        dv=DataValidation(type='list',formula1='"✔,✖,NA,-"');s.add_data_validation(dv);dv.add('D2:D18');w.save(self.template)
        self.report(statuses={'11.3':'Human Check'})
        out=self.build();w=openpyxl.load_workbook(out);s=w[OUTPUT_MAIN]
        self.assertEqual(s['D8'].value,'Human check')
        self.assertEqual(s['D3'].value,'-')
        self.assertIn('Not checked',s['E3'].value)
        self.assertIn('Issue count unconfirmed',s['E8'].value)
        human_rules=[v for v in s.data_validations.dataValidation if 'D8' in v.sqref]
        self.assertEqual(len(human_rules),1)
        self.assertIn('Human check',human_rules[0].formula1)
        gsc_rules=[v for v in s.data_validations.dataValidation if 'D2' in v.sqref]
        self.assertEqual(len(gsc_rules),1)
        self.assertNotIn('Human check',gsc_rules[0].formula1)
        report=json.loads(out.with_suffix('.validation.json').read_text(encoding='utf-8'))
        self.assertEqual(report['checks']['11.3']['status'],'Human check')
        self.assertEqual(report['checks']['11.3']['raw_status'],'Human Check')

    def test_main_rename_updates_formulas_named_ranges_and_return_links(self):
        from openpyxl.workbook.defined_name import DefinedName
        w=openpyxl.load_workbook(self.template);s=w['11. Sitemap']
        s['L2']="='SEO Implementation Checklist'!C1"
        w.defined_names.add(DefinedName('ChecklistTitle',attr_text="'SEO Implementation Checklist'!$C$1"))
        w.save(self.template)
        self.report();out=self.build();w=openpyxl.load_workbook(out)
        self.assertIn(OUTPUT_MAIN,w.sheetnames)
        self.assertNotIn('SEO Implementation Checklist',w.sheetnames)
        self.assertEqual(w['11. Sitemap']['L2'].value,"='(auto) SEO Checklist'!C1")
        self.assertEqual(w.defined_names['ChecklistTitle'].attr_text,"'(auto) SEO Checklist'!$C$1")
        self.assertIn("'(auto) SEO Checklist'!",w['11. Sitemap']['A1'].hyperlink.location)
        self.assertIn('SEO Implementation Checklist',openpyxl.load_workbook(self.template).sheetnames)

    def test_main_is_active_when_template_has_no_workbook_view(self):
        w=openpyxl.load_workbook(self.template);w.create_sheet('Summary',0);w.views=[];w.save(self.template)
        self.report();out=self.build();w=openpyxl.load_workbook(out)
        self.assertEqual(w.active.title,OUTPUT_MAIN)




class CurrentProducerCompatibilityTests(unittest.TestCase):
    """Synthetic current-producer contracts; no client data."""

    setUp = ConsolidationTests.setUp
    report = ConsolidationTests.report
    save = ConsolidationTests.save
    read = ConsolidationTests.read
    build = ConsolidationTests.build

    def sitemap(self, ident='11.5', urls='https://www.example.com/a\nhttps://www.example.com/b'):
        p = self.report(statuses={ident: 'X'})
        w = openpyxl.load_workbook(p)
        w['11. Sitemap'].append([ident + ' Finding label', urls, 'Fix source configuration', None])
        w.save(p)
        return p

    def test_current_sitemap_filled_template_and_multiline_provenance(self):
        self.sitemap()
        output = self.build()
        from workbook import verify
        self.assertEqual(verify(output)['checks'], 9)
        w = openpyxl.load_workbook(output)
        s = w['11. Sitemap']
        self.assertEqual(s['A6'].value, '11.5 Finding label')
        self.assertIsNone(s['A6'].hyperlink)
        self.assertEqual(s['B6'].value, 'https://www.example.com/a\nhttps://www.example.com/b')
        self.assertIsNone(s['B6'].hyperlink)
        self.assertTrue(s['B6'].alignment.wrap_text)
        self.assertEqual(s['C4'].value, 'preserve template note')
        sidecar = json.loads(output.with_suffix('.validation.json').read_text())
        check = sidecar['checks']['11.5']
        self.assertEqual(check['recorded_detail_rows'], 1)
        self.assertEqual(check['details'][0]['identity_fields'], ['Sitemaps'])
        self.assertEqual(len(check['details'][0]['validated_urls']['URLs Example']), 2)
        self.assertTrue(any('one native hyperlink' in x for x in sidecar['input_warnings']))
        self.assertEqual(w[OUTPUT_MAIN]['D2'].value, 'original GSC')

    def test_sitemap_63_rows_remain_63_and_clickable(self):
        p = self.report(statuses={'11.7': 'X'})
        w = openpyxl.load_workbook(p)
        for i in range(63):
            w['11. Sitemap'].append(['11.7 Finding', f'https://www.example.com/page/{i}', 'Fix', None])
        w.save(p)
        output = self.build()
        sidecar = json.loads(output.with_suffix('.validation.json').read_text())
        self.assertEqual(sidecar['checks']['11.7']['recorded_detail_rows'], 63)
        self.assertEqual(sidecar['validated_website_links'], 63)

    def test_sitemap_na_is_scoped_and_not_host_evidence(self):
        self.sitemap('11.1', 'N/A')
        r = self.read()
        self.assertEqual(r['checks']['11.1']['details'][0]['validated_urls']['URLs Example'], [])
        self.assertTrue(any('no structured URL' in x for x in r['warnings']))

    def test_sitemap_rejects_na_and_multiple_urls_for_11_7(self):
        p = self.sitemap('11.7', 'N/A')
        with self.assertRaisesRegex(ValueError, 'URL'): self.read()
        w = openpyxl.load_workbook(p)
        w['11. Sitemap']['B2'] = 'https://www.example.com/a\nhttps://www.example.com/b'
        w.save(p)
        with self.assertRaisesRegex(ValueError, 'URL count'): self.read()

    def test_sitemap_rejects_foreign_second_url_and_invalid_url(self):
        p = self.sitemap(urls='https://www.example.com/a\nhttps://other.example/b')
        with self.assertRaisesRegex(ValueError, 'host'): self.read()
        for value in ['https://www.example.com/a b', 'https://www.example.com:bad/a', 'N/A\nhttps://www.example.com/a', 'https://www.example.com/a, https://www.example.com/b']:
            with self.subTest(value=value):
                w = openpyxl.load_workbook(p); w['11. Sitemap']['B2'] = value; w.save(p)
                with self.assertRaisesRegex(ValueError, 'URL'): self.read()

    def test_sitemap_identity_conflicts_malformed_and_formula_stop(self):
        p = self.sitemap()
        w = openpyxl.load_workbook(p); s = w['11. Sitemap']; s['E1'] = 'Item No.'; s['E2'] = '11.7'; w.save(p)
        with self.assertRaisesRegex(ValueError, 'conflicting'): self.read()
        s['E2'] = '11.5'; s['A2'] = 'no id'; w.save(p)
        with self.assertRaisesRegex(ValueError, 'identifier'): self.read()
        s['A2'] = '11.5 Finding'; s['B2'] = '=1+1'; w.save(p)
        with self.assertRaisesRegex(ValueError, 'formula'): self.read()

    def content_x_only(self, statuses=None):
        p = self.report('page_titles_audit', ['17.1', '17.2', '17.3', '17.4'],
                        [['17.1', 'https://www.example.com/a', 'Title', 'New title', 'Fix']],
                        statuses or {'17.1': 'X'}, split=True)
        w = openpyxl.load_workbook(p)
        for name in ['17.2 detail', '17.3 detail', '17.4 detail']: del w[name]
        w.save(p)
        new = p.with_name(p.name.replace('17.1-17.4_', '17.1-17.4 ', 1)); p.rename(new)
        self.manifest['sources'][0]['file'] = new.name
        return new

    def test_space_name_x_only_and_human_omission(self):
        self.content_x_only({'17.1': 'X', '17.2': 'Human Check'})
        output = self.build()
        sidecar = json.loads(output.with_suffix('.validation.json').read_text())
        self.assertEqual(sidecar['checks']['17.2']['detail_table_state'], 'omitted by producer contract')
        self.assertEqual(sidecar['checks']['17.2']['raw_status'], 'Human Check')
        self.assertEqual(sidecar['checks']['17.1']['recorded_detail_rows'], 1)

    def test_all_pass_and_all_human_checklist_only(self):
        p = self.content_x_only()
        w = openpyxl.load_workbook(p); del w['17.1 detail']
        for status in ['√', 'Human Check']:
            for row in w['Checklist'].iter_rows(min_row=2): row[2].value = status
            w.save(p)
            r = self.read()
            self.assertTrue(all(c['detail_table_state']=='omitted by producer contract' for c in r['checks'].values()))
            self.assertTrue(all(c['raw_status']==status for c in r['checks'].values()))

    def test_x_detail_omission_and_empty_fail(self):
        p = self.content_x_only(); w = openpyxl.load_workbook(p)
        w['17.1 detail'].delete_rows(2); w.save(p)
        with self.assertRaisesRegex(ValueError, 'without issue'): self.read()
        del w['17.1 detail']; w.save(p)
        with self.assertRaisesRegex(ValueError, 'coverage'): self.read()

    def test_space_prefix_range_still_checked_with_legacy_permission(self):
        p = self.content_x_only()
        new = p.with_name(p.name.replace('17.1-17.4 ', '17.1-17.3 ')); p.rename(new)
        self.manifest['sources'][0].update(file=new.name, legacy_name=True)
        with self.assertRaisesRegex(ValueError, 'coverage'): self.read()

    def test_combined_split_overlap_both_orders(self):
        p = self.content_x_only(); w = openpyxl.load_workbook(p)
        s = w.create_sheet('17. Page Title'); s.append(['Item No.', 'Address'])
        for first in [True, False]:
            if first: w.move_sheet(s, offset=-1)
            else: w.move_sheet(s, offset=1)
            w.save(p)
            with self.assertRaisesRegex(ValueError, 'duplicate'): self.read()

    def test_description_exact_alias_and_duplicate_and_id_conflict(self):
        p = self.report('meta_descriptions_audit', ['18.1', '18.2', '18.3'],
                        [['18.1', 'https://www.example.com/a', 'Description', '', 'Fix']], {'18.1':'X'}, split=True)
        w = openpyxl.load_workbook(p); w['18.1 detail'].title = '18. Descriptions - Missing'
        del w['18.2 detail']; del w['18.3 detail']; w.save(p)
        r = self.read(); self.assertEqual(r['checks']['18.1']['detail_count'], 1)
        s = w.create_sheet('18.1 duplicate'); s.append(['Address']); w.save(p)
        with self.assertRaisesRegex(ValueError, 'duplicate'): self.read()
        del w[s.title]; s = w['18. Descriptions - Missing']; s['E1']='Item No.'; s['E2']='18.2'; w.save(p)
        with self.assertRaisesRegex(ValueError, 'mismatch'): self.read()

    def https_current(self, ident='9.2', status='X', address='https://www.example.com/a', related='https://cdn.example.com/a.png'):
        p = self.report('https_audit', ['9.1','9.2','9.3'], statuses={ident:status})
        w = openpyxl.load_workbook(p); del w['17. Page Title']
        s = w.create_sheet('9. Implement HTTPS Protocol')
        s.append(['Issue Type', 'Address', 'Related URL', 'Findings', 'Instructions'])
        if status != '√': s.append([ident+' Issue', address, related, 'Evidence', 'Fix'])
        w.save(p); self.manifest['asset_hosts']=['cdn.example.com']
        return p

    def test_https_resource_role_and_unmapped_values_preserved(self):
        self.https_current()
        output = self.build(); sidecar = json.loads(output.with_suffix('.validation.json').read_text())
        d = sidecar['checks']['9.2']['details'][0]
        self.assertEqual(d['source_values']['Related URL'], 'https://cdn.example.com/a.png')
        self.assertEqual(d['source_values']['Findings'], 'Evidence')
        self.assertEqual(d['validated_urls']['Related URL'], ['https://cdn.example.com/a.png'])

    def test_https_foreign_address_not_allowed_by_asset_hosts(self):
        self.https_current(address='https://cdn.example.com/a')
        with self.assertRaisesRegex(ValueError, 'host'): self.read()

    def test_https_other_checks_optional_related_and_human_details(self):
        p = self.https_current(ident='9.1', status='Human Check', related=None)
        self.assertEqual(self.read()['checks']['9.1']['status'], 'Human check')
        w = openpyxl.load_workbook(p); s=w['9. Implement HTTPS Protocol']; s['C2']='https://cdn.example.com/a'; w.save(p)
        with self.assertRaisesRegex(ValueError, 'host'): self.read()
        s['C2']='https://www.example.com/related'; w.save(p)
        self.assertEqual(self.read()['checks']['9.1']['detail_count'], 1)
        s['A2']='9.3 Issue'; w['Checklist']['C2']='√'; w['Checklist']['C4']='X'; w.save(p)
        self.assertEqual(self.read()['checks']['9.3']['detail_count'], 1)

    def test_https_header_only_and_required_headers_and_combined_status(self):
        p=self.https_current(status='√'); self.assertEqual(self.read()['checks']['9.2']['detail_count'], 0)
        w=openpyxl.load_workbook(p); w['Checklist']['C3']='X + Human Check'; w.save(p)
        with self.assertRaisesRegex(ValueError, 'status'): self.read()
        w['Checklist']['C3']='√'; w['9. Implement HTTPS Protocol']['D1']='Other'; w.save(p)
        with self.assertRaisesRegex(ValueError, 'headers'): self.read()

    def test_original_template_build_reopen_preserves_images_and_navigation(self):
        self.sitemap()
        self.manifest['template'] = str(ROOT / 'assets/Onsite Checklist - TEMPLATE.xlsx')
        original = openpyxl.load_workbook(self.manifest['template'])
        output = self.build()
        from workbook import verify, target_rows
        self.assertEqual(verify(output)['checks'], 9)
        result = openpyxl.load_workbook(output)
        self.assertEqual(result.sheetnames, [OUTPUT_MAIN if n=='SEO Implementation Checklist' else n for n in original.sheetnames])
        self.assertEqual(sum(len(s._images) for s in original), sum(len(s._images) for s in result))
        row = target_rows(result[OUTPUT_MAIN])['11.5']
        self.assertEqual(result[OUTPUT_MAIN].cell(row,4).value, '✖')
        self.assertIn('Source observation', result[OUTPUT_MAIN].cell(row,5).value)
        self.assertEqual(result[OUTPUT_MAIN].cell(row,6).value, 'Source instruction')
        self.assertTrue(result[OUTPUT_MAIN].cell(row,7).hyperlink.location.startswith("'11. Sitemap'!"))
        original.close(); result.close()


if __name__=='__main__':unittest.main()

class ConditionalTests(unittest.TestCase):
    setUp=ConsolidationTests.setUp
    save=ConsolidationTests.save
    read=ConsolidationTests.read
    build=ConsolidationTests.build

    def test_conditional_contracts(self):
        for topic,series,count,sheet in [('pagination_audit',14,6,'14. Pagination'),('hreflang_audit',15,3,'15. HREFLANG')]:
            name=f'{series}.1-{series}.{count}_Example Co_{topic}_2026-10-13.xlsx'
            self.manifest['sources']=[{'file':name,'topic':topic,'date':'2026-10-13'}]
            for mode in ['pass','human','mixed','missing','x_absent','x_empty','x_valid','x_missing_one','pass_details','header_only']:
                with self.subTest(topic=topic,mode=mode):
                    w=openpyxl.Workbook();s=w.active;s.title='Checklist';s.append(['Item No.','Item Name','Check Result','Findings','Instruction'])
                    for i in range(1,count+1):
                        if mode=='missing' and i==count:continue
                        status='√' if mode=='pass' else 'Human check'
                        if mode=='mixed':status='√' if i==4 else 'Human check'
                        if mode.startswith('x') and i==1:status='X'
                        if mode=='x_missing_one' and i==2:status='X'
                        if mode=='pass_details' and i==1:status='√'
                        s.append([f'{series}.{i}','Question',status,'Original finding','Original instruction'])
                    if mode in ['x_empty','x_valid','x_missing_one','pass_details','header_only']:
                        ds=w.create_sheet(sheet);ds.append(['Item No.','Address','Issue','Instructions'])
                        if mode in ['x_valid','x_missing_one','pass_details']:ds.append([f'{series}.1','https://www.example.com/a','Finding','Fix'])
                    w.save(self.folder/name)
                    if mode in ['missing','x_absent','x_empty','x_missing_one','pass_details']:
                        with self.assertRaises(ValueError):self.read()
                    else:
                        r=self.read();self.assertEqual(len(r['checks']),count)
                        for c in r['checks'].values():self.assertEqual(c['findings'],'Original finding')
                        self.assertEqual(r['checks'][f'{series}.1']['detail_table_state'],'observed' if mode in ['x_valid','header_only'] else 'omitted by producer contract')
                        if mode=='mixed':
                            self.manifest['template']=str(ROOT/'assets/Onsite Checklist - TEMPLATE.xlsx')
                            self.manifest['output_date']=f'2026-10-{series}'
                            from workbook import verify
                            self.assertEqual(verify(self.build())['checks'],count)
