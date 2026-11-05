"""Synthetic robots meta tags contract tests; no client evidence."""
import copy
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import openpyxl

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from inputs import read_run, check_id
from workbook import build, verify, evidence_label, internal_target
TOPIC='robots_meta_tags_audit'
SHEET='12. Robots meta tags'
IDS=['12.1','12.2','12.3']
DATE='2026-11-05'
DATED='Status / Findings (reviewed 2026-11-05)'
HEADERS=['Issue','URLs','How to Fix',DATED]

class RobotsMetaTests(unittest.TestCase):
    def setUp(self):
        (ROOT/'.runs').mkdir(exist_ok=True)
        self.tmp=tempfile.TemporaryDirectory(dir=ROOT/'.runs');self.addCleanup(self.tmp.cleanup)
        self.folder=Path(self.tmp.name);self.path=self.folder/'run.json';self.template=self.folder/'template.xlsx'
        w=openpyxl.Workbook();s=w.active;s.title='SEO Implementation Checklist'
        s.append(['','#','ITEM','Initial Check','Findings','Instruction','Links'])
        for ident in ['7.1','9.1',*IDS]:s.append(['',12.299999999999999 if ident=='12.3' else ident,ident+' question'])
        s['D2']='GSC unchanged';s['J2']='=1+2'
        d=w.create_sheet(SHEET);d['A1']='Checklist';d['A3']='Address';d['B3']='Issues';d['C3']='Screenshots (If have)'
        w.save(self.template);w.close()
        self.m={'run_id':'synthetic-robots','run_folder':'.','company':'Example Co','site_hosts':['www.example.com'],'asset_hosts':['cdn.example.com'],'output_date':DATE,'template':'template.xlsx','sources':[]}
    def report(self,statuses=None,rows=None,headers=None,token='onsite_robots_meta_tags_audit'):
        p=self.folder/f'12.1-12.3_Example Co_{token}_{DATE}.xlsx'
        w=openpyxl.Workbook();s=w.active;s.title='Checklist';s.append(['Item No.','Item Name','Check Result','Findings','Instruction'])
        for i,ident in enumerate(IDS):s.append([ident,ident+' question',(statuses or ['✔','Human check','Human check'])[i],'Original finding 原始證據; supplied crawl only','Original instruction'])
        if rows is not None:
            d=w.create_sheet(SHEET);d.append(headers or HEADERS)
            for row in rows:d.append(row)
        w.save(p);w.close();self.m['sources']=[{'file':p.name,'topic':TOPIC,'date':DATE}];return p
    def save(self):self.path.write_text(json.dumps(self.m),encoding='utf-8');return self.path
    def read(self):return read_run(self.save())
    def edit(self,p,fn):
        w=openpyxl.load_workbook(p);fn(w);w.save(p);w.close()
    def row(self,ident='12.1',url='https://www.example.com/a'):return [ident+' Confirmed issue',url,'Exact per-issue fix','Pending implementation']
    def test_nonfailed_omitted_and_observed(self):
        for statuses in (['✔','Human check','Human Check'],['√']*3,['Human check']*3):
            for rows in (None,[]):
                with self.subTest(statuses=statuses,rows=rows):
                    self.report(statuses,rows);r=self.read()
                    for i in IDS:
                        c=r['checks'][i];self.assertEqual(c['detail_count'],0);self.assertEqual(c['source_schema'],'robots-meta-tags-conditional-v1');self.assertEqual(c['detail_table_state'],'omitted by producer contract' if rows is None else 'observed');self.assertIn('原始證據',c['findings']);self.assertEqual(c['raw_status'],statuses[IDS.index(i)])
    def test_aliases_and_manifest_identity(self):
        for token in ('onsite_robots_meta_tags_audit',TOPIC):self.report(token=token);self.assertEqual(len(self.read()['checks']),3)
        for key,value in [('date','2026-11-04'),('checks',['12.1']),('topic','unregistered')]:
            saved=copy.deepcopy(self.m);self.m['sources'][0][key]=value
            with self.assertRaises(ValueError):self.read()
            self.m=saved
        self.m['company']='Wrong Co'
        with self.assertRaises(ValueError):self.read()
        self.m['company']='Example Co';p=self.report();q=p.with_name(p.name.replace('12.1-12.3','12.1-12.2'));p.rename(q);self.m['sources'][0]['file']=q.name
        with self.assertRaises(ValueError):self.read()
    def test_checklist_rejections(self):
        mutations=[lambda s:s.delete_rows(4),lambda s:s.cell(4,1,'12.1'),lambda s:s.cell(4,1,'12.4'),lambda s:s.cell(1,2,'Wrong name')]
        mutations += [lambda s,c=c:setattr(s.cell(2,c),'value',None) for c in (2,4,5)]
        mutations += [lambda s,v=v:s.cell(2,3,v) for v in ('NA','-','Pass','maybe')]
        mutations += [lambda s:s.cell(3,3,'✖'),lambda s:s.cell(2,4,'=1+2')]
        for fn in mutations:
            with self.subTest(fn=fn):
                p=self.report();self.edit(p,lambda w:fn(w['Checklist']))
                with self.assertRaises(ValueError):self.read()
    def test_failures_require_own_evidence(self):
        for statuses,rows,valid in [(['✖','✔','✔'],[self.row()],True),(['✔','✔','✖'],[self.row('12.3')],True),(['✖','Human check','✖'],[self.row(),self.row(),self.row('12.3')],True),(['✖','✔','✔'],None,False),(['✖','✔','✔'],[],False),(['✖','✔','✖'],[self.row()],False),(['✔']*3,[self.row()],False),(['Human check']*3,[self.row()],False),(['✔']*3,[self.row('12.2')],False)]:
            with self.subTest(statuses=statuses,rows=rows):
                self.report(statuses,rows)
                if valid:self.assertEqual(sum(c['detail_count'] for c in self.read()['checks'].values()),len(rows))
                else:
                    with self.assertRaises(ValueError):self.read()
    def test_detail_values_and_urls(self):
        badrows=[]
        for col in range(4):
            row=self.row();row[col]=None;badrows.append(row)
        for issue in ('12.1','12.1: ','12.1 —','12.2 issue','12.4 issue','No ID'):
            row=self.row();row[0]=issue;badrows.append(row)
        for url in ('N/A','https://foreign.example/a','https://cdn.example.com/a','ftp://www.example.com/a','https://u:p@www.example.com/a','https://www.example.com:bad/a','https://www.example.com/a b','https://www.example.com/a\nhttps://www.example.com/b','https://www.example.com/a\n'):
            row=self.row();row[1]=url;badrows.append(row)
        row=self.row();row[3]='Implemented';badrows.append(row)
        row=self.row();row[2]='=1+2';badrows.append(row)
        for row in badrows:
            with self.subTest(row=row):
                self.report(['✖','✔','✔'],[row])
                with self.assertRaises(ValueError):self.read()
    def test_header_and_sheet_rejections_even_when_empty(self):
        for hs in [HEADERS[:-1],['Issue','URLs','How to Fix','Status / Findings (reviewed 2026-11-04)'],['Issue','URLs','How to Fix','Status / Findings (reviewed 2026-02-30)'],HEADERS+['Status / Findings (reviewed 2026-11-04)'],HEADERS+['Address'],HEADERS+['URLs'],['Issues',*HEADERS[1:]],['Issue','URLs','Instructions',DATED]]:
            with self.subTest(headers=hs):
                self.report(rows=[],headers=hs)
                with self.assertRaises(ValueError):self.read()
        for title in ('12.1 detail','Wrong sheet'):
            p=self.report(rows=[]);self.edit(p,lambda w:setattr(w[SHEET],'title',title))
            with self.assertRaises(ValueError):self.read()
        for fn in (lambda w:w[SHEET].merge_cells('A1:B1'),lambda w:w[SHEET].delete_rows(1)):
            p=self.report(rows=[]);self.edit(p,fn)
            with self.assertRaises(ValueError):self.read()
    def test_reordered_extra_fields_identity_and_unnamed(self):
        hs=[DATED,'How to Fix','URLs','Issue','Item No.','Evidence'];row=['Pending implementation','Original fix','http://www.example.com/a','12.3 issue','12.3','preserve extra']
        p=self.report(['✔','✔','✖'],[row],hs);r=self.read();self.assertEqual(r['checks']['12.3']['details'][0]['values']['Evidence'],'preserve extra')
        self.edit(p,lambda w:w[SHEET].cell(2,5,'12.1'))
        with self.assertRaises(ValueError):self.read()
        p=self.report(['✖','✔','✔'],[self.row()]);self.edit(p,lambda w:w[SHEET].cell(2,6,'unheaded'))
        with self.assertRaises(ValueError):self.read()
    def test_build_original_and_synthetic_templates(self):
        for original in (False,True):
            for failed in (False,True):
                with self.subTest(original=original,failed=failed):
                    if original:shutil.copy2(ROOT/'assets/Onsite Checklist - TEMPLATE.xlsx',self.template)
                    before=openpyxl.load_workbook(self.template);sheets=before.sheetnames;images=sum(len(s._images) for s in before)
                    self.report(['✖','Human check','✖'] if failed else None,[self.row(),self.row('12.3')] if failed else None)
                    self.m['output_date']=f'2026-11-{6+int(original)*2+int(failed):02}';out=build(self.save());verify(out)
                    w=openpyxl.load_workbook(out);side=json.loads(out.with_suffix('.validation.json').read_text());main=w['(auto) SEO Checklist']
                    self.assertEqual(w.sheetnames,['(auto) SEO Checklist' if s=='SEO Implementation Checklist' else s for s in sheets]);self.assertEqual(sum(len(s._images) for s in w),images)
                    rows={}
                    for row in main:
                        try:rows[check_id(row[1].value)]=row[1].row
                        except ValueError:pass
                    if original:self.assertEqual([rows[i] for i in IDS],[77,78,79])
                    for ident in IDS:
                        r=rows[ident];c=side['checks'][ident];self.assertEqual(main.cell(r,4).value,c['status']);self.assertEqual(main.cell(r,6).value,'Original instruction');self.assertIn('[Consolidation]',main.cell(r,5).value);self.assertIn('原始證據',main.cell(r,5).value);self.assertEqual(internal_target(main.cell(r,7).hyperlink.location),tuple(c['detail_target']))
                        if not c['details']:self.assertEqual(c['detail_target'],[SHEET,'A3'])
                        else:
                            d=c['details'][0];self.assertEqual(d['source_values']['How to Fix'],'Exact per-issue fix');self.assertEqual(d['source_values'][DATED],'Pending implementation');rr=d['output_row'];self.assertEqual(w[SHEET].cell(rr,1).value,'https://www.example.com/a');self.assertEqual(w[SHEET].cell(rr,2).value,ident+' Confirmed issue');self.assertIsNone(w[SHEET].cell(rr,3).value);self.assertEqual(w[SHEET].cell(rr,1).hyperlink.target,'https://www.example.com/a')
                    self.assertEqual(internal_target(w[SHEET]['A1'].hyperlink.location),('(auto) SEO Checklist',f'B{rows["12.1"]}'));self.assertEqual([w[SHEET].cell(3,c).value for c in range(1,4)],['Address','Issues','Screenshots (If have)']);self.assertEqual(w[SHEET].max_column,before[SHEET].max_column)
                    if failed:self.assertEqual(side['checks']['12.1']['details'][0]['output_row'],4)
                    for s in before:
                        for row in s:
                            for cell in row:
                                if cell.data_type=='f' and cell.coordinate!='A1':self.assertEqual(w['(auto) SEO Checklist' if s.title=='SEO Implementation Checklist' else s.title][cell.coordinate].value,cell.value.replace("'SEO Implementation Checklist'!", "'(auto) SEO Checklist'!"))
                    w.close();before.close()
        self.assertEqual(check_id(12.299999999999999),'12.3');self.assertEqual(check_id('12.30'),'12.30')
    def test_english_labels(self):
        for status in ('-','NA','Human check'):self.assertEqual(evidence_label({'status':status,'detail_count':2}),'Issue count unconfirmed; 2 detail row(s) supplied.')
        self.assertEqual(evidence_label({'status':'✔','detail_count':0}),'Source check passed; 0 problem detail rows recorded (not independently verified).');self.assertEqual(evidence_label({'status':'✖','detail_count':2}),'2 problem detail row(s) recorded (not a count of unique issues).')
    def test_43_checks_mixed_dates(self):
        profile=json.loads((ROOT/'scripts/profiles.json').read_text());self.assertEqual(sum(len(t['ids']) for t in profile['topics'].values()),43);shutil.copy2(ROOT/'assets/Onsite Checklist - TEMPLATE.xlsx',self.template);self.report();sources=copy.deepcopy(self.m['sources'])
        for name,t in profile['topics'].items():
            if name==TOPIC:continue
            date='2026-10-13';p=self.folder/f'{t["ids"][0]}-{t["ids"][-1]}_Example Co_{t["aliases"][0]}_{date}.xlsx'
            w=openpyxl.Workbook();s=w.active;s.title='Checklist';s.append(['Item No.','Item Name','Check Result','Findings','Instruction'])
            for ident in t['ids']:s.append([ident,ident+' original name','√','Original '+ident,'Instruction '+ident])
            if not t.get('reader',{}).get('optional_nonfailed_details'):
                d=w.create_sheet(t['source_sheets'][0]);d.append({'https_audit':['Issue Type','Address','Related URL','Findings','Instructions'],'robots_audit':['Issue','Address','How to Fix'],'sitemap_audit':['Sitemaps','URLs Example','Instructions']}[name])
            w.save(p);w.close();sources.append({'file':p.name,'topic':name,'date':date})
        self.m['sources']=sources;out=build(self.save());self.assertEqual(verify(out)['checks'],43);side=json.loads(out.with_suffix('.validation.json').read_text());self.assertEqual(len(side['checks']),43)
        for ident,c in side['checks'].items():
            if ident not in IDS:self.assertEqual(c['status'],'✔');self.assertEqual(c['recorded_detail_rows'],0);self.assertEqual(c['findings'],'Original '+ident);self.assertEqual(c['instruction'],'Instruction '+ident);self.assertEqual(c['source_date'],'2026-10-13')
        self.assertEqual(side['checks']['12.1']['source_date'],DATE)
    def test_cli_and_collision(self):
        self.report();self.save()
        def cli(*args):return subprocess.run([sys.executable,str(ROOT/'scripts/consolidate.py'),*map(str,args)],capture_output=True,text=True)
        result=cli('build','--manifest',self.path);self.assertEqual(result.returncode,0,result.stderr);out=Path(json.loads(result.stdout)['output']);self.assertEqual(cli('verify',out).returncode,0);original=out.read_bytes();self.assertNotEqual(cli('build','--manifest',self.path).returncode,0);self.assertEqual(original,out.read_bytes())
if __name__=='__main__':unittest.main()
