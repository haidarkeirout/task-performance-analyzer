from datetime import date
from io import BytesIO
import json
from pathlib import Path
import sys
import tempfile
import unittest
from docx import Document
from openpyxl import Workbook, load_workbook
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
from company_performance.adapters import adapt_clickup_collection
from company_performance.application import build_company_analysis, build_company_preview, filter_company_preview
from company_performance.dashboard import DashboardFilters, build_company_dashboard
from company_performance.models import UnifiedStatus
from company_performance.normalization import deduplicate_tasks, normalize_status
from company_performance.outputs import write_company_excel, write_company_word_report

def jira_xlsx(space='Engineering', status='Done'):
    wb=Workbook();s=wb.active;s.title='Jira_Data';s.append(['Issue key','Issue id','Summary','Project key','Project name','Status','Priority','Assignee','Created','Due date','Custom field (Start date)']);s.append(['ENG-1','1','Ship feature','ENG',space,status,'High','Maya','2026-09-01T08:00:00Z','2026-09-10','2026-09-02']);b=BytesIO();wb.save(b);return b.getvalue()
class JiraPrepared:
    def __init__(self,space='Engineering'):
        self.xlsx=jira_xlsx(space);self.space_name=space;self.collected_at='2026-09-30T20:00:00Z';self.cutoff='2026-09-30'
        self.history_json=json.dumps({'ENG-1':{'history_complete':True,'history_through':'2026-09-30T20:00:00Z','initial_status':'To Do','status_events':[{'changed_at':'2026-09-02T08:00:00Z','from_status':'To Do','to_status':'In Progress'},{'changed_at':'2026-09-03T08:00:00Z','from_status':'In Progress','to_status':'Done'}]}}).encode()
class ClickPrepared:
    def __init__(self,space='Growth', task_id='cu-1', status=None, priority='2', name='Review campaign'):
        self.space_name=space;self.collected_at='2026-09-30T20:00:00Z';self.cutoff='2026-09-30'
        self.history_json=json.dumps({'tasks':[{'id':task_id,'name':name,'status':status or {'status':'Review'},'priority':{'priority':priority},'assignees':[{'username':'Zaher'}],'date_created':'1788256800000','due_date':'1790683200000','list':{'name':'Marketing'}}],'time_in_status':{}}).encode()

class CompanySourceSelectionTests(unittest.TestCase):
    def test_company_ui_uses_independent_jira_clickup_space_selectors_without_radio(self):
        text=(ROOT/'src/company_performance/ui.py').read_text(encoding='utf-8')
        self.assertIn('"Select Jira Spaces"',text)
        self.assertIn('"Select ClickUp Spaces"',text)
        self.assertNotIn('st.radio(',text)
        order=[text.index('"Unified Project Name"'),text.index('"Analysis Period Start"'),text.index('"Analysis Period End"'),text.index('"Run Analysis"')]
        self.assertEqual(order,sorted(order))
    def test_combined_preview_contains_selected_jira_and_clickup(self):
        rows=build_company_preview(jira_prepared=(JiraPrepared(),),clickup_prepared=(ClickPrepared(),))
        self.assertEqual({r.source_tool for r in rows},{'Jira','ClickUp'})
        click=next(r for r in rows if r.source_tool=='ClickUp')
        self.assertEqual(click.space,'Growth');self.assertEqual(click.original_status,'Review');self.assertEqual(click.final_status,'In Review')
    def test_preview_filtering_covers_requested_fields(self):
        rows=build_company_preview(jira_prepared=(JiraPrepared(),),clickup_prepared=(ClickPrepared(),ClickPrepared(space='Ops',task_id='cu-2',status={'status':'In Progress'},priority='1',name='Launch API')))
        self.assertEqual(len(filter_company_preview(rows,source_tools=['ClickUp'])),2)
        self.assertEqual(len(filter_company_preview(rows,spaces=['Ops'])),1)
        self.assertEqual(len(filter_company_preview(rows,task_name='launch')),1)
        self.assertEqual(len(filter_company_preview(rows,original_statuses=['Review'])),1)
        self.assertEqual(len(filter_company_preview(rows,final_statuses=['In Execution'])),1)
        self.assertEqual(len(filter_company_preview(rows,assignees=['Zaher'])),2)
        self.assertEqual(len(filter_company_preview(rows,priorities=['Critical'])),1)
        self.assertEqual(len(filter_company_preview(rows,due_date_start=date(2026,9,1),due_date_end=date(2026,9,30))),3)

class ClickUpStatusTests(unittest.TestCase):
    def test_clickup_status_name_extraction_prefers_readable_fields_over_id(self):
        for key in ('status','name','status_name','label'):
            with self.subTest(key=key):
                source={'id':'p_INTERNAL',key:'In Progress'}
                result=adapt_clickup_collection([{'id':'x','name':'X','status':source,'list':{'name':'Ops'}}],unified_project='P',source_space='S')
                rec=result.records[0];self.assertEqual(rec.raw_status,'In Progress');self.assertNotIn('ClickUp Status Unresolved',rec.data_quality_flags)
    def test_clickup_status_id_fallback_is_original_unknown_and_quality_flagged(self):
        result=adapt_clickup_collection([{'id':'x','name':'X','status':{'id':'p_INTERNAL'},'list':{'name':'Ops'}}],unified_project='P',source_space='S',collection_timestamp='2026-09-30T10:00:00Z')
        rec=deduplicate_tasks(result.records)[0]
        self.assertEqual(rec.raw_status,'p_INTERNAL');self.assertIs(normalize_status('ClickUp',rec.raw_status),UnifiedStatus.UNKNOWN);self.assertIn('ClickUp Status Unresolved',rec.data_quality_flags);self.assertIn('Unmapped Status',rec.data_quality_flags)
    def test_all_approved_clickup_mappings(self):
        expected={'To Do':UnifiedStatus.NOT_STARTED,'Planning':UnifiedStatus.NOT_STARTED,'In Progress':UnifiedStatus.IN_EXECUTION,'At Risk':UnifiedStatus.AT_RISK,'Review':UnifiedStatus.IN_REVIEW,'On Hold':UnifiedStatus.ON_HOLD,'Complete':UnifiedStatus.COMPLETED,'Cancelled':UnifiedStatus.CANCELLED}
        for raw,final in expected.items():self.assertIs(normalize_status('ClickUp',raw),final)

class CompanyClickUpIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.result=build_company_analysis(period_start=date(2026,9,1),period_end=date(2026,9,30),jira_prepared=(JiraPrepared(),),clickup_prepared=(ClickPrepared(status={'id':'internal','status':'Review'}),ClickPrepared(space='Ops',task_id='cu-2',status={'status':'In Progress'},name='Launch API')),unified_project='Company Delivery')
    def test_original_and_final_status_exposed_and_final_filter_works(self):
        click=[d for d in self.result.model.task_details if d.source_tool=='ClickUp'];self.assertEqual({d.original_status for d in click},{'Review','In Progress'});self.assertEqual({d.final_status for d in click},{'In Review','In Execution'})
        filtered=build_company_dashboard(self.result.snapshots,coverages=self.result.collection.coverages,filters=DashboardFilters(statuses=(UnifiedStatus.IN_REVIEW,)))
        self.assertEqual(len(filtered.task_details),1);self.assertEqual(filtered.task_details[0].source_tool,'ClickUp')
    def test_clickup_is_in_filters_kpi_cards_and_charts(self):
        model=self.result.model;self.assertIn('ClickUp',model.filter_options.source_tools);self.assertEqual(model.kpis.total_tasks,3);self.assertEqual(next(c for c in model.cards if c.key=='current-wip').value,'2');delivery=next(c for c in model.executive_charts if c.key=='delivery-outcome');labels={p.label:p.value for p in delivery.points};self.assertEqual(labels['In Review'],1);self.assertEqual(labels['In Execution'],1);self.assertEqual(labels['Completed'],1)
    def test_excel_task_details_has_both_status_columns(self):
        with tempfile.TemporaryDirectory() as td:
            path=write_company_excel(self.result.model,Path(td)/'out.xlsx');ws=load_workbook(path,data_only=True)['Task Details'];headers=[c.value for c in ws[1]];self.assertIn('Original Status',headers);self.assertIn('Final Status',headers);rows=list(ws.iter_rows(min_row=2,values_only=True));click=[r for r in rows if r[headers.index('Source Tool')]=='ClickUp'];self.assertTrue(click);self.assertIn('Review',{r[headers.index('Original Status')] for r in click});self.assertIn('In Review',{r[headers.index('Final Status')] for r in click})
    def test_word_contains_clickup_and_normalized_delivery_status(self):
        with tempfile.TemporaryDirectory() as td:
            path=write_company_word_report(self.result.model,Path(td)/'out.docx');doc=Document(path);text='\n'.join(p.text for p in doc.paragraphs)+'\n'+'\n'.join(c.text for t in doc.tables for r in t.rows for c in r.cells);self.assertIn('ClickUp',text);self.assertIn('In Review',text);self.assertIn('In Execution',text)
    def test_unresolved_id_stays_out_of_delivery_chart_and_is_in_data_quality(self):
        unresolved=build_company_analysis(period_start=date(2026,9,1),period_end=date(2026,9,30),clickup_prepared=(ClickPrepared(status={'id':'p_INTERNAL'}),),unified_project='P')
        self.assertEqual(unresolved.model.task_details[0].original_status,'p_INTERNAL');self.assertEqual(unresolved.model.task_details[0].final_status,'Unknown');delivery=next(c for c in unresolved.model.executive_charts if c.key=='delivery-outcome');self.assertNotIn('Unknown',[p.label for p in delivery.points]);flags={i.flag for i in unresolved.model.data_quality};self.assertIn('ClickUp Status Unresolved',flags);self.assertIn('Unmapped Status',flags)

if __name__=='__main__':unittest.main()
