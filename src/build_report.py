"""One-command report build. Profiles/plans precede export; verification never fakes review."""
from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta, timezone
import json
import math
from pathlib import Path
import re
import shutil
import sys
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
import xlsxwriter
from openpyxl import load_workbook
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether

from load import load_table, classify, profile, digest, NeedsInput
from analysis import discover, clean, capabilities, analyze, num
from charts import build_charts
from audit import save_audit

ROOT = Path(__file__).resolve().parents[1]
SHEETS = ['Summary','Findings','Charts','Recommendations','Data_Summary','Cleaned_Data','Notes']
CHECKS = [
    'Loaded correctly: shape and column count are plausible',
    'Raw file unchanged',
    'Capability check written down; no missing-field analysis',
    'Cleaning log row counts reconcile',
    'Three numbers independently spot-checked against pandas',
    'Category parts sum to total',
    'Charts have titles, unit axes and figure references',
    'Findings have a number, comparison and implication',
    'Recommendations link to findings and show impact methods',
    'Limitations list missing fields and assumptions',
    'No unresolved placeholders or copied example numbers',
    'Report opens; PDF pages and workbook sheets visually checked',
    'README explains a one-command rebuild',
]


def json_default(value):
    if isinstance(value, (np.integer,np.floating)):
        return value.item()
    if isinstance(value, (date,datetime,pd.Timestamp)):
        return value.isoformat()
    if value is pd.NA or value is pd.NaT:
        return None
    return str(value)


def save_json(path, data):
    # Strict JSON, including NaN from all-missing profile columns.
    encoded = json.dumps(data, default=json_default, ensure_ascii=False, indent=2)
    encoded = re.sub(r'(?<=[:\s])NaN(?=[,\s\]}])', 'null', encoded)
    Path(path).write_text(encoded, encoding='utf-8')


def metric_definitions(data, cap):
    metric = cap['metric']
    result = {
        'Retained records':'Count of rows after the logged cleaning steps; not necessarily transactions or customers.',
        'Observed metric records':f"Count of non-missing {metric}." if metric else 'All retained records.',
        'Primary total':f"Sum of observed {metric}; unit: {cap['unit']}." if metric else 'Number of retained records.',
        'Group share':'Group primary total divided by total primary amount across all groups, including missing-category group.',
        'Group mean':'Sum of observed outcome divided by its non-missing row count within that group.',
        'Relative mean gap':'(Higher group mean - lower group mean) / absolute lower group mean; unavailable when baseline is zero.',
        'Correlation':'Pearson correlation on complete numeric pairs; zero is the no-linear-association comparison, not proof of independence.',
    }
    if metric == 'revenue_est':
        result['Estimated revenue'] = f"{cap['roles']['price']} x {cap['roles']['quantity']}; requires both inputs. No discounts/taxes inferred."
    return result


def render_pdf(report, path):
    fonts = Path('C:/Windows/Fonts')
    family, bold = 'Helvetica','Helvetica-Bold'
    if (fonts/'arial.ttf').exists():
        pdfmetrics.registerFont(TTFont('ReportArial', str(fonts/'arial.ttf')))
        pdfmetrics.registerFont(TTFont('ReportArialBold', str(fonts/'arialbd.ttf')))
        family,bold = 'ReportArial','ReportArialBold'
    styles = getSampleStyleSheet()
    for name in ('Normal','BodyText','Title','Heading1','Heading2','Heading3'):
        styles[name].fontName = bold if name in ('Title','Heading1','Heading2','Heading3') else family
        styles[name].textColor = colors.HexColor('#203247')
    styles['Title'].fontSize, styles['Title'].leading = 25,29
    styles['Heading1'].fontSize,styles['Heading1'].leading = 17,21
    styles['Heading2'].fontSize,styles['Heading2'].leading = 12,16
    styles['BodyText'].fontSize,styles['BodyText'].leading = 10,14
    styles.add(ParagraphStyle('Small',fontName=family,fontSize=9,leading=12,textColor=colors.HexColor('#334155')))
    styles.add(ParagraphStyle('TableHeader',fontName=bold,fontSize=9,leading=12,textColor=colors.white))
    def p(text, style='BodyText'):
        return Paragraph(escape(str(text)).replace('\n','<br/>'),styles[style])
    def table(headers, rows, widths):
        values = [[p(x,'TableHeader') for x in headers]] + [[p(x,'Small') for x in row] for row in rows]
        t=Table(values,colWidths=widths,repeatRows=1,hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#203247')),
            ('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f1f5f9')]),
            ('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),
            ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
        return t
    width=A4[0]-4*cm
    story=[p('Business findings report','Title'),p(report['dataset'],'Heading2'),
           p(f"Period: {report['facts']['period']} | Generated: {report['generated']} (Asia/Calcutta)",'Small')]
    if report['synthetic']:
        story += [Spacer(1,10),p('SYNTHETIC EXAMPLE - fictional data for report review','Heading2')]
    elif report['capabilities'].get('scope_note'):
        story += [Spacer(1,10),p(report['capabilities']['scope_note'],'Small')]
    story += [Spacer(1,14),p('Executive summary','Heading1')]
    for line in report['summary'][:5]:
        story.append(p(line,'Small'))
    story += [Spacer(1,16),p('Data overview and methodology','Heading1')]
    f=report['facts']
    overview=[['Source rows', f"{report['source_rows']:,}"],['Retained rows',f"{f['rows']:,}"],
              ['Rows removed',f"{report['source_rows']-f['rows']:,} ({report['removed_pct']:.1f}%)"],
              ['Primary metric',f"{report['capabilities']['metric_label']} ({f['unit']})"],
              ['Observed metric coverage',f"{f['valid_metric_rows']:,} / {f['rows']:,} records"],
              ['Columns used',', '.join(report['used_columns']) or 'Record count only']]
    if not report['capabilities']['metric'] and 'outcome_mean' in f:
        overview.insert(-1,[f"Mean {f['outcome_column']}",f"{num(f['outcome_mean'],2)} per observed record"])
    story += [table(['Measure','Definition or value'],overview,[width*.3,width*.7]),Spacer(1,10)]
    story.append(p('Raw data is preserved. Exact duplicate rows are removed; category labels are normalized. Missing cells remain missing, and only observed values contribute to each metric. See the appendix for the complete cleaning log and column dictionary.','Small'))
    if report['capabilities']['metric']=='revenue_est':
        story.append(p(report['definitions']['Estimated revenue'],'Small'))
    story.append(p('Assumptions and skipped analyses are listed under Limitations and next steps.','Small'))
    for i in range(0,len(report['findings']),2):
        story += [PageBreak(),p('Key findings','Heading1')]
        for finding in report['findings'][i:i+2]:
            story.append(KeepTogether([p(f"Figure {finding['figure']} / {finding['id']}",'Heading2'),
                Image(finding['chart'],width=width,height=width*4.2/8),p(finding['text'],'Small'),Spacer(1,9)]))
    story += [PageBreak(),p('Recommendations','Heading1'),p('Prioritize controlled tests. Scenario ranges are conditional, not forecasts.','BodyText'),Spacer(1,10)]
    rows=[[r['priority'],r['action'],r['finding'],r['impact']] for r in report['recommendations']]
    if rows:
        story.append(table(['Priority','Action / affected group','Finding','Expected impact / method'],rows,[width*.11,width*.32,width*.1,width*.47]))
    else:
        story.append(p('No finding-linked recommendation is supported by the available evidence.'))
    story += [PageBreak(),p('Limitations and next steps','Heading1')]
    for heading,items in [('Missing fields and skipped analyses',report['capabilities']['skipped']),
                          ('Assumptions',report['capabilities']['assumptions']),
                          ('Data quality and interpretation',report['limitations'])]:
        story.append(p(heading,'Heading2'))
        for item in items:
            story.append(p(item,'Small'))
            story.append(Spacer(1,5))
    story += [p('Next steps','Heading2'),p('Validate field meanings with the data owner, obtain missing cost/customer context if needed, and rerun with the next actual dataset. Observational findings cannot establish causal effects.','Small')]
    story += [PageBreak(),p('Appendix: cleaning log','Heading1')]
    rows=[[x['step'],x['rows_before'],x['rows_after'],x['affected_cells'],x['reason']] for x in report['cleaning_log']]
    story.append(table(['Step','Before','After','Cells','Reason'],rows,[width*.24,width*.09,width*.09,width*.08,width*.5]))
    story += [PageBreak(),p('Appendix: metric definitions','Heading1'),
              table(['Metric','Definition'],list(report['definitions'].items()),[width*.28,width*.72]),
              Spacer(1,16),p('Column dictionary','Heading1')]
    rows=[[x['column'],x['class'],f"{x['null_count']} ({x['null_pct']:.1f}%)",x['unique_count']] for x in report['profile']['columns']]
    story.append(table(['Column','Classification','Source nulls','Unique'],rows,[width*.3,width*.3,width*.23,width*.17]))
    story += [Spacer(1,10),p(f"Source SHA-256: {report['source_sha256']}",'Small')]
    def page_number(canvas,doc):
        canvas.setFont(family,8)
        canvas.setFillColor(colors.HexColor('#64748b'))
        canvas.drawString(2*cm,1.25*cm,'Business findings' + (' | Synthetic example' if report['synthetic'] else ''))
        canvas.drawRightString(A4[0]-2*cm,1.25*cm,str(doc.page))
    doc=SimpleDocTemplate(str(path),pagesize=A4,leftMargin=2*cm,rightMargin=2*cm,topMargin=2*cm,bottomMargin=2*cm,
                          title='Business findings report',author='Muhammad Akram')
    doc.build(story,onFirstPage=page_number,onLaterPages=page_number)
    reader=PdfReader(path)
    main_pages=next((i for i,page in enumerate(reader.pages) if 'Appendix:' in (page.extract_text() or '')),len(reader.pages))
    if main_pages>8:
        raise NeedsInput(f"Main report has {main_pages} pages; tighten the layout before delivery.")
    return {'pages':len(reader.pages),'main_pages':main_pages}


def build_excel(report, data, path):
    wb=xlsxwriter.Workbook(str(path),{'strings_to_formulas':False,'strings_to_urls':False,'nan_inf_to_errors':False})
    wb.set_calc_mode('auto')
    body=wb.add_format({'font_name':'Arial','font_size':11,'valign':'top','text_wrap':True})
    header=wb.add_format({'font_name':'Arial','font_size':11,'bold':True,'font_color':'white','bg_color':'#203247','valign':'vcenter','text_wrap':True})
    integer=wb.add_format({'font_name':'Arial','font_size':11,'num_format':'[$-409]#,##0','valign':'top'})
    decimal=wb.add_format({'font_name':'Arial','font_size':11,'num_format':'[$-409]#,##0.00','valign':'top'})
    pct=wb.add_format({'font_name':'Arial','font_size':11,'num_format':'0.0%','valign':'top'})
    date_fmt=wb.add_format({'font_name':'Arial','font_size':11,'num_format':'yyyy-mm-dd','valign':'top'})
    currency=report['capabilities']['currency']
    money=wb.add_format({'font_name':'Arial','font_size':11,'num_format':f'[$-409]"{currency}" #,##0.00' if currency else '[$-409]#,##0.00','valign':'top'})
    sheets={name:wb.add_worksheet(name) for name in SHEETS}
    for sh in sheets.values():
        sh.hide_gridlines(2); sh.freeze_panes(1,0); sh.set_default_row(22)
        sh.set_landscape(); sh.fit_to_pages(1,0); sh.repeat_rows(0)
    def grid(name, headers, rows, widths):
        sh=sheets[name]
        for c,(h,w) in enumerate(zip(headers,widths)):
            sh.set_column(c,c,w,body); sh.write(0,c,h,header)
        sh.set_row(0,32)
        for r,row in enumerate(rows,1):
            lines=1
            for c,value in enumerate(row):
                if value is None or (not isinstance(value,str) and pd.isna(value)):
                    continue
                fmt=decimal if isinstance(value,(float,np.floating)) else integer if isinstance(value,(int,np.integer)) else body
                sh.write(r,c,value,fmt)
                if isinstance(value,str):
                    lines=max(lines,sum(max(1,math.ceil(len(line)/max(8,widths[c]-3))) for line in value.splitlines()))
            sh.set_row(r,min(400, max(24,lines*15+7)))
        return sh
    facts,cap=report['facts'],report['capabilities']
    summary=[['Dataset',report['dataset']],['Covered period',facts['period']],['Generated',report['generated']],
             ['Status','SYNTHETIC EXAMPLE - fictional data' if report['synthetic'] else cap.get('scope_note') or 'User-supplied data; interpretation subject to stated limitations'],
             ['Primary metric',f"{cap['metric_label']} ({cap['unit']})"],['Retained records',None],['Primary total',None],['Observed mean',None]]
    if not cap['metric'] and 'outcome_mean' in facts:
        summary[7][0]=f"Mean {facts['outcome_column']} per record"
    summary += [[f"Finding {f['id']}",f['title']] for f in report['findings'][:6]]
    grid('Summary',['Business findings report','Result'],summary,[28,110])
    grid('Findings',['Finding','Figure','Title','Interpretation'],[[f['id'],f['figure'],f['title'],f['text']] for f in report['findings']],[10,10,52,112])
    chart_sheet=grid('Charts',['Figure','Chart previews (snapshots; rerun to refresh)'],[],[10,110])
    for i,f in enumerate(report['findings']):
        row=2+i*25
        chart_sheet.write(row,0,f['figure'],integer)
        chart_sheet.insert_image(row,1,f['chart'],{'x_scale':1.0,'y_scale':1.0,'object_position':1,'description':f['title']})
    grid('Recommendations',['Priority','Action / affected group','Linked finding','Expected impact / calculation'],
         [[r['priority'],r['action'],r['finding'],r['impact']] for r in report['recommendations']],[12,66,15,95])
    # Keep identifiers; drop only large text fields, documenting every omission.
    omitted=[c for c in data if report['classes'].get(c)=='Free text / URL' and data[c].astype(str).str.len().sum()>100000]
    export=data.drop(columns=omitted)
    sh=sheets['Cleaned_Data']
    for c,col in enumerate(export):
        sh.write(0,c,col,header)
        sh.set_column(c,c,min(40,max(18,len(col)+3)),body)
        if pd.api.types.is_datetime64_any_dtype(export[col]):
            sh.set_column(c,c,18,date_fmt)
        elif pd.api.types.is_numeric_dtype(export[col]):
            sh.set_column(c,c,23,money if col==cap['metric'] and (col=='revenue_est' or col==cap['roles'].get('revenue')) else decimal)
        for r,value in enumerate(export[col],1):
            if pd.isna(value):
                continue
            if isinstance(value,pd.Timestamp):
                sh.write_datetime(r,c,value.to_pydatetime(),date_fmt)
            elif isinstance(value,(np.integer,np.floating)):
                sh.write_number(r,c,float(value))
            else:
                sh.write(r,c,value)
    sh.set_row(0,32)
    for r, row in enumerate(export.itertuples(index=False,name=None),1):
        lines=1
        for c,value in enumerate(row):
            if isinstance(value,str):
                w=min(40,max(18,len(str(export.columns[c]))+3))
                lines=max(lines,math.ceil(len(value)/max(8,w-3)))
        sh.set_row(r,min(400,max(22,lines*15+7)))
    sh.autofilter(0,0,len(export),len(export.columns)-1)
    metric=cap['metric']
    column_indices={c:i for i,c in enumerate(export.columns)}
    def range_for(col):
        letter=xlsxwriter.utility.xl_col_to_name(column_indices[col])
        return f"Cleaned_Data!${letter}$2:${letter}${len(export)+1}"
    summaries=[['Retained records',None,'Count of cleaned rows'],['Primary total',None,report['definitions']['Primary total']],
               ['Observed metric records',None,report['definitions']['Observed metric records']],
               ['Mean per observed record',None,'Primary total / observed metric records']]
    for col in report['groups']:
        for group in report['groups'][col]:
            summaries.append([f"{col} = {group['group']}",None,'Group primary total'])
            summaries.append([f"Share: {col} = {group['group']}",None,'Group primary total / primary total'])
    comparison_findings=[f for f in report['findings'] if f.get('group_stats')]
    for finding in comparison_findings:
        for group in finding['group_stats']:
            name=f"{finding['group']} = {group['label']}"
            summaries += [[name+' / observed records',None,f"Nonmissing {finding['outcome']} count"],
                          [name+' / recorded volume sum',None,f"SUMIFS of {finding['outcome']}; arithmetic over dataset records, not independently confirmed business sales"],
                          [name+' / mean',None,f"Recorded {finding['outcome']} sum / observed records"]]
        summaries.append([finding['group']+' / relative mean gap',None,'(Highest mean - lowest mean) / absolute lowest mean; descriptive association, not causal'])
    overall_outcome_row=len(summaries)+1
    if not cap['metric'] and 'outcome_mean' in facts:
        summaries += [[facts['outcome_column']+' / observed count',None,'Count of nonmissing recorded outcomes'],
                      [facts['outcome_column']+' / arithmetic sum',None,'Dataset arithmetic only; not independently verified business sales'],
                      [facts['outcome_column']+' / overall mean',None,'Recorded volume sum / observed outcome count']]
    stats=grid('Data_Summary',['Metric','Live calculated value','Definition / scope'],summaries,[48,26,88])
    stats.write_formula(1,1,f"=ROWS(Cleaned_Data!A2:A{len(export)+1})",integer,facts['rows'])
    if metric:
        rng=range_for(metric)
        stats.write_formula(2,1,f'=SUM({rng})',money if metric=='revenue_est' or metric==cap['roles'].get('revenue') else decimal,facts['total'])
        stats.write_formula(3,1,f'=COUNT({rng})',integer,facts['valid_metric_rows'])
    else:
        stats.write_formula(2,1,'=B2',integer,facts['total'])
        stats.write_formula(3,1,'=B2',integer,facts['valid_metric_rows'])
    stats.write_formula(4,1,'=IF(B4=0,"n.a.",B3/B4)',decimal,facts['mean'])
    rr=5
    for col,groups in report['groups'].items():
        for group in groups:
            # Escape criteria wildcards, literal leading operators, and quotation marks.
            val=str(group['group'])
            missing=val not in set(data[col].dropna().astype(str))
            criteria='' if missing else '=' + val.replace('~','~~').replace('*','~*').replace('?','~?')
            criteria=criteria.replace('"','""')
            formula=f'=SUMIFS({range_for(metric)},{range_for(col)},"{criteria}")' if metric else f'=COUNTIFS({range_for(col)},"{criteria}")'
            stats.write_formula(rr,1,formula,decimal,group['sum'])
            share=group['sum']/facts['total'] if facts['total'] else 'n.a.'
            stats.write_formula(rr+1,1,f'=IF($B$3=0,"n.a.",B{rr+1}/$B$3)',pct,share)
            rr+=2
    for finding in comparison_findings:
        mean_rows={}
        for group in finding['group_stats']:
            literal='='+str(group['label']).replace('~','~~').replace('*','~*').replace('?','~?')
            literal=literal.replace('"','""')
            by_range=range_for(finding['group']); value_range=range_for(finding['outcome'])
            stats.write_formula(rr,1,f'=COUNTIFS({by_range},"{literal}",{value_range},"<>")',integer,group['count'])
            stats.write_formula(rr+1,1,f'=SUMIFS({value_range},{by_range},"{literal}")',decimal,group['total'])
            stats.write_formula(rr+2,1,f'=IF(B{rr+1}=0,"n.a.",B{rr+2}/B{rr+1})',decimal,group['mean'])
            mean_rows[group['label']]=rr+3
            rr+=3
        hi=mean_rows[finding['higher_group']];lo=mean_rows[finding['lower_group']]
        gap=finding['difference']/abs(finding['lower_mean']) if finding['lower_mean'] else 'n.a.'
        stats.write_formula(rr,1,f'=IF(B{lo}=0,"n.a.",(B{hi}-B{lo})/ABS(B{lo}))',pct,gap)
        rr+=1
    if not cap['metric'] and 'outcome_mean' in facts:
        row=overall_outcome_row
        rng=range_for(facts['outcome_column'])
        stats.write_formula(row,1,f'=COUNT({rng})',integer,facts['outcome_count'])
        stats.write_formula(row+1,1,f'=SUM({rng})',decimal,facts['outcome_total'])
        stats.write_formula(row+2,1,f'=IF(B{row+1}=0,"n.a.",B{row+2}/B{row+1})',decimal,facts['outcome_mean'])
    # Derived estimated revenue is live in the source-detail tab as well.
    if metric=='revenue_est':
        qty,price=[xlsxwriter.utility.xl_col_to_name(column_indices[cap['roles'][r]]) for r in ('quantity','price')]
        for r in range(1,len(export)+1):
            q,p=f'{qty}{r+1}',f'{price}{r+1}'
            cached=data[metric].iloc[r-1]
            sh.write_formula(r,column_indices[metric],f'=IF(AND(ISNUMBER({q}),ISNUMBER({p})),{q}*{p},"")',money,'' if pd.isna(cached) else float(cached))
    for row,formula,value,fmt in [(6,'=Data_Summary!B2',facts['rows'],integer),(7,'=Data_Summary!B3',facts['total'],money if metric=='revenue_est' else decimal),(8,'=Data_Summary!B5',facts['mean'],decimal)]:
        sheets['Summary'].write_formula(row,1,formula,fmt,value)
    if not cap['metric'] and 'outcome_mean' in facts:
        sheets['Summary'].write_formula(8,1,f'=Data_Summary!B{overall_outcome_row+3}',decimal,facts['outcome_mean'])
    note_rows=[['Source',report['source']],['SHA-256',report['source_sha256']],['Input shape',str(report['load']['shape'])],
               ['Update behavior','Data_Summary and Summary numeric formulas recalculate from Cleaned_Data. Narrative, findings, recommendations and PNG charts are snapshots; rerun the Python command to refresh them.'],
               ['Omitted free text',', '.join(omitted) or 'None'],['Null policy','Keep nulls; no zero imputation. Available-case denominators are stated.']]
    note_rows += [['Assumption',x] for x in cap['assumptions']]
    note_rows += [['Missing / skipped',x] for x in cap['skipped']]
    note_rows += [['Limitation',x] for x in report['limitations']]
    note_rows += [[f"Metric: {k}",v] for k,v in report['definitions'].items()]
    note_rows += [[f"Cleaning: {x['step']}",f"{x['rows_before']} -> {x['rows_after']} rows; affected cells {x['affected_cells']}. {x['reason']}"] for x in report['cleaning_log']]
    note_rows += [[f"Column: {x['column']}",json.dumps(x,default=json_default,ensure_ascii=False)] for x in report['profile']['columns']]
    grid('Notes',['Topic','Evidence / method / limitation'],note_rows,[45,125])
    wb.close()
    reopened=load_workbook(path,data_only=False)
    assert reopened.sheetnames==SHEETS
    assert reopened['Cleaned_Data'].max_row==len(export)+1
    assert reopened['Data_Summary']['B3'].data_type=='f'
    reopened.close()
    cached=load_workbook(path,data_only=True)
    assert math.isclose(cached['Data_Summary']['B3'].value,facts['total'],rel_tol=1e-10)
    cached.close()
    return {'sheets':SHEETS,'formula_caches_checked':True,'omitted_columns':omitted}


def write_plan(path, meta, cap, blockers):
    lines=['# Analysis plan',f"Source: {Path(meta['source']).name}",f"Loaded shape: {meta['shape']}",'',
           '## Capability check',f"Primary metric: {cap['metric_label']} ({cap['unit']})",'']
    for q in cap['questions']:
        lines += [f"- {q['question']} Columns: {', '.join(q['columns'])}. Chart: {q['chart']}."]
    lines += ['','## Assumptions'] + ['- '+x for x in cap['assumptions']]
    lines += ['','## Skipped analyses'] + ['- '+x for x in cap['skipped']]
    lines += ['','## Decisions required'] + (['- '+x for x in blockers] or ['None detected; review semantic mappings before export.'])
    path.write_text('\n'.join(lines)+'\n',encoding='utf-8')


def report_markdown(report):
    lines=['# Business findings report',f"Dataset: {report['dataset']}",f"Period: {report['facts']['period']}",f"Generated: {report['generated']}",'']
    if report['synthetic']:
        lines+=['SYNTHETIC EXAMPLE - fictional demonstration, not actual business results.','']
    lines+=['## Executive summary']+['- '+s for s in report['summary']]
    lines+=['','## Data overview and methodology',f"{report['source_rows']:,} input records; {report['facts']['rows']:,} retained; {report['removed_pct']:.1f}% removed.",f"Columns used: {', '.join(report['used_columns'])}."]
    lines+=['','## Key findings']
    for f in report['findings']:
        lines += [f"### {f['id']}: {f['title']}",f"Figure {f['figure']}: {Path(f['chart']).name}",f['text'],'']
    lines+=['## Recommendations','| Priority | Action | Finding | Impact / method |','|---|---|---|---|']
    for r in report['recommendations']:
        lines.append('| '+' | '.join(str(r[k]).replace('|','/') for k in ['priority','action','finding','impact'])+' |')
    lines+=['','## Limitations and next steps']+['- '+x for x in report['capabilities']['skipped']+report['capabilities']['assumptions']+report['limitations']]
    lines+=['','## Appendix','### Metric definitions']+['- '+k+': '+v for k,v in report['definitions'].items()]
    lines+=['','### Cleaning log']+[f"- {x['step']}: {x['rows_before']} -> {x['rows_after']}; {x['reason']}" for x in report['cleaning_log']]
    lines+=['','### Column dictionary']+[f"- {x['column']}: {x['class']}; nulls={x['null_count']} ({x['null_pct']:.1f}%); unique={x['unique_count']}." for x in report['profile']['columns']]
    return '\n'.join(lines)+'\n'


def run(input_path, output_format='both', config_path=None, plan_only=False, as_of=None, output_root=None):
    source=Path(input_path).resolve()
    output_root=Path(output_root or ROOT/'outputs')
    raw_dir=ROOT/'data'/'raw'
    raw_dir.mkdir(parents=True,exist_ok=True)
    raw=raw_dir/source.name
    before=digest(source)
    if raw.resolve()!=source:
        if raw.exists() and digest(raw)!=before:
            raw=raw_dir/(source.stem+'_'+before[:10]+source.suffix)
        if not raw.exists():
            shutil.copy2(source,raw)
        if digest(raw)!=before:
            raise NeedsInput('Raw-copy checksum mismatch.')
    automatic=source.with_suffix('.config.json')
    config_path=Path(config_path) if config_path else automatic if automatic.exists() else None
    config=json.loads(config_path.read_text(encoding='utf-8')) if config_path else {}
    as_of=as_of or datetime.now(timezone(timedelta(hours=5,minutes=30))).date()
    if isinstance(as_of,str):
        as_of=date.fromisoformat(as_of)
    slug=re.sub(r'[^a-zA-Z0-9_-]+','_',source.stem).strip('_') or 'dataset'
    out=output_root/slug
    out.mkdir(parents=True,exist_ok=True)
    frame,meta=load_table(source,config)
    print(json.dumps({k:meta[k] for k in ('shape','columns','dtypes','first_five')},default=json_default,ensure_ascii=False))
    classes,formats=classify(frame,config)
    profiling=profile(frame,classes,formats,as_of)
    roles=discover(frame,classes,config)
    data,log,limitations,blockers=clean(frame,classes,formats,roles,config,as_of)
    cap=capabilities(data,classes,roles,config)
    if cap['currency_problem']:
        blockers.append(cap['currency_problem'])
    save_json(out/'profile.json',{'load':meta,'profile':profiling,'classes':classes})
    save_json(out/'cleaning_log.json',log)
    write_plan(out/'analysis_plan.md',meta,cap,blockers)
    if plan_only:
        save_audit(out,'plan')
        print(f"PLAN: {out/'analysis_plan.md'}; blockers: {len(blockers)}")
        return {'plan':str(out/'analysis_plan.md'),'blockers':blockers,'capabilities':cap}
    if blockers:
        raise NeedsInput('\n'.join(blockers)+f"\nReview {out/'analysis_plan.md'}")
    facts,findings,recommendations,groups,extra=analyze(data,cap,classes,config)
    limitations.extend(extra)
    if facts['valid_metric_rows']==0:
        raise NeedsInput('No observed primary metric values; select a supported measure or use counts.')
    limitations.append('Associations do not prove causation. Recommendations use review coverage or explicitly conditional scenarios, not guaranteed effects.')
    for col,kind in classes.items():
        nulls=int(data[col].isna().sum())
        if nulls:
            limitations.append(f"{col}: {nulls:,} missing values remain; only relevant analyses omit these cells.")
    reference=ROOT/'example_findings_report.md'
    if reference.exists():
        # The agent handles interpretive adaptation; retain a traceable structure reference.
        ref_text=reference.read_text(encoding='utf-8')
        expected=['Executive summary','Data overview and methodology','Key findings','Recommendations','Limitations and next steps','Appendix']
        if not all(x.casefold() in ref_text.casefold() for x in expected):
            raise NeedsInput('Reference structure differs from the required sections. Have the agent adapt it before export.')
    build_charts(findings,output_root/'charts'/slug)
    processed=ROOT/'data'/'processed'/f'{slug}_cleaned.csv'
    processed.parent.mkdir(parents=True,exist_ok=True)
    data.to_csv(processed,index=False)
    used=sorted(set(c for f in findings for c in f['columns']))
    summary=[f"{facts['rows']:,} retained records; {facts['valid_metric_rows']:,} with an observed primary metric.",
             f"Total {cap['metric_label']}: {num(facts['total'])} {cap['unit']}."]+[f['title'] for f in findings[:3]]
    if not cap['metric'] and 'outcome_mean' in facts:
        summary[1]=f"Mean {facts['outcome_column']}: {num(facts['outcome_mean'],2)} per observed record."
    report={'dataset':source.name,'source':str(raw.resolve()),'source_sha256':before,'load':meta,'profile':profiling,
            'classes':classes,'source_rows':len(frame),'facts':facts,'removed_pct':(len(frame)-len(data))/len(frame)*100,
            'generated':str(as_of),'synthetic':bool(config.get('synthetic')),'capabilities':cap,'used_columns':used,
            'summary':summary,'findings':findings,'recommendations':recommendations,'groups':groups,
            'cleaning_log':log,'definitions':metric_definitions(data,cap),'limitations':list(dict.fromkeys(limitations)),
            'reference_structure':str(reference) if reference.exists() else None}
    exports={}
    if output_format in ('pdf','both'):
        exports['pdf']=render_pdf(report,out/'findings_report.pdf')
    if output_format in ('xlsx','both'):
        exports['xlsx']=build_excel(report,data,out/'findings_report.xlsx')
    (out/'findings_report.md').write_text(report_markdown(report),encoding='utf-8')
    save_json(out/'evidence.json',report)
    assert digest(source)==before and digest(raw)==before
    reconciled=all(a['rows_after']==b['rows_before'] for a,b in zip(log,log[1:])) and log[-1]['rows_after']==len(data)
    # Independent pandas calculations from the saved processed copy, not report values.
    direct=pd.read_csv(processed)
    direct_measure=direct[cap['metric']] if cap['metric'] else pd.Series(1.,index=direct.index)
    spot=[{'metric':'rows','report':facts['rows'],'direct':len(direct)},
          {'metric':'total','report':facts['total'],'direct':float(direct_measure.sum())},
          {'metric':'observed mean','report':facts['mean'],'direct':float(direct_measure.mean())}]
    if not cap['metric'] and 'outcome_mean' in facts:
        spot[2]={'metric':f"Mean {facts['outcome_column']}",'report':facts['outcome_mean'],'direct':float(direct[facts['outcome_column']].mean())}
    checked=all(math.isclose(x['report'],x['direct'],rel_tol=1e-10,abs_tol=1e-8) for x in spot)
    parts=all(math.isclose(sum(g['sum'] for g in vals),facts['total'],rel_tol=1e-10,abs_tol=1e-8) for vals in groups.values())
    raw_ok=digest(source)==before and digest(raw)==before
    statuses=[True,raw_ok,True,reconciled,checked,parts if groups else False,
              bool(findings) and all(f.get('chart') and f.get('x_label') and f.get('y_label') for f in findings),
              bool(findings),bool(recommendations),True,
              not any(x in json.dumps(report,default=json_default) for x in ['[Your answer here]','{{','TODO']),False,True]
    verification={'checks':{k:{'done':bool(v),'evidence':'Automated check passed' if v else 'Pending or not applicable; see notes'} for k,v in zip(CHECKS,statuses)},
                  'spot_checks':spot,'exports':exports,'visual_review':{'done':False,'note':'Render and inspect PDF pages and every workbook sheet before delivery.'},
                  'formula_recalculation':{'done':False,'note':'Formula caches checked; independent recalculation still requires a spreadsheet engine.'},
                  'raw_sha256':before}
    if not groups:
        verification['checks'][CHECKS[5]]['evidence']='No eligible category breakdown; not applicable, not claimed tested.'
    save_json(out/'verification.json',verification)
    assert all(statuses[:5]), 'Core verification failed'
    audit_path=save_audit(out,'automated')
    print(f'AUDIT: {audit_path}')
    print(json.dumps({'output':str(out),'findings':[f['title'] for f in findings],'verification':verification},default=json_default,indent=2))
    return report


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input',required=True)
    parser.add_argument('--format',choices=['pdf','xlsx','both'],default='both')
    parser.add_argument('--config')
    parser.add_argument('--as-of')
    parser.add_argument('--plan-only',action='store_true')
    args=parser.parse_args()
    try:
        run(args.input,args.format,args.config,args.plan_only,args.as_of)
    except NeedsInput as exc:
        slug=re.sub(r'[^a-zA-Z0-9_-]+','_',Path(args.input).stem).strip('_') or 'dataset'
        folder=ROOT/'outputs'/slug
        folder.mkdir(parents=True,exist_ok=True)
        save_audit(folder,'blocked',error=str(exc))
        print(f'REQUIRES REVIEW: {exc}',file=sys.stderr)
        raise SystemExit(2)
    except Exception as exc:
        slug=re.sub(r'[^a-zA-Z0-9_-]+','_',Path(args.input).stem).strip('_') or 'dataset'
        folder=ROOT/'outputs'/slug
        folder.mkdir(parents=True,exist_ok=True)
        save_audit(folder,'failed',error=f'{type(exc).__name__}: {exc}')
        raise


if __name__=='__main__':
    main()
