"""Append-only reporting audits, with separate hash-bound visual review evidence."""
from __future__ import annotations
import argparse
from datetime import datetime, timedelta, timezone
import hashlib
import json
from pathlib import Path
import uuid

ROOT=Path(__file__).resolve().parents[1]
ZONE=timezone(timedelta(hours=5,minutes=30))

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def save_audit(output, phase='automated', error=None, audit_root=None):
    output=Path(output)
    folder=Path(audit_root or ROOT/'audits'/'business-findings')/output.name
    folder.mkdir(parents=True,exist_ok=True)
    prior=sorted(folder.glob('*.json'))
    previous=json.loads(prior[-1].read_text(encoding='utf-8')) if prior else None
    verification_path=output/'verification.json'
    verification=json.loads(verification_path.read_text(encoding='utf-8')) if verification_path.exists() and not error and phase!='plan' else {}
    issues=[]
    for i,(name,check) in enumerate(verification.get('checks',{}).items(),1):
        if not check['done']:
            issues.append({'id':f'CHECK-{i:02d}','classification':'verification gap','description':name,'evidence':check.get('evidence','')})
    if error:
        issues.append({'id':'BUILD-ERROR','classification':'confirmed blocker','description':str(error),'evidence':phase})
    if phase=='plan':
        issues.append({'id':'PLAN-ONLY','classification':'verification gap','description':'Profiling completed; report export and delivery checks have not run.','evidence':str(output/'analysis_plan.md')})
    old={x['id']:x for x in (previous or {}).get('issues',[])}
    current={x['id']:x for x in issues}
    transitions=[]
    for ident in sorted(set(old)|set(current)):
        if ident in old and ident in current:status='still open'
        elif ident in current:status='new'
        else:
            names=list(verification.get('checks',{}))
            index=int(ident.split('-')[1])-1 if ident.startswith('CHECK-') else None
            fresh=(index is not None and index<len(names) and verification['checks'][names[index]]['done'])
            if ident=='BUILD-ERROR':fresh=phase in {'automated','reviewed'} and bool(verification)
            if ident=='PLAN-ONLY':fresh=bool(verification)
            status='resolved' if fresh else 'not rechecked'
            if not fresh:issues.append(old[ident])
        transitions.append({'id':ident,'status':status})
    now=datetime.now(ZONE)
    stamp=now.strftime('%Y-%m-%d_%H%M%S_%f')+'_'+uuid.uuid4().hex[:6]
    record={'scope':'Business findings workflow only; not a full AIOS Four-Cs score',
            'timestamp':now.isoformat(),'phase':phase,'dataset':output.name,
            'previous':str(prior[-1]) if prior else None,'issues':issues,'transitions':transitions,
            'checks':verification.get('checks',{}),'spot_checks':verification.get('spot_checks',[]),
            'formula_recalculation':verification.get('formula_recalculation',{}),'visual_review':verification.get('visual_review',{}),
            'artifacts':{str(p):sha(p) for p in output.iterdir() if p.is_file() and p.suffix in {'.pdf','.xlsx'}},
            'next_improvement':issues[0]['description'] if issues else 'Validate the next actual user dataset; synthetic tests do not prove all unseen schemas.'}
    json_path=folder/(stamp+'.json')
    json_path.write_text(json.dumps(record,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    lines=['# Business findings audit',f"Date: {now.isoformat()}",f"Dataset: {output.name}",f"Phase: {phase}",record['scope'],'',
           f"Previous audit: {record['previous'] or 'None (baseline)'}",'','## Verification checklist']
    lines += [f"- [{'x' if v['done'] else ' '}] {k}: {v.get('evidence','')}" for k,v in record['checks'].items()]
    lines += ['','## Findings']
    lines += [f"- {x['id']} [{x['classification']}]: {x['description']} Evidence: {x['evidence']}" for x in issues] or ['No open report-check issues in the inspected scope.']
    lines += ['','## Changes since previous audit']
    lines += [f"- {x['id']}: {x['status']}" for x in transitions] or ['New baseline or no issue transitions.']
    lines += ['','## Evidence',f"Verification: {verification_path}",f"Visual review: {json.dumps(record['visual_review'])}",
              f"Formula recalculation: {json.dumps(record['formula_recalculation'])}",'','## Next improvement',record['next_improvement'],
              'Fix a verified defect, add a regression test, rebuild, inspect affected outputs, and append a fresh audit. Edited instructions alone are not resolution evidence.']
    md=json_path.with_suffix('.md');md.write_text('\n'.join(lines)+'\n',encoding='utf-8')
    assert json.loads(json_path.read_text(encoding='utf-8'))['timestamp']==record['timestamp']
    return md

def finalize_review(output, review_path):
    output=Path(output)
    review=json.loads(Path(review_path).read_text(encoding='utf-8'))
    for suffix in ('pdf','xlsx'):
        target=output/f'findings_report.{suffix}'
        if target.exists() and review.get('artifact_sha256',{}).get(suffix)!=sha(target):
            raise ValueError(f'Review is stale or missing for {suffix}. Inspect the current export.')
    if (output/'findings_report.pdf').exists() and len(set(review.get('pdf_pages',[])))<2:
        raise ValueError('At least two distinct PDF pages must be inspected.')
    if (output/'findings_report.xlsx').exists():
        required={'Summary','Findings','Charts','Recommendations','Data_Summary','Cleaned_Data','Notes'}
        if set(review.get('workbook_sheets',[]))!=required:raise ValueError('Every workbook sheet must be reviewed.')
    if not review.get('reviewer') or not review.get('notes'):raise ValueError('Actual reviewer observations are required.')
    verify_path=output/'verification.json'
    verification=json.loads(verify_path.read_text(encoding='utf-8'))
    verification['visual_review']={'done':True,**review}
    verification['formula_recalculation']=review.get('formula_recalculation',verification['formula_recalculation'])
    for name,check in verification['checks'].items():
        if name.startswith('Report opens;'):check.update(done=True,evidence=f"Inspected current exports; review: {Path(review_path).resolve()}")
    verify_path.write_text(json.dumps(verification,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    return save_audit(output,'reviewed')

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',required=True);parser.add_argument('--review-file')
    args=parser.parse_args()
    print(finalize_review(args.output,args.review_file) if args.review_file else save_audit(args.output))
