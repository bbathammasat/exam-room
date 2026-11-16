#!/usr/bin/env python3
"""Build exam data for one term from a zip of xlsx rosters.
Usage: python3 build_term.py ZIP --name "Final Exam 1/2026" --start 2026-11-23 --end 2026-12-08 --out term.json
Seat numbers are NOT stored. Names: --names masked (default, e.g. 'Miss Chawanrat I.'), full, or none."""
import zipfile,io,re,json,sys,argparse,datetime as dt,openpyxl
ap=argparse.ArgumentParser();ap.add_argument('zip');ap.add_argument('--name',required=True)
ap.add_argument('--start',required=True);ap.add_argument('--end',required=True);ap.add_argument('--out',default='term.json');ap.add_argument('--names',choices=['masked','full','none'],default='masked')
a=ap.parse_args()
S,E=dt.date.fromisoformat(a.start),dt.date.fromisoformat(a.end)
DRE=re.compile(r'Date:\s*(\w+)\s*,\s*(\w+)\s+(\d+)\s*,\s*(\d{4})\s+Time:\s*(\d+:\d+)\s*-\s*(\d+:\d+)\s+Exam Room#\s*(\S+)',re.I)
MON={m:i for i,m in enumerate('January February March April May June July August September October November December'.split(),1)}
exams=[];report=[]
def nm(r,ii):
    if a.names=='none':return ''
    parts=[p for p in ' '.join(str(x).strip() for x in r[ii+1:ii+3] if x).split() if p!='-']
    if not parts:return ''
    if a.names=='full':return ' '.join(parts)
    ti=1 if parts[0].rstrip('.').lower() in('mr','mrs','miss','ms') and len(parts)>1 else 0
    t=parts[:ti];f=parts[ti:ti+1];l=[p[0]+'.' for p in parts[ti+1:]]
    return ' '.join(t+f+l)
z=zipfile.ZipFile(a.zip)
for fn in sorted(n for n in z.namelist() if n.lower().endswith('.xlsx') and not n.startswith('__MACOSX')):
    code=re.sub(r'^\d+-\d{4}-','',fn.split('/')[-1][:-5])
    wb=openpyxl.load_workbook(io.BytesIO(z.read(fn)),data_only=True)
    found=False
    for ws in wb:
        rows=list(ws.iter_rows(values_only=True));title='';cur=None;hdr=None;blocks={}
        for r in rows:
            c0=r[0] if r else None
            if isinstance(c0,str):
                t=c0.replace('\xa0',' ').strip()
                m=DRE.search(t)
                if m:
                    wd,mo,d,y,t1,t2,room=m.groups()
                    date=dt.date(int(y),MON[mo.capitalize()],int(d));cur=(date,t1,t2,room);hdr=None
                    if date.strftime('%A')!=wd.capitalize(): report.append(f'{code}: weekday mismatch "{wd}" vs {date}')
                    blocks.setdefault(cur,[]);continue
                if re.match(r'^[A-Z]{2}\d{3}',t) and not title: title=t;continue
                if 'Student ID' in ''.join(str(x) for x in r if x):
                    hdr=[str(x or '').strip().lower() for x in r];continue
            if cur and hdr:
                ii=next((i for i,h in enumerate(hdr) if h.startswith('student id')),None)
                if ii is None or not isinstance(r[ii],(int,float)):continue
                si=next((i for i,h in enumerate(hdr) if h.startswith('comp')),None)
                if si is None: si=next((i for i,h in enumerate(hdr) if h.startswith('seat')),None)
                if si is None: si=next(i for i,h in enumerate(hdr) if h=='list')
                sv=r[si];sv=str(int(sv)) if isinstance(sv,float) and sv==int(sv) else str(sv).strip();blocks[cur].append([str(int(r[ii])),sv,nm(r,ii)])
        if blocks:
            found=True
            for (date,t1,t2,room),rs in blocks.items():
                if not rs:continue
                if not(S<=date<=E):
                    report.append(f'EXCLUDED {code} [{ws.title}] {date} room {room}: outside term range ({len(rs)} students)');continue
                exams.append(dict(code=code,title=title,date=date.isoformat(),start=t1.zfill(5),end=t2.zfill(5),room=room,rows=rs))
    if not found: report.append(f'NO DATA {code}')
# merge rooms per exam, checks
ex={}
for e in exams:
    k=(e['code'],e['date'],e['start'])
    x=ex.setdefault(k,dict(code=e['code'],title=e['title'],date=e['date'],start=e['start'],end=e['end'],rooms={}))
    x['rooms'].setdefault(e['room'],[]).extend(e['rows'])
for k,x in ex.items():
    ids=[r[0] for v in x['rooms'].values() for r in v]
    d=len(ids)-len(set(ids))
    if d:report.append(f'{x["code"]}: {d} duplicate student ID(s)')
    bad=[i for i in ids if len(i)!=10]
    if bad:report.append(f'{x["code"]}: non-10-digit IDs {bad[:3]}')
    seats={}
    for rm,v in x['rooms'].items():
        for i,s,_ in v:seats.setdefault((rm,s),[]).append(i)
    dd=sum(1 for v in seats.values() if len(v)>1)
    if dd:report.append(f'{x["code"]}: {dd} seat(s) assigned to more than one student')
out=dict(name=a.name,start=a.start,end=a.end,exams=sorted(ex.values(),key=lambda x:(x['date'],x['start'],x['code'])))
for x in out['exams']:x['rooms']={k:[[i,n] if a.names!='none' else i for i,_,n in v] for k,v in sorted(x['rooms'].items())}  # seat numbers are never published
json.dump(out,open(a.out,'w',encoding='utf-8'),ensure_ascii=False,separators=(',',':'))
n=sum(len(v) for x in out['exams'] for v in x['rooms'].values())
print(f'{len(out["exams"])} exams, {n} seat records -> {a.out}');print('\n'.join(report) or 'no issues')
