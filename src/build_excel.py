import json, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter as L
o=json.load(open('results/results.json'))
df=pd.read_excel('data/1--AI_Job_Market_Dataset.xlsx',sheet_name=1)
N=len(df)+1
wb=Workbook()
F=lambda **k:Font(name='Arial',size=k.pop('size',10),**k)
HF=PatternFill('solid',fgColor='1F3A5F');BAND=PatternFill('solid',fgColor='EEF2F8')
def hdr(ws,r,vals,c0=1):
    for i,v in enumerate(vals):
        c=ws.cell(r,c0+i,v);c.font=F(bold=True,color='FFFFFF');c.fill=HF;c.alignment=Alignment(horizontal='center',wrap_text=True)
def title(ws,t,sub=None):
    ws['A1']=t;ws['A1'].font=F(bold=True,size=14,color='1F3A5F')
    if sub: ws['A2']=sub;ws['A2'].font=F(italic=True,color='666666')
def style(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font.name!='Arial': c.font=F()
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[L(i)].width=x
rng=lambda col:f"Data!${col}$2:${col}${N}"
# DATA
wd=wb.active;wd.title='Data'
for j,cn in enumerate(df.columns,1): wd.cell(1,j,cn)
hdr(wd,1,list(df.columns)+['n_skills'])
for i,row in enumerate(df.itertuples(index=False),2):
    for j,v in enumerate(row,1): wd.cell(i,j,v.item() if hasattr(v,'item') else v)
    wd.cell(i,20,f"=SUM(J{i}:N{i})")
wd.freeze_panes='A2';wd.auto_filter.ref=f"A1:T{N}";widths(wd,[8,26,26,16,11,10,12,10,11,10,10,10,12,10,10,10,10,10,10,9])
wd.sheet_properties.tabColor='888888'
def pivot(ws,r0,name,labels,col,numeric=False,fmt_label=None):
    hdr(ws,r0,[name,'Postings','Total openings','Avg salary'])
    for i,l in enumerate(labels,1):
        r=r0+i;ws.cell(r,1,l)
        ws.cell(r,2,f"=COUNTIFS({rng(col)},$A{r})");ws.cell(r,3,f"=SUMIFS({rng('S')},{rng(col)},$A{r})");ws.cell(r,4,f"=AVERAGEIFS({rng('O')},{rng(col)},$A{r})")
        ws.cell(r,2).number_format=ws.cell(r,3).number_format=ws.cell(r,4).number_format='#,##0'
    return r0+1,r0+len(labels)
def chart(ws,kind,t,r1,r2,valcol,anchor,catcol=1,w=15,h=7.5,ytitle=None,lo=None):
    ch=BarChart() if kind=='bar' else LineChart()
    if kind=='bar': ch.type='col'
    ch.title=t;ch.height=h;ch.width=w;ch.legend=None
    ch.add_data(Reference(ws,min_col=valcol,min_row=r1-1,max_row=r2),titles_from_data=True)
    ch.set_categories(Reference(ws,min_col=catcol,min_row=r1,max_row=r2))
    ch.y_axis.delete=False;ch.x_axis.delete=False
    if lo is not None: ch.y_axis.scaling.min=lo
    ws.add_chart(ch,anchor);return ch
def sorted_u(c): return sorted(df[c].unique())
# DEMAND
wdm=wb.create_sheet('Job Demand');title(wdm,'Job demand analysis','All figures are live formulas over the Data sheet (openings = sum of job_openings).')
r=4;specs=[('Job title','B'),('Industry','D'),('Country','E'),('Company size','C'),('Posting year','Q'),('Posting month','P'),('Hiring urgency','R')]
src={'B':'job_title','D':'company_industry','E':'country','C':'company_size','Q':'job_posting_year','P':'job_posting_month','R':'hiring_urgency'}
a=0
for nm,col in specs:
    labs=[int(x) if col in 'QP' else x for x in sorted_u(src[col])]
    r1,r2=pivot(wdm,r,nm,labs,col)
    chart(wdm,'line' if col in 'QP' else 'bar',f'Openings by {nm.lower()}',r1,r2,3,f"G{r-0}",w=14,h=6.5)
    r=max(r2+3,r+14)
widths(wdm,[24,12,15,13,3,3,10])
# WORKFORCE
ww=wb.create_sheet('Workforce Analytics');title(ww,'Workforce analytics','Average salary as recorded (currency and period are not stated in the source file).')
r=4
for nm,col in [('Experience level','G'),('Company size','C'),('Education','I'),('Remote type','F'),('Country','E'),('Industry','D'),('Job title','B')]:
    r1,r2=pivot(ww,r,nm,sorted_u(src.get(col,{'G':'experience_level','I':'education_level','F':'remote_type'}.get(col))),col)
    chart(ww,'bar',f'Avg salary by {nm.lower()}',r1,r2,4,f"G{r}",w=14,h=6.5,lo=0)
    r=max(r2+3,r+14)
# years of experience
hdr(ww,r,['Years experience','Postings','Total openings','Avg salary'])
for y in range(15):
    q=r+1+y;ww.cell(q,1,y);ww.cell(q,2,f"=COUNTIFS({rng('H')},$A{q})");ww.cell(q,3,f"=SUMIFS({rng('S')},{rng('H')},$A{q})");ww.cell(q,4,f"=AVERAGEIFS({rng('O')},{rng('H')},$A{q})")
    for c in (2,3,4): ww.cell(q,c).number_format='#,##0'
chart(ww,'line','Avg salary by years of experience',r+1,r+15,4,f"G{r}",w=14,h=6.5,lo=0)
r+=19
ww.cell(r,1,'Avg salary: job title x experience level').font=F(bold=True)
levels=['Entry','Mid','Senior'];hdr(ww,r+1,['Job title']+levels)
for i,t in enumerate(sorted_u('job_title')):
    q=r+2+i;ww.cell(q,1,t)
    for j,l in enumerate(levels,2):
        ww.cell(q,j,f"=AVERAGEIFS({rng('O')},{rng('B')},$A{q},{rng('G')},{L(j)}${r+1})").number_format='#,##0'
widths(ww,[24,12,15,13,3,3,10])
# SKILLS
sk=['skills_python','skills_sql','skills_ml','skills_deep_learning','skills_cloud'];nm={'skills_python':'Python','skills_sql':'SQL','skills_ml':'Machine Learning','skills_deep_learning':'Deep Learning','skills_cloud':'Cloud'}
col={s:L(10+i) for i,s in enumerate(sk)}
ws_=wb.create_sheet('Skill Trends');title(ws_,'Skill demand trend analysis','Prevalence = share of postings whose skill flag = 1.')
hdr(ws_,4,['Posting year']+[nm[s] for s in sk])
years=sorted_u('job_posting_year')
for i,y in enumerate(years,5):
    ws_.cell(i,1,int(y))
    for j,s in enumerate(sk,2): ws_.cell(i,j,f"=AVERAGEIFS({rng(col[s])},{rng('Q')},$A{i})").number_format='0.0%'
lc=LineChart();lc.title='Skill prevalence by posting year';lc.height=8;lc.width=18
lc.add_data(Reference(ws_,min_col=2,max_col=6,min_row=4,max_row=4+len(years)),titles_from_data=True);lc.set_categories(Reference(ws_,min_col=1,min_row=5,max_row=4+len(years)))
lc.y_axis.scaling.min=0.4;lc.y_axis.scaling.max=0.6;lc.y_axis.number_format='0%';lc.y_axis.delete=False;lc.x_axis.delete=False;ws_.add_chart(lc,'H4')
r=15;hdr(ws_,r,['Skill','Avg salary WITH','Avg salary WITHOUT','Premium (unadjusted)','Overall prevalence'])
for i,s in enumerate(sk,1):
    q=r+i;ws_.cell(q,1,nm[s]);ws_.cell(q,2,f"=AVERAGEIFS({rng('O')},{rng(col[s])},1)");ws_.cell(q,3,f"=AVERAGEIFS({rng('O')},{rng(col[s])},0)");ws_.cell(q,4,f"=B{q}-C{q}");ws_.cell(q,5,f"=AVERAGE({rng(col[s])})")
    for c in (2,3,4): ws_.cell(q,c).number_format='#,##0;-#,##0'
    ws_.cell(q,5).number_format='0.0%'
chart(ws_,'bar','Salary premium when skill required',r+1,r+5,4,'H21',w=18,h=7.5)
r=23;hdr(ws_,r,['Job title']+[nm[s] for s in sk])
for i,t in enumerate(sorted_u('job_title'),1):
    q=r+i;ws_.cell(q,1,t)
    for j,s in enumerate(sk,2): ws_.cell(q,j,f"=AVERAGEIFS({rng(col[s])},{rng('B')},$A{q})").number_format='0.0%'
r=32;hdr(ws_,r,['Skills required (0-5)','Postings','Avg salary'])
for k in range(6):
    q=r+1+k;ws_.cell(q,1,k);ws_.cell(q,2,f"=COUNTIFS({rng('T')},$A{q})");ws_.cell(q,3,f"=IFERROR(AVERAGEIFS({rng('O')},{rng('T')},$A{q}),0)")
    ws_.cell(q,2).number_format=ws_.cell(q,3).number_format='#,##0'
widths(ws_,[24,16,18,20,18,10])
# SALARY MODELS (static from Python)
wm=wb.create_sheet('Salary Models');title(wm,'Salary prediction models','Values below were computed in Python (scikit-learn): 80/20 split, random_state=42, 5-fold CV. They are static results, not formulas.')
hdr(wm,4,['Model','Test R2','MAE','RMSE','CV R2 mean','CV R2 sd'])
for i,m in enumerate(o['sal_models'],5):
    for j,k in enumerate(['model','r2','mae','rmse','cv_r2','cv_sd'],1): wm.cell(i,j,m[k])
    wm.cell(i,3).number_format=wm.cell(i,4).number_format='#,##0'
hdr(wm,11,['Feature (grouped)','Permutation importance (drop in test R2)'])
imp=[(k,v) for k,v in o['sal_imp'].items()];
for i,(k,v) in enumerate(imp,12): wm.cell(i,1,k);wm.cell(i,2,max(v,0)).number_format='0.000'
chart(wm,'bar','Feature importance (Gradient Boosting)',12,11+len(imp),2,'H4',w=20,h=9)
wm.cell(30,1,'Reading: salary is driven by experience level and job title, then company size and ML / Deep Learning / Cloud flags. Other features add ~0. Linear Regression matches tree models, so the pattern is close to additive.').font=F(italic=True)
widths(wm,[28,22,10,10,12,10])
# CLASSIFICATION
wc=wb.create_sheet('ML Classification');title(wc,'Machine learning classification tasks','Static Python results; stratified 80/20 split and 5-fold CV. Majority baseline = always predict the most common class.')
r=4
for c in o['clf']:
    wc.cell(r,1,c['name']+'  |  '+', '.join(f"{k}: {v:,}" for k,v in c['classes'].items())).font=F(bold=True)
    hdr(wc,r+1,['Model','Test accuracy','Macro F1','CV accuracy'])
    for i,m in enumerate(c['results'],r+2):
        for j,k in enumerate(['model','acc','f1','cv_acc'],1): wc.cell(i,j,float(m[k]) if j>1 else m[k]).number_format='0.0%' if j>1 else 'General'
    chart(wc,'bar',c['name'],r+2,r+5,2,f"G{r}",w=12,h=5.5,lo=0)
    r+=13
wc.cell(r,1,'Hiring urgency is a deterministic rule in this file: 7+ openings = High; otherwise High if ML and Deep Learning both required, else Low (<=7 yrs experience) / Medium (>7). 100% tree accuracy reflects that rule, not general hiring insight. Remote type is not predictable (~chance).').font=F(italic=True)
widths(wc,[28,14,12,12])
# SUMMARY
s=wb.create_sheet('Summary',0);title(s,'AI Job Market Dashboard - Summary','Source: uploaded Excel, CLEANED Dataset sheet (identical to RAW). No external data or assumptions added.')
k=[('Postings',f"=COUNTA({rng('A')})",'#,##0'),('Total openings',f"=SUM({rng('S')})",'#,##0'),('Mean salary',f"=AVERAGE({rng('O')})",'#,##0'),('Median salary',f"=MEDIAN({rng('O')})",'#,##0'),('Min salary',f"=MIN({rng('O')})",'#,##0'),('Max salary',f"=MAX({rng('O')})",'#,##0'),('First posting year',f"=MIN({rng('Q')})",'0'),('Last posting year',f"=MAX({rng('Q')})",'0')]
hdr(s,4,['KPI','Value'])
for i,(a,f,n) in enumerate(k,5): s.cell(i,1,a);s.cell(i,2,f).number_format=n
s['A15']='Key findings';s['A15'].font=F(bold=True,size=12,color='1F3A5F')
fi=["Salary prediction: Linear Regression, Random Forest and Gradient Boosting all reach test R2 ~0.99 (MAE ~2,450) vs ~0 for a mean baseline.",
"Salary drivers: experience level and job title dominate; company size, ML, Deep Learning and Cloud flags follow. Country, industry, education, remote type and years of experience add ~nothing.",
"Pay gaps: Entry ~89k, Mid ~114k, Senior ~138k; AI Engineer and ML Engineer ~140k vs ~100k for the other four titles; multinational corporations ~20k above other company sizes.",
"Job demand: postings and openings are spread very evenly across titles, industries, countries and years (see Job Demand sheet).",
"Classification: hiring urgency follows a deterministic rule (tree models 100%); experience level ~93-99%; salary above median ~96-97%; remote type ~chance.",
"Skills: each skill is required in ~47-54% of postings every year with no sustained trend; ML, Deep Learning and Cloud carry a salary premium (~14.6k, ~15k, ~9.7k), Python and SQL do not.",
"Caveats: salary currency/period not stated; 2026 postings included as recorded; very uniform patterns suggest the file may be synthetic - avoid presenting as real labour-market findings."]
for i,t in enumerate(fi,16): s.cell(i,1,t).alignment=Alignment(wrap_text=True,vertical='top');s.merge_cells(start_row=i,start_column=1,end_row=i,end_column=6);s.row_dimensions[i].height=32
s['A24']='Sheets: Job Demand, Workforce Analytics, Skill Trends (live formulas) | Salary Models, ML Classification (static Python results) | Data (cleaned dataset + n_skills helper column).'
s['A24'].font=F(italic=True,color='666666')
widths(s,[34,16,16,16,16,16])
for w in wb.worksheets:
    if w.title!='Data': style(w)
    w.sheet_view.showGridLines=False if w.title!='Data' else True
wb.save('outputs/AI_Job_Market_Dashboard.xlsx')
