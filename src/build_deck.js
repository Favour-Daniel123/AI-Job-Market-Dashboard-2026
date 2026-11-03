const pptxgen=require('pptxgenjs');const o=require('../results/results.json');
const p=new pptxgen();p.layout='LAYOUT_WIDE';p.title='AI Job Market Analysis';
const NAVY='1F3A5F',TEAL='12B886',AMB='F59F00',INK='1B2030',MU='5B6478',BG='F4F6FA',W='FFFFFF';
const HF='Cambria',BF='Calibri';
const f0=n=>Math.round(n).toLocaleString('en-US');
function title(s,t,sub){s.background={color:W};s.addText(t,{x:0.6,y:0.35,w:12.1,h:0.8,fontFace:HF,fontSize:32,bold:true,color:NAVY,margin:0,isTextBox:true});
 if(sub)s.addText(sub,{x:0.6,y:1.12,w:12.1,h:0.4,fontFace:BF,fontSize:14,color:MU,margin:0,isTextBox:true});}
function bullets(s,arr,x,y,w,h){s.addText(arr.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<arr.length-1,paraSpaceAfter:10}})),{x,y,w,h,fontFace:BF,fontSize:15,color:INK,valign:'top',margin:0,isTextBox:true});}
const ax={catAxisLabelFontFace:BF,valAxisLabelFontFace:BF,catAxisLabelFontSize:11,valAxisLabelFontSize:11,catAxisLabelColor:MU,valAxisLabelColor:MU,valGridLine:{color:'E3E6EF',size:0.5},catGridLine:{style:'none'},dataLabelFontFace:BF,dataLabelFontSize:11,dataLabelColor:INK,titleFontFace:BF,titleFontSize:13,titleColor:INK,showTitle:true,showLegend:false};
// 1 title
let s=p.addSlide();s.background={color:NAVY};
s.addText('AI Job Market Analysis',{x:0.8,y:2.2,w:11.5,h:1.2,fontFace:HF,fontSize:48,bold:true,color:W,margin:0,isTextBox:true});
s.addText('Salary prediction · Job demand · ML classification · Workforce analytics · Skill trends',{x:0.8,y:3.5,w:11.5,h:0.5,fontFace:BF,fontSize:18,color:'CADCFC',margin:0,isTextBox:true});
s.addText('10,345 job postings, 2020–2026 (source: uploaded Excel, CLEANED Dataset sheet)',{x:0.8,y:4.2,w:11.5,h:0.4,fontFace:BF,fontSize:14,color:'9FB4D6',margin:0,isTextBox:true});
// 2 overview
s=p.addSlide();title(s,'The dataset at a glance','RAW and CLEANED sheets are identical: no missing values, no duplicates');
[['10,345','job postings'],['51,767','total openings'],['113,438','mean salary (as recorded)'],['2020–2026','posting years']].forEach((k,i)=>{
 const x=0.6+i*3.1;s.addShape(p.ShapeType.roundRect,{x,y:1.9,w:2.8,h:1.8,fill:{color:BG},line:{color:BG},rectRadius:0.1});
 s.addText(k[0],{x,y:2.1,w:2.8,h:0.8,fontFace:HF,fontSize:34,bold:true,color:NAVY,align:'center',margin:0,isTextBox:true});
 s.addText(k[1],{x,y:2.95,w:2.8,h:0.5,fontFace:BF,fontSize:14,color:MU,align:'center',margin:0,isTextBox:true});});
bullets(s,['6 job titles, 7 countries, 6 industries, 4 company sizes, 3 experience levels','Salary currency and period are not stated in the file, so values are shown as recorded','Postings dated 2026 are included exactly as they appear'],0.6,4.1,12,2.4);
// 3 salary
s=p.addSlide();title(s,'Salary is predictable: R² ≈ 0.99','Three models, 80/20 hold-out and 5-fold cross-validation');
const imp=Object.entries(o.sal_imp).filter(e=>e[1]>0.001);
s.addChart(p.charts.BAR,[{name:'Importance',labels:imp.map(e=>e[0].replace('skills_','skill: ').replace(/_/g,' ')),values:imp.map(e=>e[1])}],{x:0.6,y:1.7,w:6.6,h:4.9,barDir:'bar',chartColors:[NAVY],showValue:true,dataLabelFormatCode:'0.00',dataLabelPosition:'outEnd',title:'What drives salary (drop in R² when shuffled)',catAxisOrientation:'maxMin',...ax});
const hd=['Model','R²','MAE'].map(t=>({text:t,options:{bold:true,color:W,fill:{color:NAVY},fontFace:BF,fontSize:13}}));
s.addTable([hd,...o.sal_models.map(m=>[m.model,m.r2.toFixed(3),f0(m.mae)].map(t=>({text:t,options:{fontFace:BF,fontSize:13,color:INK}})))],{x:7.6,y:1.8,w:5.1,colW:[2.7,1.1,1.3],border:{type:'solid',color:'E3E6EF',pt:0.5}});
bullets(s,['Experience level and job title dominate; company size and ML / Deep Learning / Cloud flags follow','Country, industry, education, remote type and years of experience add almost nothing','Linear Regression matches tree models, so the relationship is close to additive'],7.6,3.9,5.1,2.7);
// 4 demand
s=p.addSlide();title(s,'Demand is evenly spread','Total openings (sum of job_openings) across the dataset');
const T=o.by_job_title;
s.addChart(p.charts.BAR,[{name:'Openings',labels:T.map(t=>t.job_title),values:T.map(t=>t.openings)}],{x:0.6,y:1.7,w:7,h:4.9,barDir:'bar',chartColors:[TEAL],valAxisMinVal:0,showValue:true,dataLabelPosition:'outEnd',dataLabelFormatCode:'#,##0',title:'Openings by job title',catAxisOrientation:'maxMin',...ax});
const Y=o.by_job_posting_year;
s.addChart(p.charts.LINE,[{name:'Openings',labels:Y.map(y=>String(y.job_posting_year)),values:Y.map(y=>y.openings)}],{x:7.9,y:1.7,w:4.8,h:3,chartColors:[NAVY],lineSize:2,valAxisMinVal:0,title:'Openings by year',...ax});
bullets(s,['Titles, industries, countries and years each hold a similar share','High urgency: 50.2% of postings'],7.9,4.9,4.8,1.7);
// 5 classification
s=p.addSlide();title(s,'Classification: strong where a rule exists','Best model test accuracy vs always predicting the majority class');
const C=o.clf,nm=['Hiring urgency','Experience level','Remote type','Salary > median'];
const base=C.map(c=>c.results[0].acc),best=C.map(c=>Math.max(...c.results.slice(1).map(r=>r.acc)));
s.addChart(p.charts.BAR,[{name:'Majority baseline',labels:nm,values:base},{name:'Best model',labels:nm,values:best}],{x:0.6,y:1.7,w:7.2,h:4.9,barDir:'col',chartColors:['B8C0D4',NAVY],showValue:true,dataLabelFormatCode:'0%',dataLabelPosition:'outEnd',valAxisMaxVal:1.1,valAxisMinVal:0,valAxisLabelFormatCode:'0%',showLegend:true,legendPos:'b',legendFontFace:BF,legendFontSize:12,...Object.assign({},ax,{showLegend:true}),title:'Test accuracy',});
bullets(s,['Hiring urgency follows a deterministic rule in the data (7+ openings, or ML + Deep Learning required, else years of experience), so trees hit 100%','Experience level (years excluded) and above-median salary reach ~99% and ~97%','Remote type is at chance: no signal in the available columns'],8.1,1.8,4.6,4.8);
// 6 workforce
s=p.addSlide();title(s,'Experience level and company size set pay','Average salary, as recorded');
const E=o.by_experience_level,Z=o.by_company_size;
s.addChart(p.charts.BAR,[{name:'Avg salary',labels:['Entry','Mid','Senior'],values:['Entry','Mid','Senior'].map(l=>Math.round(E.find(e=>e.experience_level==l).avg_sal))}],{x:0.6,y:1.7,w:5.9,h:4.0,barDir:'col',chartColors:[NAVY],showValue:true,dataLabelFormatCode:'#,##0',dataLabelPosition:'outEnd',valAxisMinVal:0,title:'By experience level',...ax});
s.addChart(p.charts.BAR,[{name:'Avg salary',labels:Z.map(z=>z.company_size.replace('MultiNational Corporations','Multinational')),values:Z.map(z=>Math.round(z.avg_sal))}],{x:6.8,y:1.7,w:5.9,h:4.0,barDir:'col',chartColors:[TEAL],showValue:true,dataLabelFormatCode:'#,##0',dataLabelPosition:'outEnd',valAxisMinVal:0,title:'By company size',...ax});
s.addText('AI Engineer and ML Engineer average ~140k vs ~100k for the other four titles. Country, industry, education and remote type differ by only a few thousand.',{x:0.6,y:5.9,w:12.1,h:0.8,fontFace:BF,fontSize:15,color:INK,margin:0,isTextBox:true});
// 7 skills
s=p.addSlide();title(s,'Skills: stable demand, uneven pay','Each skill appears in ~47–54% of postings in every year, with no sustained trend');
const K=[['Python','skills_python'],['SQL','skills_sql'],['Machine Learning','skills_ml'],['Deep Learning','skills_deep_learning'],['Cloud','skills_cloud']];
s.addChart(p.charts.BAR,[{name:'Premium',labels:K.map(k=>k[0]),values:K.map(k=>o.skill_prem[k[1]][2])}],{x:0.6,y:1.7,w:7,h:4.9,barDir:'col',chartColors:[AMB],showValue:true,dataLabelFormatCode:'#,##0',dataLabelPosition:'outEnd',title:'Salary premium when skill is required (unadjusted)',...ax});
bullets(s,['Deep Learning (+15.2k), Machine Learning (+14.6k) and Cloud (+9.7k) postings pay more','Python and SQL show no meaningful gap','Prevalence stays within a narrow band every year'],7.9,1.9,4.8,4.5);
// 8 caveats
s=p.addSlide();s.background={color:NAVY};
s.addText('Takeaways and caveats',{x:0.8,y:0.6,w:11.5,h:0.9,fontFace:HF,fontSize:36,bold:true,color:W,margin:0,isTextBox:true});
s.addText([['Pay is set by experience level, job title and company size; ML, Deep Learning and Cloud skills add a premium.'],['Demand and skill prevalence are flat across years, industries and countries.'],['Hiring urgency is a rule baked into the file; remote type is unpredictable.'],['Currency and period of salary are unstated, and 2026 postings are included as recorded.'],['The patterns are very uniform, which suggests synthetic data: avoid presenting these as real labour-market findings.']].map((t,i,a)=>({text:t[0],options:{bullet:true,breakLine:i<a.length-1,paraSpaceAfter:12}})),{x:0.8,y:1.8,w:11.7,h:4.8,fontFace:BF,fontSize:18,color:W,valign:'top',margin:0,isTextBox:true});
p.writeFile({fileName:'outputs/AI_Job_Market_Analysis.pptx'}).then(()=>console.log('ok'));
