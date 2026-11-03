import pandas as pd, numpy as np, json
from sklearn.model_selection import train_test_split, cross_val_score, StratifiedKFold, KFold
from sklearn.linear_model import LinearRegression, LogisticRegression
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor, RandomForestClassifier, GradientBoostingClassifier
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error, accuracy_score, f1_score, confusion_matrix
from sklearn.dummy import DummyClassifier, DummyRegressor
from sklearn.inspection import permutation_importance
df=pd.read_excel('data/1--AI_Job_Market_Dataset.xlsx',sheet_name=1)
sk=['skills_python','skills_sql','skills_ml','skills_deep_learning','skills_cloud']
cats=['job_title','company_size','company_industry','country','remote_type','experience_level','education_level','hiring_urgency']
out={}
# demand
out['n']=len(df); out['openings_total']=int(df.job_openings.sum())
def g(col): 
    t=df.groupby(col).agg(postings=('job_id','count'),openings=('job_openings','sum'),avg_sal=('salary','mean'),med_sal=('salary','median')).round(0).reset_index()
    return t.to_dict('records')
for c in ['job_title','company_industry','country','company_size','remote_type','experience_level','education_level','hiring_urgency','job_posting_year']:
    out['by_'+c]=g(c)
out['by_month']=df.groupby('job_posting_month').agg(postings=('job_id','count'),openings=('job_openings','sum')).reset_index().to_dict('records')
out['title_x_year_openings']=df.pivot_table(index='job_posting_year',columns='job_title',values='job_openings',aggfunc='sum').reset_index().to_dict('records')
out['title_x_urg']=pd.crosstab(df.job_title,df.hiring_urgency).reset_index().to_dict('records')
out['ind_x_title']=pd.crosstab(df.company_industry,df.job_title).reset_index().to_dict('records')
# skills trends
out['skill_year']=df.groupby('job_posting_year')[sk].mean().mul(100).round(1).reset_index().to_dict('records')
out['skill_title']=df.groupby('job_title')[sk].mean().mul(100).round(1).reset_index().to_dict('records')
out['skill_overall']=df[sk].mean().mul(100).round(1).to_dict()
prem={}
for s in sk:
    a=df[df[s]==1].salary.mean(); b=df[df[s]==0].salary.mean(); prem[s]=[round(a),round(b),round(a-b)]
out['skill_prem']=prem
df['nskills']=df[sk].sum(1)
out['nskills_sal']=df.groupby('nskills').salary.agg(['count','mean']).round(0).reset_index().to_dict('records')
# workforce
out['yrs_sal']=df.groupby('years_experience').salary.mean().round(0).reset_index().to_dict('records')
out['exp_yrs']=df.groupby('experience_level').years_experience.agg(['mean','min','max']).round(1).reset_index().to_dict('records')
out['exp_title_sal']=df.pivot_table(index='job_title',columns='experience_level',values='salary',aggfunc='mean').round(0).reset_index().to_dict('records')
out['country_title_sal']=df.pivot_table(index='country',columns='job_title',values='salary',aggfunc='mean').round(0).reset_index().to_dict('records')
out['sal_year']=df.groupby('job_posting_year').salary.agg(['mean','median']).round(0).reset_index().to_dict('records')
out['sal_hist']=np.histogram(df.salary,bins=np.arange(40000,210001,10000))
out['sal_hist']=[out['sal_hist'][0].tolist(),out['sal_hist'][1].tolist()]
out['corr']=df[['salary','years_experience','job_openings','nskills']+sk].corr().round(2).to_dict()
# salary models
feat=['job_title','company_size','company_industry','country','remote_type','experience_level','education_level','years_experience','job_posting_year','job_posting_month','hiring_urgency','job_openings']+sk
X=pd.get_dummies(df[feat],drop_first=False).astype(float); y=df.salary
Xtr,Xte,ytr,yte=train_test_split(X,y,test_size=.2,random_state=42)
models={'Mean baseline':DummyRegressor(),'Linear Regression':LinearRegression(),'Random Forest':RandomForestRegressor(300,min_samples_leaf=3,n_jobs=-1,random_state=42),'Gradient Boosting':GradientBoostingRegressor(random_state=42)}
res=[]
kf=KFold(5,shuffle=True,random_state=42)
for n,m in models.items():
    m.fit(Xtr,ytr); p=m.predict(Xte)
    cv=cross_val_score(m,X,y,cv=kf,scoring='r2',n_jobs=-1)
    res.append(dict(model=n,r2=round(r2_score(yte,p),3),mae=round(mean_absolute_error(yte,p)),rmse=round(mean_squared_error(yte,p)**.5),cv_r2=round(cv.mean(),3),cv_sd=round(cv.std(),3)))
    models[n]=(m,p)
out['sal_models']=res
best=max(res[1:],key=lambda r:r['r2'])['model']; out['best_sal']=best
m,p=models[best]
out['avp']=[[int(a),int(b)] for a,b in zip(yte.values[:300],p[:300])]
# grouped permutation importance on original features
rf=models['Gradient Boosting'][0] if best=='Gradient Boosting' else models[best][0]
imp={}
Xte_=Xte.copy()
rng=np.random.RandomState(0)
base=r2_score(yte,rf.predict(Xte))
for f in feat:
    cols=[c for c in X.columns if c==f or c.startswith(f+'_')]
    drops=[]
    for _ in range(5):
        Xp=Xte.copy(); idx=rng.permutation(len(Xp)); Xp[cols]=Xte[cols].values[idx]
        drops.append(base-r2_score(yte,rf.predict(Xp)))
    imp[f]=round(float(np.mean(drops)),4)
out['sal_imp']=dict(sorted(imp.items(),key=lambda kv:-kv[1]))
if best=='Linear Regression':
    lr=models[best][0]; out['lin_coef']=dict(zip(X.columns,lr.coef_.round(0)))
# classification
def clf(target,drop,name,balanced=False):
    feats=[f for f in feat+['salary'] if f not in drop and f!=target]
    Xc=pd.get_dummies(df[feats],drop_first=False).astype(float); yc=df[target]
    Xtr,Xte,ytr,yte=train_test_split(Xc,yc,test_size=.2,random_state=42,stratify=yc)
    ms={'Majority baseline':DummyClassifier(strategy='most_frequent'),'Logistic Regression':LogisticRegression(max_iter=3000),'Random Forest':RandomForestClassifier(300,min_samples_leaf=3,n_jobs=-1,random_state=42),'Gradient Boosting':GradientBoostingClassifier(random_state=42)}
    r=[]; cms={}
    skf=StratifiedKFold(5,shuffle=True,random_state=42)
    for n,m in ms.items():
        m.fit(Xtr,ytr); p=m.predict(Xte)
        cv=cross_val_score(m,Xc,yc,cv=skf,scoring='accuracy',n_jobs=-1)
        r.append(dict(model=n,acc=round(accuracy_score(yte,p),3),f1=round(f1_score(yte,p,average='macro'),3),cv_acc=round(cv.mean(),3)))
        labs=sorted(yc.unique()); cms[n]=dict(labels=labs,m=confusion_matrix(yte,p,labels=labs).tolist())
    return dict(name=name,target=target,classes=yc.value_counts().to_dict(),results=r,cms=cms,n_features=len(feats))
out['clf']=[clf('hiring_urgency',[],'Hiring urgency (High/Medium/Low)'),
            clf('experience_level',['years_experience'],'Experience level (Entry/Mid/Senior)'),
            clf('remote_type',[],'Remote type (Remote/Hybrid/Onsite)')]
df['high_salary']=np.where(df.salary>df.salary.median(),'Above median','At/below median')
out['clf'].append(clf('high_salary',['salary'],'Salary above median (binary)') if False else None)
out['clf']=[c for c in out['clf'] if c]
# high-salary: needs salary excluded
feats=[f for f in feat]
Xc=pd.get_dummies(df[feats],drop_first=False).astype(float); yc=df.high_salary
Xtr,Xte,ytr,yte=train_test_split(Xc,yc,test_size=.2,random_state=42,stratify=yc)
ms={'Majority baseline':DummyClassifier(strategy='most_frequent'),'Logistic Regression':LogisticRegression(max_iter=5000),'Random Forest':RandomForestClassifier(300,min_samples_leaf=3,n_jobs=-1,random_state=42),'Gradient Boosting':GradientBoostingClassifier(random_state=42)}
r=[];cms={}
for n,m in ms.items():
    m.fit(Xtr,ytr);p=m.predict(Xte)
    cv=cross_val_score(m,Xc,yc,cv=StratifiedKFold(5,shuffle=True,random_state=42),scoring='accuracy',n_jobs=-1)
    r.append(dict(model=n,acc=round(accuracy_score(yte,p),3),f1=round(f1_score(yte,p,average='macro'),3),cv_acc=round(cv.mean(),3)))
    labs=sorted(yc.unique());cms[n]=dict(labels=labs,m=confusion_matrix(yte,p,labels=labs).tolist())
out['clf'].append(dict(name='Salary above median (binary)',target='salary > median ($%d)'%int(df.salary.median()),classes=yc.value_counts().to_dict(),results=r,cms=cms,n_features=len(feats)))
json.dump(out,open('results/results.json','w'),default=lambda o:o.item() if hasattr(o,'item') else str(o))
print(json.dumps(out['sal_models'],indent=0)); print(out['sal_imp']); print(out['skill_prem']); print(out['skill_overall'])
for c in out['clf']: print(c['name'],c['classes']); [print(' ',x) for x in c['results']]
print(out['sal_year']); print(out['by_job_title']); print(out['skill_year'])
