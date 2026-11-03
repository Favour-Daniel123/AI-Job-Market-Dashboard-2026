import json
o=json.load(open('results/results.json'))
sk=['skills_python','skills_sql','skills_ml','skills_deep_learning','skills_cloud']
D={
 'n':o['n'],'open':o['openings_total'],
 'title':o['by_job_title'],'ind':o['by_company_industry'],'ctry':o['by_country'],'size':o['by_company_size'],
 'rem':o['by_remote_type'],'exp':o['by_experience_level'],'edu':o['by_education_level'],'urg':o['by_hiring_urgency'],
 'yr':o['by_job_posting_year'],'mon':o['by_month'],'ty':o['title_x_year_openings'],
 'sm':o['sal_models'],'imp':o['sal_imp'],'avp':o['avp'][:250],
 'clf':[{'name':c['name'],'classes':c['classes'],'res':c['results']} for c in o['clf']],
 'sy':o['skill_year'],'st':o['skill_title'],'sp':o['skill_prem'],'ns':o['nskills_sal'],'ys':o['yrs_sal'],
 'ets':o['exp_title_sal'],
}
html=open('src/dashboard_template.html').read().replace('__DATA__',json.dumps(D,separators=(',',':')))
open('outputs/ai_job_market_dashboard.html','w').write(html)
print(len(html))
