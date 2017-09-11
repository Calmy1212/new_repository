# -*- coding: utf-8 -*-
"""
Created on Mon Aug 28 11:44:39 2017

@author: Quan.Li
"""

from sqlalchemy import create_engine
import pandas as pd

engine = create_engine('mysql+pymysql://ExternalPage:webaccess@75.96.219.150:3306/go_external_devtest')
company = pd.read_sql_query("SELECT DISTINCT cc.Name AS company_name, ccs.Name AS mm_status, au.first_name AS pm_first_name,au.last_name AS pm_last_name, concat('(',gvc.Code,') ',gvc.Name) AS gv_company,cc.KYCCompletedDate, cco.Name,cc.DecisionMakers, cc.JoinedDate,csd.Company_id, csd.Companycategory_id, ccc.Level,if(ccc.Level=0, ccc.value, 0) as mm_segment,if(ccc.Level=1, ccc.value,0) as level_2,if(ccc.Level=2, ccc.value, 0) as level_3 FROM go_external_devtest.contacts_company cc INNER JOIN go_external_devtest.contacts_companystatus AS ccs ON cc.Status_id= ccs.id INNER JOIN go_external_devtest.contacts_profile AS cp ON cp.id=cc.ManagerPrimary_id INNER JOIN go_external_devtest.auth_user AS au ON au.id=cp.user_id INNER JOIN go_external_devtest.contacts_country AS cco ON cco.iso=cc.ReportCountry_id INNER JOIN go_external_devtest.globalvision_gvcompany AS gvc ON gvc.GVCompanyID = cc.GVCompany_id LEFT JOIN go_external_devtest.contacts_companysupplementarydata AS csd ON cc.id = csd.Company_id INNER JOIN go_external_devtest.contacts_companycategory ccc ON ccc.id = csd.Companycategory_id WHERE cc.Status_id=1 AND csd.TagIdentifier='CATEGORY';", engine)

df = pd.DataFrame()
company = company.groupby('company_name')
for name, group in company:
    mm_segment = ""
    level_2 = ""
    level_3 = ""
    for index, row in group.iterrows():
        if row["mm_segment"] != "0":
            mm_segment = row["mm_segment"]
        if row["level_2"] !="0":
            level_2 = row["level_2"]
        if row["level_3"] != "0":
            level_3 = row["level_3"]
    new_row = group.iloc[[0]]
    new_row['mm_segment'] = mm_segment
    new_row['level_2'] = level_2
    new_row['level_3'] = level_3
    df = df.append(new_row)        

writer = pd.ExcelWriter('C:/Users/Quan.Li/Desktop/Jiaqi/company.xlsx')
df.to_excel(writer,"Sheet1",index=False)