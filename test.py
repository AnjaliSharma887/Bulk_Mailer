import pandas as pd
import numpy as np
data=pd.read_excel("email_excel_file.xlsx")
#print(data['email'])

if 'email' in data.columns:
    emails=list(data["email"])
    c=[]
    for i in emails:
         if pd.isnull(i)==False:
            c.append(i)
    emails=np.array(c)
    print(emails)
    print(type(emails))
else:
    print("Not Exist")

