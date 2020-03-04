#!/usr/bin/env python
# coding: utf-8

# In[1]:


import win32com.client
import os
import pandas as pd
from datetime import datetime


# In[2]:


outlook=win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")


# In[3]:


#inbox = outlook.GetDefaultFolder(6)
root_folder = outlook.Folders.Item(1)
subfolder = root_folder.Folders['RPA']
messages = subfolder.Items


# In[4]:


#message=inbox.Items
print(messages.count)


# In[13]:


a = pd.read_excel(r'C:\Users\praseena.s\Desktop\PES\Redmine data\Project.xlsx')


# In[14]:


final=pd.read_excel(r"C:\Users\praseena.s\Desktop\PES\Invoice track\final.xls")


# In[15]:


result = pd.concat([a,final], ignore_index=True, sort=False, join = 'inner')


# In[16]:


final=final.append(result)


# In[19]:


final.to_excel(r'C:\Users\praseena.s\Desktop\PES\Invoice track\final.xls')


# In[20]:


#final=pd.read_excel(r"C:\Users\praseena.s\Desktop\PES\Invoice track\final.xls")
f=final[final.columns[2]]


# In[21]:


f


# In[22]:


for message in messages:
    body=message.body
    subject=message.subject
    #print(subject)
    #print(body)
    for i in f.index:
        match = body.find(f.loc[i])
        match2=subject.find(f.loc[i])
        if match is not -1:
            date=message.senton.date()   
            #print(date)
            final.loc[i,'PM Submitted date'] = date.strftime("%d/%m/%Y")
            final.update(final)
        if match2 is not -1:
            date2=message.senton.date()
            #print(date2)
            final.loc[i,'BDM Submitted Date'] = date2.strftime("%d/%m/%Y")
            final.update(final)
    final.to_excel('C:\Users\praseena.s\Desktop\PES\Invoice track\final.xls',index=False) 


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




