
# coding: utf-8

# In[66]:

import datetime
import webbrowser
now = datetime.datetime.now()


# In[54]:

file = open("C:\\Users\\Splinta\\GoogleDrive\\GitStuff\\PortalUpdate\\bi_state.txt","r")
filedata = file.read()


# In[56]:

# Replace the target string
filedata = filedata.replace('xDay', now.strftime("%A"))
filedata = filedata.replace('xMonth', now.strftime("%B"))
filedata = filedata.replace('xNum', str(now.day))
filedata = filedata.replace('xYear', str(now.year))


# In[65]:

# Write the file out again
file = open("C:\\Users\\Splinta\\GoogleDrive\\GitStuff\\PortalUpdate\\u_bi_state.txt", "w")
file.write(filedata)
file.close()


# In[67]:

webbrowser.open('http://mlfess.mapleleaf.ca/irj/portal')


# In[64]:




# In[ ]:



