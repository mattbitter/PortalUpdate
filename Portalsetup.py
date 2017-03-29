
# coding: utf-8

# In[4]:

from distutils.core import setup
import py2exe

setup(console=['PortalUpdate.py'],
     options = {
         'py2exe': { 'packages' : ['datetime','webbrowser']
                   }
             }
     )


# In[ ]:



