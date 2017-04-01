
# coding: utf-8

# In[15]:

import datetime
import webbrowser
now = datetime.datetime.now()

import win32api, win32con
import keyboard
import time

from O365 import Message
import pandas as pd

from PortalUpdateConfig import SAPun, SAPpw, OFFICEun, OFFICEpw, chrome_path


# In[87]:

xl = pd.ExcelFile("\\\camis2-mlf-fp01\\Project Leapfrog\\9. Business Analytics\\NWPortalUpdateFile\\VeryNewReconAllBU.xlsx")


# In[88]:

df = xl.parse('AllBU')


# In[89]:

def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)


# In[90]:

def compare():
    file = open("\\\camis2-mlf-fp01\\Project Leapfrog\\9. Business Analytics\\NWPortalUpdateFile\\bi_state.txt","r")
    filedata = file.read()
    
    # Replace the target string
    filedata = filedata.replace('xDay', now.strftime("%A"))
    filedata = filedata.replace('xMonth', now.strftime("%B"))
    filedata = filedata.replace('xNum', str(now.day))
    filedata = filedata.replace('xYear', str(now.year))
    
    # Write the file out again
    file = open("\\\camis2-mlf-fp01\\Project Leapfrog\\9. Business Analytics\\NWPortalUpdateFile\\bi_state.html", "w")
    file.write(filedata)
    file.close()

    webbrowser.get(chrome_path).open_new_tab('http://caeppap01.mapleleaf.ca:53900/irj/go/km/navigation/documents/MLF/pages?StartUri=/')

    time.sleep(2)

    
    
    keyboard.write(SAPun)
    keyboard.press_and_release('tab')
    time.sleep(1)
    keyboard.write(SAPpw)

    time.sleep(1)
    keyboard.press_and_release('enter')
    time.sleep(2)

    click(201,75)
    time.sleep(1)

    click(201,92)
    time.sleep(1)
    click(252,113)
    time.sleep(1)

    time.sleep(1)
    click(109,204)

    time.sleep(1)
    #click(788,205)
    keyboard.press_and_release('Alt+d')
    time.sleep(1)

    text_path = str('\\\camis2-mlf-fp01\\Project Leapfrog\\9. Business Analytics\\NWPortalUpdateFile')

    keyboard.write(text_path)
    keyboard.press_and_release('enter')

    time.sleep(1)
    keyboard.press_and_release('Alt+n')
    #click(408,632)

    keyboard.write('bi_state.html')
    keyboard.press_and_release('enter')

    time.sleep(2)
    click(34,264)

    time.sleep(1)
    click(20,203);


# In[91]:

def sendemail():
    authenticiation = (OFFICEun,OFFICEpw)
    m = Message(auth=authenticiation)

    m.setRecipients('matthew.bitter@mapleleaf.com')
    m.setSubject('BW and DS chains are completed. Portal updated. Reconciliation reports tie')
    m.setBody('')
    m.sendMessage();


# In[ ]:

if (df['Variance'] > 2).any() or (df['Variance'] < -2).any():
    print("Recon reports do not match - exiting")
else:
    compare()
    sendemail()


# In[ ]:



