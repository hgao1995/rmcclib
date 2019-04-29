"quickly get desire mailbox and its folder"

#Created by Henry Gao on the 14th Jan 2019

import win32com.client
import os

def getinbox(acc):
    
    ol=win32com.client.Dispatch("Outlook.Application")

    myNameSpace=ol.GetNamespace('MAPI')

    myAcc=myNameSpace.CreateRecipient(acc+'@hilton.com')

    inbox=myNameSpace.GetSharedDefaultFolder(myAcc,6)

    return inbox
        

