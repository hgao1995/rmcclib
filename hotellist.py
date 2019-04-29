"To get RMCC Hotel list under HOTELS Folder"
#Created by Henry Gao on the 9th Jan 2019
import os

def gethotellist():
    hotellist=[]
    hotel_folder_path=r'L:\RES\RevOM\HOTELS'
    for hotelfolder in os.listdir(hotel_folder_path):
        if len(hotelfolder)==5:
            hotellist.append(hotelfolder)
    return hotellist
        
