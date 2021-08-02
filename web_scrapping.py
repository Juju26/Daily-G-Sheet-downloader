#https://sites.google.com/site/gdocs2direct/home
import webbrowser
import datetime
import os
import time
from pathlib import Path


def file_download():
    #before change chrome to default to avoid edge problems
    url='https://docs.google.com/spreadsheets/d/1tZkMDnfOzQbpCaYlM7ek0DXzhEIHca8mBdmr1pP_CQQ/export?format=xlsx'
    webbrowser.open_new(url)

def move_file():
    #for moving file
    try:
        '''get_date=str(datetime.date.today())
        newname=str(get_date+' Department of CSE- Fee Receipt Collector (Responses).xlsx')'''
        Path("C:/Users/Asus/Downloads/Department of CSE- Fee Receipt Collector (Responses).xlsx").rename("C:/Users/Asus/OneDrive/Desktop/projects rough/final/Department of CSE- Fee Receipt Collector (Responses).xlsx") #hijack file from download to destination folder
        time.sleep(5)
    except FileNotFoundError:
        print("File not downloaded yet")

def change_name():
    # to change the name to date file
    file_name='pyd'
    get_date=str(datetime.date.today())
    hou=str(datetime.datetime.now().hour)
    mi=str(time.strftime("%M"))
    tim=hou+'-'+mi
    newname=str(get_date+' '+tim+' '+file_name+'.xlsx')
    os.chdir("C:/Users/Asus/OneDrive/Desktop/projects rough/final/")  #destination folder path
    try:
        os.rename("Department of CSE- Fee Receipt Collector (Responses).xlsx",newname)    #download file name, replace name 
    except FileExistsError as e:
        print('File already found')
    #print(new_path)

#driver area
print("Downloading files")
file_download()
time.sleep(6)
os.system("taskkill /f /im msedge.exe") #kill the browser
print("Moving files")
#script should be in destination folder
move_file()
print("changing name")
change_name()
print('Done')