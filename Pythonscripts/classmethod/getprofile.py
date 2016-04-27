import os,re
import getpass
def get_profile():
    user=getpass.getuser()
    path="C:/Users/" + user+'/AppData/Roaming/Mozilla/Firefox/Profiles'
    filenames=os.listdir(path)
    for file in filenames:
        if re.match("\S{8}.default",file):
            profile=path+'/'+file
            return profile




