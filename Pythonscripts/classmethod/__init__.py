from os.path import dirname
import os
path=dirname(__file__)
list=os.listdir(path)
a=[]
for i in list:
    if i.split('.')[1]=='py':
        a.append(i.split('.')[0])
__all__=a
