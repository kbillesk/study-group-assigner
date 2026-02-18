# -*- coding: utf-8 -*-
"""
Created on Tue May 24 20:54:36 2022

@author: krist
"""

import os
import json
os.chdir(r'C:\Users\krist\Desktop\classplanning')
with open('classes.txt','r', encoding=('utf8')) as d:
    classes = list(line.lower().rstrip("\n") for line in d)
#    print(classes)

classroomset = set()
with open('classrooms.txt','r', encoding=('utf8')) as d:
    for line in d:
        line = line.partition(",")
        classroomset.add(line[0])
#    print(classroomset)
weeklyslots = set()
with open('weeklyslots.txt','r', encoding=('utf8')) as d:
    for line in d:
        line = line.rstrip("\n")
        weeklyslots.add(line)
#    print(weeklyslots)
#here the set of available classes are generated in availableclasses combining weeklyslots with the classrooms
availableclasses = list()
for line in weeklyslots:
    for rooms in classroomset:
        availclass = line+"-"+rooms
        availableclasses.append(availclass)

#print(availableclasses)
#This assignes classes to available slots and classrooms
assignedclasses = dict()
for nclass in classes:
    i=1
    while availableclasses[i]:
        assignedclass = availableclasses[i-1]
        print(assignedclass)
        if assignedclasses.get(nclass):
            else:
                assignedclasses[assignedclass]= nclass
                availableclasses.pop(i-1)
        i+=1
    
print(assignedclasses)
#print(availableclasses)
