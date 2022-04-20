import sys
from os import listdir
from os.path import isfile, join
import re

# print arguments from 1 (no script name)
#print(sys.argv[1:])

# print files in current dir
#print(os.listdir(os.getcwd()))

# list files in current dir
#allFiles = os.listdir(sys.argv[1])

onlyFiles = [f for f in listdir(sys.argv[1]) if isfile(join(sys.argv[1], f))]

#r = re.compile("^.*[xls,xlsx,ppt,pptx,doc,docx]$")
r = re.compile(".*\.ppt[x]{,1}$", re.IGNORECASE)
onlyPowerpoint = list(filter(r.match, onlyFiles))
r = re.compile(".*\.doc[x]{,1}$", re.IGNORECASE)
onlyWord = list(filter(r.match, onlyFiles))
r = re.compile(".*\.xls[x]{,1}$", re.IGNORECASE)
onlyExcel = list(filter(r.match, onlyFiles))
r = re.compile(".*\.pdf$", re.IGNORECASE)
onlyPDF = list(filter(r.match, onlyFiles))

for x in onlyPowerpoint:
  print(x)

for x in onlyWord:
  print(x)

for x in onlyExcel:
  print(x)

for x in onlyPDF:
  print(x)
