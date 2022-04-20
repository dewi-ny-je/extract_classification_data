import sys
from os import listdir
from os.path import isfile, join

# print arguments from 1 (no script name)
#print(sys.argv[1:])

# print files in current dir
#print(os.listdir(os.getcwd()))

# list files in current dir
#allFiles = os.listdir(sys.argv[1])

onlyFiles = [f for f in listdir(sys.argv[1]) if isfile(join(sys.argv[1], f))]

for x in onlyFiles:
  print(x)
