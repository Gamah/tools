import os
import sys
from fnmatch import fnmatch

def getFiles():
    branchPath = input("Enter path to be searched: ")
    pattern = input("Enter file name pattern to match: ")
    checked = 0
    found = 0
    foundFiles = {}
    
    for path, subdirs, files in os.walk(branchPath):
        for name in files:
            checked = checked + 1
            if fnmatch(name,pattern):
                foundFiles[found] = str(path + "/" +name)
                found = found + 1
    while(True):
        response = input("Checked %s files and found %s matches, continue? ([Y]es, [N]o, [L]ist files)\r\n" % (checked,len(foundFiles)))
        if response == "Y":
            return foundFiles
        elif response == "L":
            for file in foundFiles:
                print(foundFiles[file])
        elif response == "N":
            sys.exit()


def getInputs():
    search = []
    replace = []

    while(True):
        tmpSearch = input("Enter string to be matched... (Leave blank to finish): \r\n")
        if tmpSearch == "":
            return(search,replace)
        tmpReplace = input("Enter string to replace it with... (Leave blank to delete): \r\n")
        search.append(tmpSearch)
        replace.append(tmpReplace)


def doChanges(foundFiles, search, replace):
    hits = {}
    newFiles = {}
    
    #initialize hits variable...
    for s in search:
        hits[s] = 0
    
    print("Searching %s files for search hits..." % len(foundFiles))
    for foundFile in foundFiles:
        localFileData = ""
        makeNewFile = 0

        with open(foundFiles[foundFile],'r') as localFile:
            #print("Reading file %s..." % (foundFiles[foundFile]))
            localFileData = localFile.read()
    
            for s,r in zip(search,replace):
                localhits = localFileData.count(s)
                if localhits > 0:
                    hits[s] = hits[s] + localhits
                    localFileData = localFileData.replace(s,r)
                    makeNewFile = 1
                    #print("Found match on [%s], replacing with [%s]" % (s,r))
            if makeNewFile == 1:
                newFiles[foundFiles[foundFile]] = localFileData
    totalhits = 0
    for s,r,h in zip(search,replace,hits):
        print("    %s hits on '%s' (replacing with '%s')" % (hits[h],s,r))
        totalhits = totalhits + hits[h]
    while(True):
        response = input("Found %s total hits in %s file(s), write changes? ([Y]es, [N]o, [L]ist Changed Files)\r\n" % (totalhits,len(newFiles)))
        if response == "L":
            for file in newFiles:
                print(file)
        elif response == "N":
            sys.exit()
        elif response == "Y":
            for file in newFiles:
                print("Writing changes to %s..." % file)
                with open(file,'w') as newFile:
                    newFile.write(newFiles[file])
            return
     

def main():
    fileList = getFiles()
    search,replace = getInputs()
    doChanges(fileList,search,replace)
    

main()
