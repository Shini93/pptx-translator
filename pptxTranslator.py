#takes pptx in german
#makes xml with many files out of it
#takes files and gets english sentences from it
#writes english sentences in xml files
#changes xml to pptx again

#imports
import subprocess       #to read and write pptx
import shutil           #to move files
import os               #tp read folder
import glob             #to move in subfolders
import re               #to find multiple substrings easy
import deepl            #to translate on deepl.com

#globals
fileNames = []      #list of files with .pptx in the folder
saveFolder = "saveFolderPython"
saveFolderConverts = saveFolder + "/convertedFiles"
slideFolder = "/ppt/slides"
auth_key = "getYourOwnKey"
target_language = "EN-US"


def main():
    getAllPPTXs()
    for name in fileNames:
        editFile(name)
    
#finds all pptxs in the folder
def getAllPPTXs():
    print("allpptx")
    global fileNames
    fileNames = getAllFiles('.',".pptx")

def getAllSlides(name):
    print("allSlides")
    global slideFolder
    global saveFolderConverts
    path = saveFolderConverts+'/'+name+slideFolder
    slides =  getAllFiles(path,".xml")
    return slides

def getAllFiles(path, ending):
    array = []
    if path == '.':
    
        for file_path in os.listdir(path):
            if os.path.isfile(os.path.join('.', file_path)) and ending in file_path:
                file_path = file_path[0:-len(ending)]
                array.append(file_path)
        return array
    else:
        array = glob.glob(path+"/*.xml")
        for a in range (len(array)):
            index = array[a].find("slide",len(array[a])-15)
            array[a] = array[a][index:-4]
        return array
        

#reads file, changes to xml, translate and changes to pptx
def editFile(name):
    convertToXMLFiles(name)
    slides = getAllSlides(name)
    iterateSlides(slides, name)
    convertToPPTXFile(name)

#converts pptx to xml files
def convertToXMLFiles(name):
    global saveFolderConverts
    save_dir = saveFolderConverts + "/" + name
    name_ = name + ".pptx"
    subprocess.run(f'opc extract {name_} {save_dir}', shell=True, check=True)

#saves xml files to one new ppt again
def convertToPPTXFile(name):
    global saveFolderConverts
    global saveFolder
    save_dir = saveFolderConverts + "/" + name
    newName = "new_"+ name + ".pptx"
    subprocess.run(f'opc repackage {save_dir} {newName}', shell=True, check=True)
    shutil.move(newName, saveFolder+"/"+newName)

#iterates through all slides and finds text and translates it
def iterateSlides(slides, name):
    for s in slides:
        readAndTranslateSlide(s, name)

def readAndTranslateSlide(slide, name):
    global saveFolderConverts
    global slideFolder
    sliderDir = saveFolderConverts+'/'+name+slideFolder
    fullPath = sliderDir + "/" + slide+".xml"
    f = open(fullPath,"r", encoding="utf8")
    lines = f.read()
    f.close()
    start = findAllSubStrings(lines,"<a:t>")
    end = findAllSubStrings(lines,"</a:t>")
    allText = []    
    for i in range(len(start)):
        allText.append(lines[start[i]+5:end[i]])

    translated = getTranslation(allText)
    for i in range(len(start)-1,0,-1):
        dummy1 = lines[0:start[i]+5]
        dummy2 = lines[end[i]:]
        lines = dummy1+translated[i]+dummy2
    
    overWriteXML(lines,fullPath)

def getTranslation(allText):
    new = []
    for i in range(len(allText)):
        translation = getTranslationDeepl(allText[i])
        new.append(translation)
    return new

def getTranslationDeepl(text):
    if text == "" or text == " " or text == "  " or text == "    ":
        return text
    if text.isnumeric() == True:
            return text
    print(text)
    translator = deepl.Translator(auth_key) 
    result = translator.translate_text(text, target_lang=target_language) 
    translated_text = result.text
    print(translated_text)
    return translated_text


def overWriteXML(text,path):
    os.remove(path) 
    f = open(path, "a", encoding="utf8")
    f.write(text)
    f.close()

def findAllSubStrings(ori, sub):
    return [i.start() for i in re.finditer(sub, ori)]

    
main()
