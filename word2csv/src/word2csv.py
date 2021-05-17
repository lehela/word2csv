import os
import time
import sys
import gc
import errno
import alive_progress as prgr
import configparser
import numpy as np
import zipfile
import pandas as pd
#from pandas import DataFrame, isna
import csv
import lxml.etree as etree
#from lxml import etree
from threading import current_thread
from types import SimpleNamespace


# Helper Functions

def getText(xmlElem, delim : str = "") -> str:
    tNodes = xmlElem.xpath(".//w:t", namespaces=xmlElem.nsmap)
    # tNodes = xmlElem.findall(".//w:t", xmlDoc.nsmap)
    snippets = []
    for t in tNodes:
        snippets.append(t.text)
    result = delim.join(snippets) 
    return result.strip()

def resolveNS(text, namespaces):
    tokens = text.split(":")
    resolved = "{"+namespaces[tokens[0]]+"}"+tokens[1]
    return resolved

def loadDocx(docPath):
    zipf = zipfile.ZipFile(docPath)
    return zipf

def getComments(zipf):

    dictComments = {}

    try:
        raw = zipf.read("word/comments.xml")
    except:
        return dictComments

    xmlComments = etree.fromstring(raw)
    for xmlComment in xmlComments.xpath("./w:comment", namespaces=xmlComments.nsmap):
   
        id = xmlComment.get(resolveNS("w:id", namespaces=xmlComment.nsmap))
        author = xmlComment.get(resolveNS("w:author", namespaces=xmlComment.nsmap))
        metatext = getText(xmlComment)

        dictComments[id] = {
            "ID" : id,
            "Author" : author,
            "Metatext" : metatext
        }

    return dictComments

def getFootnotes(zipf):

    dictFootnotes = {}

    try:
        raw = zipf.read("word/footnotes.xml")
    except:
        return dictFootnotes

    xmlFootnotes = etree.fromstring(raw)
    for xmlFootnote in xmlFootnotes.xpath("./w:footnote", namespaces=xmlFootnotes.nsmap):
   
        id = xmlFootnote.get(resolveNS("w:id", namespaces=xmlFootnote.nsmap))

        dictFootnotes[id] = {
            "ID" : id,
            "Metatext" : getText(xmlFootnote)
        }

    return dictFootnotes

def getStyles(zipf):

    dictStyles = {}

    try:
        raw = zipf.read("word/styles.xml")
    except:
        return dictStyles

    xmlStyles = etree.fromstring(raw)

    for xmlStyle in xmlStyles.xpath("./w:style", namespaces=xmlStyles.nsmap):
   
        styleId = xmlStyle.get(resolveNS("w:styleId", namespaces=xmlStyle.nsmap))
        styleType = xmlStyle.get(resolveNS("w:type", namespaces=xmlStyle.nsmap))
        
        outlineLvl = np.nan
        for xmloutlineLvl in xmlStyle.xpath(".//w:outlineLvl", namespaces=xmlStyle.nsmap):
            outlineLvl = xmloutlineLvl.get(resolveNS("w:val", namespaces=xmlStyles.nsmap))

        dictStyles[styleId] = {
            "Type" : styleType,
            "outlineLvl" : outlineLvl
        }

    return dictStyles

def getDocumentBody(zipf):
    raw = zipf.read("word/document.xml")
    xmlDoc = etree.fromstring(raw)
    xmlDocBody = xmlDoc.xpath(".//w:body", namespaces=xmlDoc.nsmap)[0]
    return xmlDocBody


def parseHeading(df: pd.DataFrame, row: dict, xmlNode, dStyles: dict) -> pd.DataFrame:

    # Exit immediately if the row is not representing a paragraph, or the paragraph has no content
    if row["tag"] != "p" or getText(xmlNode) == "":
        return df

    # Fetch the paragraph's Style ID
    styleId = np.nan
    for styleNode in xmlNode.xpath(".//w:pStyle", namespaces=xmlNode.nsmap):
        styleId = styleNode.get(resolveNS("w:val", namespaces=xmlNode.nsmap),"")
    row["styleId"] = styleId

    # Fetch the Outline Level from the Style
    row["outlineLvl"] = dStyles.get(styleId, {"outlineLvl" : np.nan})["outlineLvl"]

    # If the Style has no Outline Level then exit
    if pd.isna( row["outlineLvl"] ):
        return df

    # Main Parsing Logic
    # ---------------------
    global dOutlineNumbers

    # Increase current level outline
    currentOutlineLevel = int(row["outlineLvl"])
    dOutlineNumbers[currentOutlineLevel] = dOutlineNumbers.get(currentOutlineLevel, 0) + 1

    # Reset lower level outlines
    lvl_low = currentOutlineLevel + 1
    lvl_high = len(dOutlineNumbers.keys())
    for idx in range(lvl_low, lvl_high):
        dOutlineNumbers[idx] = 0    

    # Generate Outline Reference
    refList = list(dOutlineNumbers.items())[0:int(row["outlineLvl"]) + 1]
    refText = ".".join(str(itm[1]) for itm in refList)

    # Populate record attributes
    row["paraID"] = xmlNode.get(resolveNS("w14:paraId", namespaces=xmlNode.nsmap))
    row["ChapterOutline"] = refText
    row["ChapterName"] = getText(xmlNode)

    row["Type"] = "Heading"
    row["Text"] = getText(xmlNode)

    # Save row to dataframe
    df = df.append(row, ignore_index=True)
    
    # Check for Comments & Footnotes
    df = parseComments(df, row, xmlNode)
    df = parseFootnotes(df, row, xmlNode)
    
    return df

def parseComments(df: pd.DataFrame, row: dict, xmlNode) -> pd.DataFrame:
    
    for cmt in xmlNode.xpath(".//w:commentRangeStart", namespaces=xmlNode.nsmap):
        # Collect all text from subsequent nodes until end of comment range
        anchor = ""
        curr = cmt.xpath(".//following-sibling::*")[0]
        done = False
        while etree.QName(curr.tag).localname != "commentRangeEnd" and done == False:
            anchor = anchor+getText(curr)
            following = curr.xpath(".//following-sibling::*")
            if len(following) > 0:
                curr = following[0]
            else:
                done = True
        
        row["ID"] = cmt.get(resolveNS("w:id", namespaces=cmt.nsmap))
        commentDef = dComments[row["ID"]]

        row["Type"] = "Comment"
        row["Text"] = anchor
        row["Metatext"] = commentDef.get("Metatext")
        row["Author"] = commentDef.get("Author")
        df = df.append(row, ignore_index=True)

    return df

def parseFootnotes(df: pd.DataFrame, row: dict, xmlNode) -> pd.DataFrame:
    
    global footnoteIndex

    for fnt in xmlNode.xpath(".//w:footnoteReference", namespaces=xmlNode.nsmap):

        # Get preceding runs
        runs = fnt.xpath("./..//preceding-sibling::w:r", namespaces = xmlNode.nsmap)

        # Try to isolate runs with DiagramShape style
        anchors = []
        anchor = ""
        lastStyle = ""
        firstPass = True
        for run in runs:
            try:
                Styles = run.xpath("./w:rPr/w:rStyle", namespaces=xmlNode.nsmap)
                currentStyle = Styles[0].get(resolveNS("w:val", namespaces=xmlNode.nsmap))
            except:
                currentStyle = "Default"
            if lastStyle == currentStyle or firstPass == True:
                anchor = anchor+getText(run)
            else:
                if currentStyle == "DiagramShape":
                    anchor = getText(run)
                else:
                    if len(anchor) > 0:
                        anchors.append(anchor)
                    anchor = getText(run)

            lastStyle = currentStyle
            firstPass = False

        if currentStyle == "DiagramShape" and len(anchor) > 0:
            anchors.append(anchor)
            anchor = ""

        if len(anchors) > 0:
            marked = anchors[-1]+anchor
        else:
            marked = anchor
        
        xmlID = fnt.get(resolveNS("w:id", namespaces=fnt.nsmap))
        footnoteDef = dFootnotes[xmlID]

        footnoteIndex += 1
        row["ID"] = footnoteIndex
        row["Type"] = "Footnote"
        row["Text"] = marked.strip()
        row["Metatext"] = footnoteDef.get("Metatext")

        df = df.append(row, ignore_index=True)

    return df

def parseParagraph(df: pd.DataFrame, row: dict, xmlNode) -> pd.DataFrame:

    # If row not a paragraph, exit immediately 
    if row["tag"] != "p":
        return df

    # If paragraph is empty, then exit immediately
    if getText(xmlNode) == "":
        return df

    # If paragraph is already a header, then exit immediately
    if not(pd.isna(row["outlineLvl"])):
        return df

    row["paraID"] = xmlNode.get(resolveNS("w14:paraId", namespaces=xmlNode.nsmap))
    row["Type"] = "Paragraph"
    row["Text"] = getText(xmlNode)
    row["ParagraphRef"] = parseParRef(row["Text"])

    df = df.append(row, ignore_index=True)
    df = parseComments(df, row, xmlNode)
    df = parseFootnotes(df, row, xmlNode)
    
    return df

def parseTable(df: pd.DataFrame, row: dict, xmlNode) -> pd.DataFrame:

    # Not a table -> Return
    if row["tag"] != "tbl" :
        return df

    row["Type"] = "Paragraph"

    for tr in xmlNode.xpath("w:tr", namespaces=xmlNode.nsmap):
        row["paraID"] = tr.get(resolveNS("w14:paraId", namespaces=xmlNode.nsmap))
        row["Text"] = getText(tr, "|")
        row["ParagraphRef"] = parseParRef(row["Text"])
        df = df.append(row, ignore_index=True)
        df = parseComments(df, row, tr)
        df = parseFootnotes(df, row, tr)
    
    return df

def parseParRef(text: str = " [3.1.3] Test") -> str:
    parsed = text.strip().split("]", 1)[0].split("[", 1)
    if len(parsed) > 1:
        token = "["+parsed[1]+"]"
    else:
        token = np.nan
    return token

def fillNaN(df : pd.DataFrame) -> pd.DataFrame :
    
    df.sort_index(axis=0)

    # Fill down specific columns
    cols = ["Document","ChapterOutline", "ChapterName"]
    df.loc[:,cols] = df.loc[:,cols].ffill()
    
    return df

def getFilePaths(argv):
    
    #print(argv)
    docxPaths = []
    print("\n")
    for path in argv[1:]:
        if os.path.splitext(path)[1] == ".docx":
            docxPaths.append(path)
        else:
            print("Not a Word file: {path}\n".format(path = path))
    
    return docxPaths

def checkPath(path):
    try:
        checkFile = os.path.join(path,"check.txt")
        with open(checkFile, 'w'):
            pathExists = True
    except IOError as x:
        pathExists = False
        msg = "'{path}' is invalid. "
        if x.errno == errno.EACCES:
            msg = msg + "It cannot be written to.. "
        print("--> "+msg.format(path = path))
    if pathExists:
        os.remove(checkFile)

    return pathExists

def getOutputPath(argv):

    # Fetch the output path from ini file
    ownPath = os.path.split(sys.argv[0])[0]
    configPath = ownPath + "\\word2csv.ini"
    cp = configparser.ConfigParser()
    if os.path.isfile(configPath):
        cp.read(configPath)
    else:
        cp.add_section("output")
        cp.set("output", "path", "C:\\")
    path = cp.get("output","path")

    # Confirm Output Path
    done = False
    while done == False:
        pathNew = input("\n\nPress return to accept default output location, or enter new path.\n\n[ {currPath} ]: ".format(currPath = path))
        if pathNew == "":
            done = True
        else:
            pathExists = checkPath(pathNew) 
            if pathExists:
                path = pathNew
                cp.set("output", "path", path)
                with open(configPath, 'w') as configfile:
                    cp.write(configfile)                
                done = True

    if path == "":
        path = ownPath
    return os.path.realpath(path)

def convertDocx(docPath : str, outputPath : str ) :
    docName = os.path.split(docPath)[1]

    docx = loadDocx(docPath)
    
    dStyles = getStyles(docx)
    
    global dComments
    dComments = getComments(docx)
    
    global dFootnotes 
    dFootnotes = getFootnotes(docx)
    
    global dOutlineNumbers  
    dOutlineNumbers = {}

    global footnoteIndex
    footnoteIndex = 0

    # Main Loop

    treeBodyNodes = getDocumentBody(docx)

    idx=0
    df_csv = pd.DataFrame()
    with prgr.alive_bar(total=len(treeBodyNodes), spinner="stars") as bar:
        for xmlNode in treeBodyNodes:
            
            row = {
                "Index" : idx, 
                "Document" : docName,
                "ChapterOutline" : np.nan,
                "ChapterName" : np.nan,
                "ParagraphRef" : np.nan,
                "Page" : np.nan,
                "Type" : np.nan,
                "ID" : np.nan,
                "Text" : np.nan,
                "Metatext" : np.nan,
                "Author" : np.nan,
                "tag" : etree.QName(xmlNode.tag).localname, 
                "xml": xmlNode 
                }

            df_csv = parseHeading(df_csv, row, xmlNode, dStyles)
            df_csv = parseParagraph(df_csv, row, xmlNode)
            df_csv = parseTable(df_csv, row, xmlNode)

            idx+=1
            bar()

    df_csv = fillNaN(df_csv)

    # Save to CSV
    csvCols = [
                "Index", 
                "Document",
                "ChapterOutline",
                "ChapterName",
                "ParagraphRef",
                "Page",
                "Type",
                "ID",
                "Text",
                "Metatext",
                "Author"
    ]

    df_csv.to_csv(os.path.join( outputPath, docName+".csv"), index=False, header=True, columns=csvCols, quoting=csv.QUOTE_NONNUMERIC)    
    return

# Globals
dComments = {}
dFootnotes = {}
dOutlineNumbers = {}
footnoteIndex = 0
indent = "\t"

os.system('cls' if os.name == 'nt' else 'clear')
print(  
    "\n"+indent+"Word2Csv Converter\n"+indent+"=================="+
    "\n\n"+
    "IMPORTANT: Word documents must have all tracked changes ACCEPTED, or else the conversion may fail.."
    )

outPath = getOutputPath(sys.argv)

# Get dropped files
docPaths = getFilePaths(sys.argv)
print(  
    "\n\nConverting {docCount} documents. ".format(docCount = len(docPaths))
    )

if len(docPaths) == 0:
    print("\nTry dragging & dropping Word files next time...")

# Console Printout
docDisplay = "{idx}/"+str(len(docPaths))+" : Convert {docTitle}"

# Main Loop
idx = 0
for docPath in docPaths:
    idx += 1
    docTitle = os.path.split(docPath)[1]
    print(docDisplay.format(idx = idx, docTitle = docTitle))
    convertDocx(docPath, outPath)
    print("\n")

print("\n\nConversion finished.")

time.sleep(5) 