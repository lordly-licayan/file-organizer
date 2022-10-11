import configparser, os, re, time, xlsxwriter
import logging
import hashlib
from pathlib import Path
from os.path import exists, join
from shutil import copy2
from datetime import datetime

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

configFileName= 'config.ini'

currentPath= Path(__file__).resolve().parent
config_file = join(currentPath, configFileName)
config = configparser.ConfigParser()
config.read(config_file, encoding='UTF-8')

def getFileDateTime(mDateTime):
    return datetime.strptime(mDateTime, '%a %b %d %X %Y')

def getModifiedDateTime(filePath):
    return getFileDateTime(time.ctime(os.path.getmtime(filePath)))

def getMd5(filePath):
    md5Checksum= ''
    md5Hash = hashlib.md5()
    with open(filePath,"rb") as f:
        # Read and update hash in chunks of 4K
        for byte_block in iter(lambda: f.read(4096),b""):
            md5Hash.update(byte_block)
        md5Checksum= md5Hash.hexdigest()
    return md5Checksum

def listFiles(path, fileSearchPattern):
    fileInfo = {}
    
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
        print(f'\n::=> ' + r)
        for file in f:
            filePath= join(r, file)
            baseFile= re.search(fileSearchPattern, filePath, re.IGNORECASE)
                        
            if baseFile:
                fileData= []
                md5Checksum= getMd5(filePath)
                
                #Phase 1: Create a dictionary containing the MD5 as key and file data as the value.
                if md5Checksum in fileInfo:
                    fileData= fileInfo[md5Checksum]
                else:
                    fileInfo[md5Checksum]= fileData

                #Phase 2: Create a list containing the list of paths.
                fileData.append(filePath)
    return fileInfo

def getMd5Map(path, fileSearchPattern):
    md5Map = {}
    
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
        print(f'\n:: ' + r)
        for file in f:
            filePath= join(r, file)
            baseFile= re.search(fileSearchPattern, filePath, re.IGNORECASE)
                        
            if baseFile:
                md5List = []
                md5Checksum= getMd5(filePath)
                if md5Checksum in md5Map:
                    md5List = md5Map[md5Checksum]
                else:
                    md5Map[md5Checksum] = md5List
                
                md5List.append(filePath)
                
    return md5Map

def log(fileSheetRow, fileName, filePath, destPath, remarks):
    print(f'\nNo:{fileSheetRow}; Filename:{fileName}; Source={filePath}; Destination={destPath}; Remarks={remarks}')

def fillUpReport(fileSheet, rowNo, colNo, rowColor, md5, counter, fileName, sourcePath, destPath, remarks):
    fileSheet.write_string(rowNo, colNo, str(rowNo), rowColor)
    fileSheet.write_string(rowNo, colNo + 1, md5, rowColor)
    fileSheet.write_string(rowNo, colNo + 2, str(counter), rowColor)
    fileSheet.write_string(rowNo, colNo + 3, fileName, rowColor)
    fileSheet.write_string(rowNo, colNo + 4, sourcePath, rowColor)
    fileSheet.write_string(rowNo, colNo + 5, destPath, rowColor)
    fileSheet.write_string(rowNo, colNo + 6, remarks, rowColor)

def process(outputPath, fileInfo, md5Map):
    workbook_name = join(outputPath, f'File_organization_{time.time_ns()}.xlsx')
    workbook = xlsxwriter.Workbook(workbook_name)
    headerStyle = workbook.add_format({'border' : 1,'bg_color' : '#C6EFCE', 'bold': True})
    rowStyle = workbook.add_format({'border' : 1,'bg_color' : '#FFE5CC'})
    borderStyle = workbook.add_format({'border': 1})

    #Create Sheet for unchanged Files.
    if fileInfo:
        fileSheet = workbook.add_worksheet('List of files')
        fileSheet.set_column('A:A',5)
        fileSheet.set_column('B:B',40)
        fileSheet.set_column('C:C',8)
        fileSheet.set_column('D:D',35)
        fileSheet.set_column('E:E',50)
        fileSheet.set_column('F:F',50)
        fileSheet.set_column('G:G',15)
        
        fileSheet.write('B1','MD5 checksum', headerStyle)
        fileSheet.write('C1','Count', headerStyle)
        fileSheet.write('D1','Filename', headerStyle)
        fileSheet.write('E1','Source', headerStyle)
        fileSheet.write('F1','Destination', headerStyle)
        fileSheet.write('G1','Remarks', headerStyle)
        
        fileSheetRow = 0
        fileSheetCol = 0
        md5Counter = 0
        
        #Iterate the list of md5 with list of file paths.
        for md5Checksum in fileInfo:
            fileList = fileInfo[md5Checksum]
            rowColor =  rowStyle if (md5Counter%2 == 0) else borderStyle
            md5Counter += 1
            md5List = []
            remarks = 'Duplicate file!'
            
            #Check if file already exist in the destination folder
            if md5Checksum in md5Map:
                counter = 0
                remarks = 'Already exist!'
                md5List = md5Map[md5Checksum]
                #Iterate the md5List
                for md5FilePath in md5List:
                    fileCounter = 0
                    for filePath in fileList:
                        fileSheetRow += 1
                        fileCounter += 1
                        fileName = os.path.basename(filePath)
        
                        fillUpReport(fileSheet, fileSheetRow, fileSheetCol, rowColor, md5Checksum, fileCounter, fileName, filePath, md5FilePath, remarks)            
                        log(fileSheetRow, fileName, filePath, md5FilePath, remarks)
                continue
            
            #If file not in the destination folder then proceed.  
            md5Map[md5Checksum] = md5List
            md5FilePath = fileList[0]
            md5FileModification= getModifiedDateTime(md5FilePath)
            fileCounter = 0
            isCopied =  False
            
            #Identify the oldest file of the same md5.
            for filePath in fileList:
                fileModification= getModifiedDateTime(filePath)
                if(fileModification < md5FileModification):
                    md5FileModification =  fileModification
                    md5FilePath = filePath
            
            #Iterate the files to process
            for filePath in fileList:
                fileName = os.path.basename(filePath)
                dateModified = time.ctime(os.path.getmtime(filePath))
                fileModification= getFileDateTime(dateModified)
            
                timeStamp= datetime.strptime(dateModified, '%c')
                newFilename = str(timeStamp.strftime('%Y-%b-%d_%H.%M.%S'))
                    
                year= str(timeStamp.year)
                yearDir= os.path.join(outputPath, year)
                
                month= str(timeStamp.strftime('%b'))
                monthDir= os.path.join(yearDir, month)
                
                fileExt= Path(fileName).suffix
                destPath = os.path.join(monthDir, newFilename + fileExt)
                                   
                if not os.path.exists(yearDir):
                    os.makedirs(yearDir)
                
                if not os.path.exists(monthDir):
                    os.mkdir(monthDir)

                fileSheetRow += 1
                fileCounter += 1
                counter= 0
                isForCopying= False
                remarks = 'Duplicate file!'
                    
                if (not isCopied) and (fileModification == md5FileModification):
                    if not os.path.exists(destPath):
                        isForCopying= True
                    else:
                        dirList = os.listdir(monthDir)
                        for item in dirList:
                            #Check the same file name for file renaming.
                            fileNamePattern = re.compile("_\((.*)\)".join([newFilename, fileExt]), re.IGNORECASE)
                            hasMatch= fileNamePattern.search(item)
                            
                            if hasMatch:
                                index= int(hasMatch.groups()[0])
                                if(index > counter):
                                    counter = index
                                newFilename += '_(' + str(counter + 1) + ')'
                                destPath = os.path.join(monthDir, newFilename + fileExt)
                                isForCopying= True
                                break
                                
                if isForCopying:
                    copy2(filePath, destPath)
                    md5List.append(destPath)
                    isCopied= True
                    remarks = 'Copied'
                
                fillUpReport(fileSheet, fileSheetRow, fileSheetCol, rowColor, md5Checksum, fileCounter, fileName, filePath, destPath, remarks)
                log(fileSheetRow, fileName, filePath, destPath, remarks)

    workbook.close()

if __name__ == "__main__":
    try:
        start = datetime.now()
        print(f'\nInitializing...\nTime started: {start}')
        
        sourcePath= config['PATH']['SOURCE']
        outputPath= config['PATH']['OUTPUT_PATH']
        fileSearchPattern = config['OTHERS']['FILES_SEARCH_PATTERN']
        
        fileInfo = listFiles(sourcePath, fileSearchPattern)
        md5Map = getMd5Map(outputPath, fileSearchPattern)
        
        if outputPath:
            destinationPath= outputPath
        else:
            destinationPath= currentPath

        #process and create report
        process(destinationPath, fileInfo, md5Map)

        finish = datetime.now()
        print(f'\nTime elapsed:\n{finish - start}')
    except Exception as err:
        logger.error(str(err), exc_info=True)
        #input("")
        print(f'\Error...')
