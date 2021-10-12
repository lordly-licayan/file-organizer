import configparser, os, re, datetime, time, xlsxwriter
import logging
from pathlib import Path
from os.path import exists, join
from shutil import copyfile
#from distutils.util import strtobool

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

configFileName= 'config.ini'

currentPath= Path(__file__).resolve().parent
config_file = join(currentPath, configFileName)
config = configparser.ConfigParser()
config.read(config_file, encoding='UTF-8')

def listFiles(path, fileSearchPattern):
    fileInfo = {}
    
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
        for file in f:
            filePath= join(r, file)
            baseFile= re.search(fileSearchPattern, filePath, re.IGNORECASE)
                        
            if baseFile:
                fileData= {}
                filename = os.path.basename(filePath)
                fileModification = time.ctime(os.path.getmtime(filePath))
                
                #Phase 1: Create a dictionary containing the date of modification as key and file data as the value.
                if fileModification in fileInfo:
                    fileData= fileInfo[fileModification]
                else:
                    fileInfo[fileModification]= fileData

                #Phase 2: Create a dictionary containing the filename and the list of paths.
                if filename in fileData:
                    fileList = fileData[filename]
                    fileList.append(filePath)
                else:
                    fileList= [filePath]
                    fileData[filename]=  fileList
    return fileInfo

def makeReport(outputPath, fileInfo):
    workbook_name = join(outputPath, f'File_organization_{time.time_ns()}.xlsx')
    workbook = xlsxwriter.Workbook(workbook_name)
    header = workbook.add_format({'border' : 1,'bg_color' : '#C6EFCE', 'bold': True})
    border = workbook.add_format({'border': 1})

    #Create Sheet for unchanged Files.
    if fileInfo:
        fileSheet = workbook.add_worksheet('List of files')
        fileSheet.set_column('A:A',5)
        fileSheet.set_column('B:B',30)
        fileSheet.set_column('C:C',20)
        fileSheet.set_column('D:D',50)
        fileSheet.set_column('E:E',50)
        fileSheet.set_column('F:F',18)

        fileSheet.write('B1','Creation time', header)
        fileSheet.write('C1','Filename', header)
        fileSheet.write('D1','Source', header)
        fileSheet.write('E1','Destination', header)
        fileSheet.write('F1','Remarks', header)
        fileSheetRow = 1
        fileSheetCol = 0
  
        for fileModification in fileInfo:
            fileData= fileInfo[fileModification]
            for filename in fileData:
                fileList = fileData[filename]
                for filePath in fileList:
                    timeStamp= datetime.datetime.strptime(fileModification, '%c')
                    newFilename = str(timeStamp.strftime('%Y-%m-%d_%H.%M.%S'))
                    year= str(timeStamp.year)
                    month= str(timeStamp.strftime('%b'))
                    fileSheet.write_number(fileSheetRow, fileSheetCol, fileSheetRow, border)
                    fileSheet.write_string(fileSheetRow, fileSheetCol+1, fileModification, border)
                    fileSheet.write_string(fileSheetRow, fileSheetCol+2, filename, border)
                    fileSheet.write_string(fileSheetRow, fileSheetCol+3, filePath, border)
                    
                    yearDir= os.path.join(outputPath, year)
                    monthDir= os.path.join(yearDir, month)
                    
                    if not os.path.exists(yearDir):
                        os.makedirs(yearDir)
                    
                    if not os.path.exists(monthDir):
                        os.mkdir(monthDir)

                    destPath = os.path.join(monthDir, newFilename + Path(filename).suffix)
                    fileSheet.write_string(fileSheetRow, fileSheetCol+4, destPath, border)
                    remarks= "Done"
                    
                    print(f'\nFilename:{filename}; Source={filePath}; Destination={destPath}; Modification={fileModification}; Year={year}; Month={month}')
                    
                    if not os.path.exists(destPath):
                        try:
                            copyfile(filePath, destPath)
                        except Exception as err:
                            remarks= "Error: Can't copy"
                    else:
                        remarks= "Already exist!"
                    
                    fileSheet.write_string(fileSheetRow, fileSheetCol+5, remarks, border)
                    fileSheetRow +=1
    workbook.close()

if __name__ == "__main__":
    try:
        start = datetime.datetime.now()
        print(f'\nInitializing...\nTime started: {start}')
        
        sourcePath= config['PATH']['SOURCE']
        filesSearchPattern = config['OTHERS']['FILES_SEARCH_PATTERN']
        fileInfo = listFiles(sourcePath, filesSearchPattern)

        outputPath= config['PATH']['OUTPUT_PATH']
        if outputPath:
            destinationPath= outputPath
        else:
            destinationPath= currentPath

        #create report
        makeReport(destinationPath, fileInfo)

        finish = datetime.datetime.now()
        print(f'\nTime elapsed:\n{finish - start}')
    except Exception as err:
        logger.error(str(err), exc_info=True)
        #input("")
        print(f'\nTesting...')
