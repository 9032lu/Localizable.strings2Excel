# -*- coding:utf-8 -*-

from optparse import OptionParser
from XlsFileUtil import XlsFileUtil
from XmlFileUtil import XmlFileUtil
from StringsFileUtil import StringsFileUtil
from Log import Log
import os
import time
import shutil

def addParser():
    parser = OptionParser()

    parser.add_option("-f", "--fileDir",
                      help="Xls files directory.",
                      metavar="fileDir")

    parser.add_option("-t", "--targetDir",
                      help="The directory where the strings files will be saved.",
                      metavar="targetDir")

    parser.add_option("-e", "--excelStorageForm",
                      type="string",
                      default="single",
                      help="The excel(.xls) file storage forms including single(single file), multiple(multiple files), default is multiple.",
                      metavar="excelStorageForm")

    parser.add_option("-a", "--additional",
                      help="additional info.",
                      metavar="additional")

    (options, args) = parser.parse_args()
    print("options: %s, args: %s" % (options, args))

    return options


def convertFromSingleForm(options, fileDir, targetDir):
    for _, _, filenames in os.walk(fileDir):
        # print(f'---filenames--{filenames}-')
        xlsFilenames = [fi for fi in filenames if fi.endswith(".xlsx")]
        if len(xlsFilenames)==0:
            Log.info('-----没有可用.xlsx文件')
        for file in xlsFilenames:
            Log.info(f"======={file}===")
            xlsFileUtil = XlsFileUtil(fileDir+"/"+file)
            table = xlsFileUtil.getTableByIndex(0)
            countryCode = []
            for r in list(list(table.rows)[0])[1:]:
                countryCode.append(r.value)
            keys =[]
            for r in list(list(table.columns)[0])[1:]:
                keys.append(r.value)

            # print(countryCode)
            # firstRow =  list(list(table.rows)[0]) 
            # keys = list(list(table.columns)[0])
            # print(f"------countryCode-------------{countryCode}")
            # del keys[0]
            # print(f"-----------keys--------{keys}")
            for index in range(len(countryCode)):
                # if index <= 0:
                #     continue
                languageName = countryCode[index]
                values=[]
                for v in list(list(table.columns)[index+1])[1:]:
                    values.append(v.value)
                # print(f"------{index}--{languageName}--{values}-")

                # values = table.col_values(index)
                # del values[0]
                # print(f"-----options--{options}--")
                StringsFileUtil.writeToFile(
                    keys, values, targetDir + "/"+languageName+".lproj/", file.replace(".xlsx", "")+".strings", options.additional)
    print ("Convert %s successfully! you can see strings file in %s" % (
        fileDir, targetDir))
    Log.info('转换完成，速度杠杠的！！！')


def convertFromMultipleForm(options, fileDir, targetDir):
    for _, _, filenames in os.walk(fileDir):
        xlsFilenames = [fi for fi in filenames if fi.endswith(".xls")]
        for file in xlsFilenames:
            xlsFileUtil = XlsFileUtil(fileDir+"/"+file)
            langFolderPath = targetDir + "/" + file.replace(".xls", "")
            if not os.path.exists(langFolderPath):
                os.makedirs(langFolderPath)

            for sheet in xlsFileUtil.getAllTables():
                iosDestFilePath = langFolderPath + "/" + sheet.name
                iosFileManager = open(iosDestFilePath, "wb")
                for row in sheet.get_rows():
                    content = "\"" + row[0].value + "\" " + \
                        "= " + "\"" + row[1].value + "\";\n"
                    iosFileManager.write(content)
                if options.additional is not None:
                    iosFileManager.write(options.additional)
                iosFileManager.close()
    Log.info ("Convert %s successfully! you can see strings file in %s" % (
        fileDir, targetDir))


def startConvert(options):
    fileDir = options.fileDir
    targetDir = options.targetDir

    print ("Start converting")
    print(f"===={options.excelStorageForm}==")
    if fileDir is None:
        print ("xls files directory can not be empty! try -h for help.")
        return

    if targetDir is None:
        print ("Target file directory can not be empty! try -h for help.")
        return

    targetDir = targetDir + "/xls-files-to-strings_" + \
        time.strftime("%Y%m%d_%H%M%S")
    if not os.path.exists(targetDir):
        os.makedirs(targetDir)

    if options.excelStorageForm == "single":
        convertFromSingleForm(options, fileDir, targetDir)
    else:
        convertFromMultipleForm(options, fileDir, targetDir)

# 脚本执行放开main
# def main():
#     options = addParser()
#     startConvert(options)


# main()


class Xlsx2Strings:
    '工具用来执行方法'
    @staticmethod
    def startConvertXlsxToStrings(fileDir,targetDir):
        Log.info(f'xlsx路径:{fileDir}')
        Log.info(f'保存路径:{targetDir}')
        if fileDir == targetDir:
            Log.info('路径一致会被清空！！！！')
            return
        if not os.path.exists(targetDir):
            os.makedirs(targetDir)
        # else:
            # shutil.rmtree(targetDir)
            # os.makedirs(targetDir)

        
        convertFromSingleForm(addParser(), fileDir,targetDir)
