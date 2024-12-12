import PySimpleGUI as sg
import json
from Xls2Strings import Xlsx2Strings as XS
from Strings2Xls import Strings2Xlsx as SX
from Xml2Xls import Xml2Xlsx
from Xls2Xml import Xlsx2Xml
from Log import Log

settings = sg.UserSettings()

lista=['xlsxToStrings','stringsToXlsx','xlsxToxml','xmlToXlsx']
layout=[
    [sg.FolderBrowse('选择文件夹',size=(8,1)),sg.In(key='-input_file_dir-',disabled=True,default_text=settings.get('-input_file_dir-',''))],
    [sg.FolderBrowse('保存到',size=(8,1)),sg.In(key='-output_file_dir-',disabled=True,default_text=settings.get('-output_file_dir-',''))],
    [sg.R(i,default= (i==lista[0]),group_id=1,key=i) for i in lista],
    [sg.B('执行',size=(15,1))],
    [sg.ML(size=(65,10),reroute_cprint=True)],
   
]
window = sg.Window("多语言转换工具",layout)
while True:
    event,values=window.read()
    if  values== None:
        break
    input_dir = values['-input_file_dir-']
    output_dir = values['-output_file_dir-']
    if input_dir is None:
        Log.info('请选择源文件夹')
    else:
         settings['-input_file_dir-']=values['-input_file_dir-']

    if output_dir is None:
        Log.info('请选择保存文件夹')
    else:
        settings['-output_file_dir-']=values['-output_file_dir-']

    
    # if len(output_dir)>0:
    #     output_dir =  output_dir +'/sources/'
    # print(len(output_dir))
    if event == None:
        break
    if event=='执行' and len(output_dir)>0 and len(input_dir)>0:
        if values['xlsxToStrings'] == True:
            Log.info('xlsxToStrings')
            XS.startConvertXlsxToStrings(input_dir,output_dir)
            # XS.startConvertXlsxToStrings(input_dir,output_dir+'/strs_sources')
        elif values['stringsToXlsx'] == True:
            Log.info('stringsToXlsx')
            SX.startConvertStringToXlsx(input_dir,output_dir+'/xlsx_sources')
        elif values['xlsxToxml'] == True:
            Log.info('xlsxToxml')
            Xlsx2Xml.convertFromSingleForm(input_dir,output_dir+'/xml_sources')
        elif values['xmlToXlsx'] == True:
            Log.info('xmlToXlsx')
            Xml2Xlsx.convertToSingleFile(input_dir,output_dir+'/xlsx_sources')
        else:
            Log.info('要转换成什么格式？')

        

window.close()
