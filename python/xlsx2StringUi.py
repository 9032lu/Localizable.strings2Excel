import PySimpleGUI as sg
from Xls2Strings import Xlsx2Strings as XS
from Strings2Xls import Strings2Xlsx as SX
import colorsys
from Log import Log
lista=['xlsxToStrings','stringsToXlsx']
layout=[
    [sg.FolderBrowse('选择文件夹',size=(8,1)),sg.In(key='-input_file_dir-')],
    [sg.FolderBrowse('保存到',size=(8,1)),sg.In(key='-output_file_dir-')],
    [sg.R(i,group_id=1,key=i) for i in lista],
    [sg.B('执行',size=(15,1))],
    [sg.ML(size=(65,10),reroute_cprint=True)],
   
]
window = sg.Window("python",layout)
while True:
    event,values=window.read()
    if  values== None:
        break
    input_dir = values['-input_file_dir-']
    output_dir = values['-output_file_dir-']
    print(input_dir)
    if len(input_dir)==0:
        Log.info('请选择源文件夹')
    elif len(output_dir)==0:
        Log.info('请选择保存文件夹')
    # if len(output_dir)>0:
    #     output_dir =  output_dir +'/sources/'
    print(len(output_dir))
    if event == None:
        break
    if event=='执行' and len(output_dir)>0 and len(input_dir)>0:
        if values['xlsxToStrings'] == True:
            Log.info('xlsxToStrings')
            XS.startConvertXlsxToStrings(input_dir,output_dir+'/strs_sources')
        elif values['stringsToXlsx'] == True:
            Log.info('stringsToXlsx')
            SX.startConvertStringToXlsx(input_dir,output_dir+'/xlsx_sources')
        else:
            Log.info('要转换成什么格式？')

        

window.close()
