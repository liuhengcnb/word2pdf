# -*- coding: utf-8 -*-
"""
功能：1.将文件夹rootpath中的非word文档复制到outputfolder文件夹下；
     2.将文件夹rootpath中的word文档转为同名pdf保存到outputfolder文件夹下。     
使用：修改rootpath后点击运行。
@author: lh
"""
import win32com.client, os, shutil

class Word_2_PDF(object):

    def __init__(self, filepath, Debug=False):      #param Debug: 控制过程是否可视化                
        self.wordApp = win32com.client.Dispatch('word.Application')
        self.wordApp.Visible = Debug
        self.myDoc = self.wordApp.Documents.Open(filepath)

    def export_pdf(self, output_file_path):     #将Word文档转化为PDF文件         
        self.myDoc.ExportAsFixedFormat(output_file_path, 17, Item=7, CreateBookmarks=0)

if __name__ == '__main__':
    
    '''↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓只修改此处rootpath↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓'''
    rootpath = r'C:\测试'       # 文件夹根目录
    newfolder_name = rootpath[rootpath.rfind('\\')+1:] + 'pdf'
    outputfolder = os.path.join(rootpath, newfolder_name)       #输出文件夹路径，路径为原文件夹下新创建的'原文件名+pdf'文件夹 
    os.mkdir(outputfolder)
        
    filelist = os.listdir(rootpath)
    docfilelist = [i for i in filelist if (i.endswith('doc') or i.endswith('docx'))]
    notdocfilelist = list(set(filelist)-set(docfilelist))
    notdocfilelist.remove(newfolder_name)
        
    for eachfilename in notdocfilelist:     #复制非word文档文件到新文件夹             
        shutil.copyfile(os.path.join(rootpath, eachfilename), os.path.join(outputfolder, eachfilename))
      
    for eachdocname in docfilelist:
        w2p = Word_2_PDF(os.path.join(rootpath, eachdocname), False)        
        eachpdfname = eachdocname[:eachdocname.rfind('.')]+'.pdf'
        w2p.export_pdf(os.path.join(outputfolder, eachpdfname))
        w2p.myDoc.Close()

    w2p.wordApp.Quit()
    print('转换完成！')
  
    
