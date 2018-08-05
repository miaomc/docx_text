# -*- coding: cp936 -*-
import win32com.client
import os
import docx_test2

def doc2docx(old_file, new_file):
    app=win32com.client.Dispatch('Word.Application')
    doc=app.Documents.Open(old_file)
    doc.SaveAs(new_file, 16)
    doc.Close()
    app.Quit()

##def search_file():
##    import os
##    import docx
##    for filename in os.listdir(os.getcwd()):
##        if filename.endswith('.docx'):
##            print(filename[:-5])
##            doc = docx.Document(filename)
##            for para in doc.paragraphs:
##                print (para.text)

def search_file(search_path = r'F:\分销产品彩页'):
    lst_caiye = []
    for path, dirs, files in os.walk(search_path):
        for f in files:
             lst_caiye.append(os.path.join(path, f))

    return lst_caiye

if __name__ == '__main__':
##    old_file = r'E:\投标资料模板\py_replace\612.doc'
##    new_file = r'E:\投标资料模板\py_replace\B612.docx'
    l = search_file()

    # make all *.doc SaveAs *.docx 
    for i in l:
        if i.endswith('.doc'):
            doc2docx(i,i+'x')



    



