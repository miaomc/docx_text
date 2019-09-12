git remote add origin https://github.com/miaomc/docx_text.git 
git rm '....py'
git commit -m "..."
git config --global http.postBuffer 524288000

一、*.exe 版本
0、主程序shouquan_fuwu.py需要两个原始的*.docx的文件没有上传(u'制造厂商授权函【Uniview】无时效性限制20160101_template.docx'和 u'售后服务承诺模板【Uniview】20160901_template.docx')
1、需要考虑docx中的paragraphs和tables，其中tables最后一层cell的内容也是个paragraphs
2、通过paragraph.runs的内联模块来确保格式
3、tmp.py为草稿
4、paragraphs和tables只扫描和替换了一层，这里可能少一个递归，overstack上有个严格的例子，还讲到了格式问题，下次再深究

二、tkinter GUI 版本
1、 mainframe.py 和 shouquan_fuwu_tk.py

P.S.
1、注意是python-docx而不是docx
2、修改源
pip install python-docx -i https://mirrors.aliyun.com/pypi/simple/
3、注意是pyinstaller而不是pyinstall，pip install PyInstaller -i https://mirrors.aliyun.com/pypi/simple/
4、pyinstaller -F shouquan_fuwu.py
5、pyinstaller -F main_frame.py --noconsole


