git remote add origin https://github.com/miaomc/docx_text.git 
git rm '....py'
git commit -m "..."
git config --global http.postBuffer 524288000

һ��*.exe �汾
0��������shouquan_fuwu.py��Ҫ����ԭʼ��*.docx���ļ�û���ϴ�(u'���쳧����Ȩ����Uniview����ʱЧ������20160101_template.docx'�� u'�ۺ�����ŵģ�塾Uniview��20160901_template.docx')
1����Ҫ����docx�е�paragraphs��tables������tables���һ��cell������Ҳ�Ǹ�paragraphs
2��ͨ��paragraph.runs������ģ����ȷ����ʽ
3��tmp.pyΪ�ݸ�
4��paragraphs��tablesֻɨ����滻��һ�㣬���������һ���ݹ飬overstack���и��ϸ�����ӣ��������˸�ʽ���⣬�´����

����tkinter GUI �汾
1�� mainframe.py �� shouquan_fuwu_tk.py

P.S.
1��ע����python-docx������docx
2���޸�Դ
pip install python-docx -i https://mirrors.aliyun.com/pypi/simple/
3��ע����pyinstaller������pyinstall��pip install PyInstaller -i https://mirrors.aliyun.com/pypi/simple/
4��pyinstaller -F shouquan_fuwu.py
5��pyinstaller -F main_frame.py --noconsole


