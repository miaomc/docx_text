# -*- coding: cp936 -*-
import Tkinter as tk
import shouquan_fuwu_tk
    

jiafang_name = u'�׷���XX��λ��'
xiangmu_name = u'��Ŀ���ƣ�����ʡXX��Ŀ���ɹ���XX����'
jichengshang_name = u'�����̣�����XX�Ƽ���˾��'
jichengshang_dizhi = u'�����̵�ַ��ܽ����XX¥XX��XX����XX����'
time_weibao = u'ά��ʱ�䣨2��'
time_yyyy = u'�꣨2018��'
time_mm = u'�£�7��'
time_dd = u'�գ�22��'

replace_dir = {u'JIAFANG':jiafang_name,
               u'XIANGMU':xiangmu_name,
               u'JICHENGSHANG_NAME':jichengshang_name,
               u'JICHENGSHANG_DIZHI':jichengshang_dizhi,
               u'WEIBAO':time_weibao,
               u'TIME_YYYY':time_yyyy,
               u'TIME_MM':time_mm,
               u'TIME_DD':time_dd}
lst_mannaul = [u'JIAFANG',
               u'XIANGMU',
               u'JICHENGSHANG_NAME',
               u'JICHENGSHANG_DIZHI',
               u'WEIBAO',
               u'TIME_YYYY',
               u'TIME_MM',
               u'TIME_DD']

mainframe = tk.Tk()
mainframe.title(u'һ����Ȩ&�ۺ󺯹���V2.1')
#��ǩ�ؼ�����ʾ�ı���λͼ��չʾ�ڵ�һ��

entry_list = []
for n,i in enumerate(lst_mannaul):
    label = tk.Label(mainframe,text=replace_dir[i]+':').grid(row=n*2,sticky=tk.E)#����
    tmp = tk.Entry(mainframe, width=45)
    tmp.grid(row=n*2,column=1)

    entry_list.append(tmp)
    
label = tk.Label(mainframe,text='      0          ').grid(row=len(lst_mannaul)*2+2,sticky=tk.W)#����

making_times = 0

def do_replace():
    global making_times
    making_times += 1
    label = tk.Label(mainframe,text='    making..').grid(row=len(lst_mannaul)*2+2,sticky=tk.W)#����
    for n,i in enumerate(lst_mannaul):
        replace_dir[i] = entry_list[n].get()#!!!!!!h01768.decode(sys.stdin.encoding)
    
    shouquan_fuwu_tk.shouquan_fuwu_replace(replace_dir)
    label = tk.Label(mainframe,text='      '+str(making_times)+'  ok.     ').grid(row=len(lst_mannaul)*2+2,sticky=tk.W)#����
    print 'ok'


tk.Button(mainframe,text=u'���ɷ���&��Ȩ��',width=15,height=2,command=do_replace).grid(row=len(lst_mannaul)*2+2, column=1)

mainframe.mainloop()
