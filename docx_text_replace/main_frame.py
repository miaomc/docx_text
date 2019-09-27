# -*- coding: cp936 -*-
import Tkinter as tk
import shouquan_fuwu_tk
    

jiafang_name = u'甲方（XX单位）'
xiangmu_name = u'项目名称（湖南省XX项目（采购号XX））'
jichengshang_name = u'集成商（湖南XX科技公司）'
jichengshang_dizhi = u'集成商地址（芙蓉区XX楼XX号XX栋，XX房）'
time_weibao = u'维保时间（2）'
time_yyyy = u'年（2018）'
time_mm = u'月（7）'
time_dd = u'日（22）'

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
mainframe.title(u'一键授权&售后函工具V2.1')
#标签控件，显示文本和位图，展示在第一行

entry_list = []
for n,i in enumerate(lst_mannaul):
    label = tk.Label(mainframe,text=replace_dir[i]+':').grid(row=n*2,sticky=tk.E)#靠右
    tmp = tk.Entry(mainframe, width=45)
    tmp.grid(row=n*2,column=1)

    entry_list.append(tmp)
    
label = tk.Label(mainframe,text='      0          ').grid(row=len(lst_mannaul)*2+2,sticky=tk.W)#靠右

making_times = 0

def do_replace():
    global making_times
    making_times += 1
    label = tk.Label(mainframe,text='    making..').grid(row=len(lst_mannaul)*2+2,sticky=tk.W)#靠右
    for n,i in enumerate(lst_mannaul):
        replace_dir[i] = entry_list[n].get()#!!!!!!h01768.decode(sys.stdin.encoding)
    
    shouquan_fuwu_tk.shouquan_fuwu_replace(replace_dir)
    label = tk.Label(mainframe,text='      '+str(making_times)+'  ok.     ').grid(row=len(lst_mannaul)*2+2,sticky=tk.W)#靠右
    print 'ok'


tk.Button(mainframe,text=u'生成服务&授权函',width=15,height=2,command=do_replace).grid(row=len(lst_mannaul)*2+2, column=1)

mainframe.mainloop()
