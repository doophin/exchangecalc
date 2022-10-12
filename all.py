 # -*- coding: utf-8 -*-
"""
Created on Thu Nov 11 19:56:30 2021

@author: koi
"""
import sys
from PyQt5.QtWidgets import QPushButton,QWidget,QLabel,QGridLayout,QApplication,QLineEdit,QInputDialog,QMessageBox,QCheckBox,QMainWindow
from PyQt5.QtGui import QIcon
import pandas as pd 
import datetime
from dateutil.relativedelta import relativedelta
import os
import matplotlib as mpl
import seaborn as sns
from operator import itemgetter
import math
# import traceback

mpl.rcParams['font.sans-serif'] = ['KaiTi']
mpl.rcParams['font.serif'] = ['KaiTi']
mpl.rcParams['axes.unicode_minus'] = False # 解决保存图像是负号'-'显示为方块的问题,或者转换负号为字符串
sns.set_style("darkgrid",{"font.sans-serif":['KaiTi', 'Arial']})   #这是方便seaborn绘图得时候得字体设置
pd.set_option('mode.chained_assignment', None)

class Wintool(QMainWindow):
    def __init__(self):
        super().__init__()
        ckpasswd,ok = QInputDialog.getText(self,'登录','请输入密码')
        if ok is False:
            sys.exit(0)   
        elif ckpasswd == 'msqh':
            self.initTool()
        else:
            QMessageBox.critical(self, '错误',"密码错误")
            sys.exit(0)

    def initTool(self):
        self.cname = QLabel('请输入柜台表文件名日期')
        self.cnameEdit = QLineEdit()
        now = datetime.datetime.today().__format__('%Y%m%d')
        self.cnameEdit.setText(now)
        self.wzmdate = now
        self.cnameEdit.textChanged[str].connect(self.wzmdatefunc)

        self.btn1 = QPushButton('确认')    
        self.btn1.setShortcut("Return")
        self.btn1.clicked.connect(self.checkevent)
        
        self.t1text =  QLabel()
        self.t2text =  QLabel() 
        self.t3text =  QLabel()
        self.t4text =  QLabel()
        self.t5text =  QLabel()
        self.t6text =  QLabel()
        
        self.t1ck = QLabel()
        self.t2ck = QLabel()
        self.t3ck = QLabel()
        self.t4ck = QLabel()
        self.t5ck = QLabel()
        self.t6ck = QLabel()
        
        self.myinfo = QPushButton('说明')
        self.myinfo.clicked.connect(self.checkinfo)
        
        self.check1 = QCheckBox('测算郑商所', self)
        self.check1.setChecked(True)
        self.czceflag = 1
        self.check1.stateChanged.connect(self.choose)
        
        self.check2 = QCheckBox('测算上期所', self)
        self.check2.setChecked(True)  
        self.shfeflag = 1
        self.check2.stateChanged.connect(self.choose)
        
        self.check3 = QCheckBox('测算能源中心', self)
        self.check3.setChecked(True)
        self.ineflag = 1
        self.check3.stateChanged.connect(self.choose)
        
        self.check4 = QCheckBox('测算大商所', self)
        self.check4.setChecked(True)
        self.dceflag = 1
        self.allflag = 4    
        self.check4.stateChanged.connect(self.choose)
        
        self.check5 = QCheckBox('生成持仓图', self) 
        self.imgflag = 0    
        self.check5.stateChanged.connect(self.choose)
        
        self.check6 = QCheckBox('生成明细表', self)
        self.chosflag = 0 
        self.check6.stateChanged.connect(self.choose)
        
        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(self.cname,1,1)
        grid.addWidget(self.cnameEdit,1,2)

        grid.addWidget(self.btn1,1,3)
        grid.addWidget(self.t1text,2,1)
        grid.addWidget(self.t2text,3,1)
        grid.addWidget(self.t3text,4,1)
        grid.addWidget(self.t4text,5,1)
        grid.addWidget(self.t5text,6,1)
        grid.addWidget(self.t6text,7,1)
        
        grid.addWidget(self.t1ck,2,2)
        grid.addWidget(self.t2ck,3,2)
        grid.addWidget(self.t3ck,4,2)
        grid.addWidget(self.t4ck,5,2)
        grid.addWidget(self.t5ck,6,2)
        grid.addWidget(self.t6ck,7,2)
        
        grid.addWidget(self.check1,2,3)
        grid.addWidget(self.check2,3,3)
        grid.addWidget(self.check3,4,3)
        grid.addWidget(self.check4,5,3)
        grid.addWidget(self.check5,6,3)
        grid.addWidget(self.check6,7,3)
        grid.addWidget(self.myinfo,8,3)
        
        widget = QWidget()
        widget.setLayout(grid)
        
        if getattr(sys, 'frozen', False): #是否Bundle Resource
            icofile = os.path.join(sys._MEIPASS,'b.ico')
        else:
            icofile = ('b.ico')
        self.setWindowIcon(QIcon(icofile))
        self.setCentralWidget(widget)
        self.setWindowTitle('交易所测算')
        self.show()
        self.resize(600, 250)
        
    def checkinfo(self):
        QMessageBox.information(self, '说明',"请在当前table目录下放入从CTP主席柜台导出的六张CSV格式表\n产品设置\n交易所期货结算保证金率指定日查询\n交易所结算明细\n历史结算价查询\n投资者资料查询\n投资者持仓查询\n\
                                \n在文本框中输入六张表后的日期，即导出表当天的日期\n点击触认后会对表存在进行检查。如果六张表存在，请在弹出的文本框中输入客户持仓表中最后的交易日期。若是今日结算后导出的数据，文本框可为空，日期会默认为今天日期\
                                \n勾选生成明细表会多生成一张包列计算的所有列明细表\
                                \n版本：3.0\
                                \n                                                                 Update:2022-01-13\n                                                                 Author:Koi")
    
    def choose(self):            
        if self.check1.isChecked():
            self.czceflag = 1
        else:
            self.czceflag = 0  
            
        if self.check2.isChecked():
            self.shfeflag = 1
        else:
            self.shfeflag = 0  
            
        if self.check3.isChecked():
            self.ineflag = 1
        else:
            self.ineflag = 0  
        
        if self.check4.isChecked():
            self.dceflag = 1
        else:
            self.dceflag = 0      
        
        if self.check5.isChecked():
            self.imgflag = 1
        else:
            self.imgflag = 0 
            
        if self.check6.isChecked():
            self.chosflag = 1
        else:
            self.chosflag = 0
            
        
        self.allflag = self.czceflag + self.shfeflag + self.ineflag + self.dceflag

    def wzmdatefunc(self,a):
        self.wzmdate = a
        
    def checkevent(self, event):
        if self.allflag == 0:
            QMessageBox.warning(self, '错误',"未选择交易所")
        else:
            self.lz = 0
            self.frflag = 0
            self.t6text.clear()
            self.t6ck.clear()
 
            tzzcccx = cs(self.wzmdate).tzzcccxt           
            jysjsmx = cs(self.wzmdate).jysjsmxt
            cpsz = cs(self.wzmdate).cpszt          
            lsjsjcx = cs(self.wzmdate).lsjsjcxt
            bzjl = cs(self.wzmdate).bzjlt
            
            tzzzlcx = cs(self.wzmdate).tzzzlcxt
        
            self.t1text.setText(bzjl)  
            self.t2text.setText(tzzcccx)
            self.t3text.setText(jysjsmx)
            self.t4text.setText(cpsz)      
            self.t5text.setText(lsjsjcx)
            
            self.existcz(self.t1ck,os.path.exists(bzjl))    
            self.existcz(self.t2ck,os.path.exists(tzzcccx))
            self.existcz(self.t3ck,os.path.exists(jysjsmx))
            self.existcz(self.t4ck,os.path.exists(cpsz))
            self.existcz(self.t5ck,os.path.exists(lsjsjcx))
        
            if self.czceflag == self.dceflag == 0:
                if self.lz == 5:
                    self.lz = 'ok'           
            else:
                self.t6text.setText(tzzzlcx)
                self.existcz(self.t6ck,os.path.exists(tzzzlcx))
                self.frflag = 1
                if self.lz == 6:
                    self.lz = 'ok'
  
            self.adjustSize()
            if self.lz == 'ok':
                tddate,ok = QInputDialog.getText(self,'检查通过','请输入最后交易日期',text=self.wzmdate)
                if ok and tddate == '':
                    QMessageBox.warning(self, '注意',"日期将设置为今天日期")
                    self.jyscs(self.wzmdate,tddate,self.czceflag,self.shfeflag,self.ineflag,self.dceflag,self.chosflag,self.imgflag,self.frflag)
                elif ok:
                    self.jyscs(self.wzmdate,tddate,self.czceflag,self.shfeflag,self.ineflag,self.dceflag,self.chosflag,self.imgflag,self.frflag)
               
    def jyscs(self,a,b,c,d,e,f,g,h,i):      
        try:
            cs(a).operjs(b,c,d,e,f,g,h,i)
            QMessageBox.about(self, '完成',"生成完毕！")
        except Exception as e:
            QMessageBox.critical(self, '错误',"生成错误\n\n>>>>>"+str(e))
        # except:
            # QMessageBox.critical(self, '错误',"生成错误\n\n>>>>>"+str(traceback.format_exc()))

    def existcz(self,a,b):
        if b == True:
            flag = '文件存在'
            self.lz = self.lz + 1
        else:
            flag = '文件不存在'
        a.setText(flag)   

class cs():
    def toxlsx(self,a,b,c):         
        a.to_excel(b,index = False,freeze_panes=(1,2),sheet_name=c)
    def wzsz(self,a):                                                               
        tpath = os.path.join(os.getcwd(),'table')
        a = datetime.datetime.strptime(a, '%Y%m%d').__format__('-%Y-%m-%d')
        self.tzzcccxt = os.path.join(tpath,'投资者持仓查询'+a+'.CSV')
        self.tzzzlcxt = os.path.join(tpath,'投资者资料查询'+a+'.CSV')
        self.bzjlt = os.path.join(tpath,'交易所期货结算保证金率指定日查询'+a+'.CSV')
        self.cpszt = os.path.join(tpath,'产品设置'+a+'.CSV')
        self.jysjsmxt = os.path.join(tpath,'交易所结算明细'+a+'.CSV')
        self.lsjsjcxt = os.path.join(tpath,'历史结算价查询'+a+'.CSV') 
    
    def operjs(self,a,b,c,d,e,f,g,h):
        self.td = a
        self.czceflag = b
        self.shfeflag = c
        self.ineflag = d
        self.dceflag = e
        self.chosflag = f 
        self.imgflag = g
        self.frflag = h
        
        self.cinit()
  #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if self.dceflag == 1:
            self.exchoose = 'DCE'
            self.exchange = '大商所'
            self.cominit() 
            self.company()
            if self.imgflag == 1:
                self.drawimg(self.tzzcccxqh)
            self.framework(self.tzzzlcxt)
            self.comqz = self.comqz.rename(columns={'保证金':'今日公司保证金占用'})
            ssf=self.dsjysjsmx     
            bzj=self.bzjl
            bili=pd.read_csv('config/bili.csv',encoding="GBK", thousands=',',low_memory=False)#                      
            jizhun= pd.read_csv('config/jizhun.csv',encoding="GBK", low_memory=False)#取基准                     
            settlementp = pd.merge(jizhun,self.zlhy,on = '产品')[['产品','合约','结算价']]           
            frcc1= self.tzzcccxqh[['结算日','投资者代码','产品','合约','手数','客户性质']]#只提取需要内容            
            frcc2=frcc1.loc[frcc1['客户性质']=='法人']
            zrrcc=frcc1.loc[frcc1['客户性质']=='自然人']           
            maxday = int(self.td)            
            tds = self.totalday          
            donetds = self.traday           
            undonetds = self.leftday
            nexttds=self.nextotalday
            tdfrcc=frcc2.loc[frcc2['结算日']==maxday]
            tdfrcchz=tdfrcc.groupby('产品').agg({'手数':'sum'}).reset_index()
            tdfrcchz=tdfrcchz.rename(columns={'手数':'法人今日持仓'})
            frcchz=frcc2.groupby('产品').agg({'手数':'sum'}).reset_index()
            frcchz=frcchz.rename(columns={'手数':'法人持仓汇总'})
            frcchz=pd.merge(frcchz,tdfrcchz,on='产品',how='outer').fillna(0)
            frcchz['法人持仓汇总加权']=frcchz['法人持仓汇总']*1.5
            frcchz['法人今日持仓加权']=frcchz['法人今日持仓']*1.5
            tdzrrcc=zrrcc.loc[zrrcc['结算日']==maxday]
            tdzrrcchz=tdzrrcc.groupby('产品').agg({'手数':'sum'}).reset_index()
            tdzrrcchz=tdzrrcchz.rename(columns={'手数':'自然人今日持仓'})
            zrrcchz=zrrcc.groupby('产品').agg({'手数':'sum'}).reset_index()
            zrrcchz=zrrcchz.rename(columns={'手数':'自然人持仓汇总'})
            zrrcchz=pd.merge(zrrcchz,tdzrrcchz,on='产品',how='outer').fillna(0)
            basic=pd.merge(frcchz,zrrcchz,on='产品',how='outer').fillna(0)
            basic=pd.merge(basic,settlementp,on='产品',how='outer').fillna(0)
            basic=pd.merge(basic,ssf,on='产品',how='inner').fillna(0)
            basic=pd.merge(basic,self.comqz,on='产品',how='left').fillna(0)
            basic['持仓汇总']=basic['法人持仓汇总']+basic['自然人持仓汇总']
            basic['今日持仓']=basic['法人今日持仓']+basic['自然人今日持仓']
            basic['加权持仓汇总']=basic['法人持仓汇总加权']+basic['自然人持仓汇总']
            basic['加权今日持仓']=basic['法人今日持仓加权']+basic['自然人今日持仓']
            basic['总交易天数']=tds
            basic['已交易天数']=donetds
            basic['未交易天数']=undonetds
            basic['预计持仓']=[(x+y*z) for (x,y,z) in zip (basic['持仓汇总'],basic['今日持仓'],basic['未交易天数'])]
            basic['加权预计持仓']=[(x+y*z) for (x,y,z) in zip (basic['加权持仓汇总'],basic['加权今日持仓'],basic['未交易天数'])]
            basic['加权持仓日均']=basic['加权持仓汇总']/basic['已交易天数']
            basic['预计日均']=basic['预计持仓']/basic['总交易天数']
            basic['加权预计日均']=basic['加权预计持仓']/basic['总交易天数']
            jizhun=pd.merge(jizhun,bzj,on='产品',how='left')
            jizhun.fillna(0,inplace=True)#na转化为0
            jizhun = pd.merge(jizhun,self.cpsz,on = '产品',how='left')
            jizhun['保证金占用'] = jizhun['每手数量'] * jizhun['保证金率']
            basic=pd.merge(jizhun,basic,on='产品',how='outer').fillna(0)
            basic['非加权相差日均']=basic['基准']-basic['预计日均']
            basic['非加权相差手数双边']=[(x*y/self.nleftday) for (x,y) in zip (basic['非加权相差日均'],basic['总交易天数'])]
            # basic['非加权相差手数双边']=[(x*y/z) for (x,y,z) in zip (basic['非加权相差日均'],basic['总交易天数'],basic['未交易天数'])]
            self.anadce(basic)
            basic = self.atmp
            basic['可达到减收比例']=0
            basic['可达到减收比例']=[b if a>=1 else c for (a,b,c) in zip (basic['A类完成百分比'],basic['A类减收比例'],basic['可达到减收比例'])]
            basic['可达到减收比例']=[b if a>=1 else c for (a,b,c) in zip (basic['B类完成百分比'],basic['B类减收比例'],basic['可达到减收比例'])]
            basic['可达到减收比例']=[b if a>=1 else c for (a,b,c) in zip (basic['C类完成百分比'],basic['C类减收比例'],basic['可达到减收比例'])]
            basic['可达到减收比例']=[b if a>=1 else c for (a,b,c) in zip (basic['D类完成百分比'],basic['D类减收比例'],basic['可达到减收比例'])]
            basic['预计减收手续费']=basic['上交手续费']*basic['可达到减收比例']
            basic['预计减收手续费收益']=[(a/b*c) for (a,b,c) in zip(basic['预计减收手续费'],basic['已交易天数'],basic['总交易天数'])]
            basic['已持仓收益率']=[(x/y*365/30) if y!=0 else 0 for (x,y) in zip(basic['预计减收手续费收益'],basic['今日公司保证金占用'])]
            basic['条件A非加权']=[(x-(y-z)) for (x,y,z) in zip (basic['基准'],basic['预计日均'],basic['预计考核月内公司占有日均(双边)'])]
            basic['条件B加权']=[((x-(y-1.5*z))/1.5) for (x,y,z) in zip (basic['D类达标日均'],basic['加权预计日均'],basic['预计考核月内公司占有日均(双边)'])]
            basic['下月预计相差手数双边']= [y if x < y else x for (x,y) in zip (basic['条件A非加权'],basic['条件B加权'])]
            basic['下月预计相差手数单边']=[math.ceil(x/2) for x in  basic['下月预计相差手数双边']]
            basic['下月资金预计占用(万元)']=[(x*y*z/10000) for (x,y,z) in zip (basic['结算价'],basic['保证金占用'],basic['下月预计相差手数单边'])]
            basic = basic.set_index('产品')
            basic.loc['生猪','下月资金预计占用(万元)'] = basic.loc['生猪','下月资金预计占用(万元)']*2
            basic = basic.reset_index()
            basic['下月预计减收手续费']=[(x*y*nexttds/tds) for (x,y) in zip (basic['上交手续费'],basic['D类减收比例'])]
            basic['下月预计收益率']=[(x/(y*10000)*365/30) if y!=0 else 0 for (x,y) in zip(basic['下月预计减收手续费'],basic['下月资金预计占用(万元)'])]
            basic['今日公司保证金占用'] = basic['今日公司保证金占用']/10000
            basic = basic.rename(columns={'今日公司保证金占用':'今日公司保证金占用(万元)'})
            basic['达D档增持手数（单边）']=basic['D类达标相差手数单边']
            basic['增持需要保证金（万元）']=basic['D类资金预计占用']/10000            
            basic['增持年化收益率']=[((x-y)*365/30/z) for (x,y,z) in zip (basic['D类收益'],basic['预计减收手续费收益'],basic['D类资金预计占用'])]
            basic['自有资金已持仓手数（单边）']=basic['公司今日占有(双边)']*0.5
            basic['已投入保证金(万元)']=basic['今日公司保证金占用(万元)']
            basic['已投入年化收益率']=basic['已持仓收益率']
            basic['预计达标总持仓手数']=[b if a<0 else (a+b) for (a,b) in zip (basic['达D档增持手数（单边）'],basic['自有资金已持仓手数（单边）'])]
            basic['预计达标总保证金']=[b if a<0 else (a+b) for (a,b) in zip (basic['增持需要保证金（万元）'],basic['今日公司保证金占用(万元)'])]
            basic['预计达标年化收益率']=[(x*365/30/(y*10000)) if y!=0 else 0 for (x,y) in zip (basic['D类收益'],basic['预计达标总保证金'])]
            #final=basic
            final=basic.loc[:,['产品','加权持仓汇总','加权持仓日均','加权预计日均','上交手续费','预计减收手续费','预计减收手续费收益','加权基准','基准','非加权相差日均','A类达标日均','B类达标日均','C类达标日均','D类达标日均','A类完成百分比','B类完成百分比','C类完成百分比','D类完成百分比','A类减收手续费','B类减收手续费','C类减收手续费','D类减收手续费','加权今日持仓','今日持仓','合约','结算价','D类达标相差手数双边','D类达标相差手数单边','可达到减收比例','预计减收手续费','公司今日占有(双边)','今日公司保证金占用(万元)','达D档增持手数（单边）','增持需要保证金（万元）','增持年化收益率','自有资金已持仓手数（单边）','已投入保证金(万元)','已投入年化收益率','预计达标总持仓手数','预计达标总保证金','预计达标年化收益率','条件A非加权','条件B加权','下月预计相差手数双边','下月预计相差手数单边','下月资金预计占用(万元)','下月预计减收手续费','下月预计收益率']]
            c=['A类完成百分比','B类完成百分比','C类完成百分比','D类完成百分比','可达到减收比例','下月预计收益率','增持年化收益率','已投入年化收益率','预计达标年化收益率']  
            def zh(d):
                final[d] = final[final[d]!=''][d].astype(float).apply(lambda x: format(x,'.2%'))
            for n in c:
                zh(n)
            x=len(final.index)#取行数
            y=len(bili.columns)#取列数
            ts = ["总交易天数","已交易天数","未交易天数","下月交易天数"]
            sz= [(x) for x in [tds,donetds,undonetds,nexttds]]
            list_to_tuple = list(zip(ts,sz))
            tssm= pd.DataFrame(list_to_tuple,columns=["ts","sz"])#生成天数dataframe
            name = self.exchange+self.cnday+'减收测算'+ self.moned+'.xlsx'
            with pd.ExcelWriter(os.path.join(self.path,self.cnday,self.exchange,name)) as writer:
                final.to_excel(writer,index=False)
                bili.to_excel(writer, startrow=x+1, header=None, index=False)
                tssm.to_excel(writer, startrow=x+1,startcol=y,header=None, index=False,freeze_panes=(1,1))

  #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++          
        if self.czceflag == 1:
            self.exchoose = 'CZCE'  
            self.exchange = '郑商所'
            th1 = '郑商所月末'
            th2 = '郑商所原始'
            czcewj = pd.read_excel(self.exchangwj,sheet_name= self.czcewjt)
            self.cominit() 
            self.company()
            if self.imgflag == 1:
                self.drawimg()
            hy = self.tzzcccxqh.groupby(['产品','合约']).sum('手数').reset_index()[['产品','合约','手数']]
            zdhy = pd.merge(self.zdhyfr,hy,on='合约',how='left')[['产品','合约','手数']].set_index('合约').dropna()
            
            tdhy = self.tdtzzcccxqh.groupby(['产品','合约']).sum('手数').reset_index()[['产品','合约','手数']].set_index('合约')
            zdhy['当天重点合约平均手数(双边)'] = tdhy['手数']
            zdhy = zdhy.fillna(0).reset_index()
            zdhy['预计重点合约平均总持仓(双边)']  = zdhy['当天重点合约平均手数(双边)']*self.leftday + zdhy['手数']
            zdhy['预计重点合约平均日均持仓(双边)'] = zdhy['预计重点合约平均总持仓(双边)']/self.totalday

            zdhy = zdhy.groupby('产品').mean().reset_index()
            zdhy = pd.merge(zdhy,self.zdhynum,how='left',on='产品').rename(columns={'手数':'当前重点合约平均手数(双边)'})
            zdhy['重点合约数量'] = zdhy['重点合约数量'].astype(int)
            self.zdhy = zdhy
            
         
 
            
            base = pd.merge(czcewj.fillna(''),self.qhpz,how='left',on='产品').fillna(0)
            base = pd.merge(base,self.zlhy,on='产品',how='left')
            # base = pd.merge(base,zdhy,on='产品',how='left').fillna('')
            self.comother(base)
            base = self.atmp  
            self.czcefr()
            base = pd.merge(base,self.frzb,on='产品',how='left').fillna(0)
            base = pd.merge(base,self.comqz,on='产品',how='left').fillna(0)
            
            base['保证金'] = base['保证金']/10000
            base = base.rename(columns={'保证金':'当前已投入(万元)'})
            self.czcedbjs(base)
            self.base = self.atmp
            
            
            # self.base['去除公司占有达标满减差额日均'] = self.base['满减所需日均手数'] - self.base['预计日均持仓（加权双边）'] + self.base['预计考核月内公司占有日均(双边加权)']
            # self.base['权重'] = self.base['重点合约'].apply(lambda x:2 if x != ''  else 1)
            # self.base['去除公司占有达标满减差额日均'] = self.base['去除公司占有达标满减差额日均'] / self.base['权重']
            # self.base['达标满减需投入(万元)'] = self.base['去除公司占有达标满减差额日均'] * self.base['每手保证金'] / 20000
            # self.base['去除公司占有今日法人持仓'] = self.base['今日法人持仓(双边)'] - self.base['公司今日占有(双边)']
            # self.base['下月要求标准'] = self.base['当月法人日均持仓占比最低标准'] + 0.01
            # self.base['法人下一日需开仓(今日计算)'] = (self.base['下月要求标准'] * ( self.base['今日持仓'] - self.base['公司今日占有(双边)']) - self.base['去除公司占有今日法人持仓'] )/  (1 - self.base['下月要求标准']) 
            # self.base['法人达标需投入(今日计算)(万元)'] = self.base['法人下一日需开仓(今日计算)']  *  self.base['每手保证金'] / 20000
            # self.base['当前法人日均'] = self.base['当前法人总持仓(双边)'] / self.traday
            # self.base['去除公司占有法人日均'] = self.base['当前法人日均'] - self.base['当前日均(双边)']
            # self.base['法人下一日需开仓(日均计算)'] =  (self.base['下月要求标准'] * (self.base['预计当月日均持仓'] - self.base['当前日均(双边)']) - self.base['去除公司占有法人日均']) / (1 - self.base['下月要求标准']) 
            # self.base['法人达标需投入(日均计算)(万元)'] = self.base['法人下一日需开仓(日均计算)']  *  self.base['每手保证金'] / 20000
            
            # self.base['去除公司后达标满减预计年化'] = 365*self.base['达标减收手续费（元）']/10000/( self.base['达标满减需投入(万元)'] *30)
            # self.base['去除公司后法人达标年化(今日计算)'] = 365*self.base['持仓及法人占比均达标可减收(元)']/10000/( self.base['法人达标需投入(今日计算)(万元)'] *30)
            # self.base['去除公司后法人达标年化(日均计算)'] = 365*self.base['持仓及法人占比均达标可减收(元)']/10000/( self.base['法人达标需投入(日均计算)(万元)'] *30)
            
            czcelist = ['达标完成率','上限达标年化收益率','今日法人持仓占比','当前法人持仓占比日均值','法人达标日均差值',\
                        '法人达比增投年化收益','法人达标完成率','去除公司占有上限达标年化收益率','重点合约上限达标年化收益率']

            self.reout(th1,th2,czcelist)
            


 #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++           
        if self.shfeflag == 1:
            self.exchoose = 'SHFE'
            self.exchange = '上期所'
            th1 = '上期所月末'
            th2 = '上期所原始'
            self.shcom(self.exchoose,self.shfezlwjt,self.shfedbwjt)
            a = self.base
            b = self.dbwj
            b['预计当月日均持仓总值'] = self.totalnum
            b['预计满足85%达标总日均还差'] = b['85%'] - b['预计当月日均持仓总值']
            b['预计达标85%下一日需共开仓'] = b['预计满足85%达标总日均还差'] * self.totalday/ self.nleftday
            if b['预计当月日均持仓总值'][0] >= b['85%'][0]:
                b['达标情况'] = '已达标'
                b['当前达标减收比例'] = b['大于等于85%']
            elif b['预计当月日均持仓总值'][0] < b['85%'][0] and b['预计当月日均持仓总值'][0] >= b['75%'][0]:
                b['达标情况'] = '已达标75%未达标85%'
                b['当前达标减收比例'] = b['小于85%大于等于75%']
            elif b['预计当月日均持仓总值'][0] < b['75%'][0] and b['预计当月日均持仓总值'][0] >= b['50%'][0]:
                b['达标情况'] = '已达标50%未达标75%'
                b['当前达标减收比例'] = b['小于75%大于等于50%']
            else:
                b['达标情况'] = '未达标50%'
                b['当前达标减收比例'] = b['小于50%']
            
            self.dbwz = b
            
            for t in range(0,len(a)): 
                if a['达到基准需下一日开仓'][t] != 0:
                    a.loc[t,'开仓权重'] =  min ((b['预计达标85%下一日需共开仓'][0] - a['达到基准需下一日开仓'][t]) * self.leftday / self.totalday  * a['大于基准每多一手可增加减收'][t] , a['增量减收可减最大值'][t] ) / a['每手保证金'][t]
                else:
                    a.loc[t,'开仓权重'] = 0
            mslist = a.sort_values(by = '开仓权重',ascending = False).index
            n = mslist[0]
            if a['开仓权重'][n] > 0:
                a.loc[n,'满足85%补产品仓位前仓数'] = b['预计达标85%下一日需共开仓'][0]
                if a['满足85%补产品仓位前仓数'][n] >= a['获得所有增量减收需下日开仓'][n]:
                    a.loc[n,'补仓方案'] = a['获得所有增量减收需下日开仓'][n]
                else:
                    a.loc[n,'补仓方案'] =  a['满足85%补产品仓位前仓数'][n]
                a.loc[n,'此产品补完后剩余需补仓数'] = a['满足85%补产品仓位前仓数'][n] - a['补仓方案'][n]
                a.loc[n,'补仓所需金额(万元)'] = a['补仓方案'][n]*a['每手保证金'][n]/20000
                a.loc[n,'补仓所获增收额度'] = (a['补仓方案'][n] - a['达到基准需下一日开仓'][n])*self.leftday/self.totalday*a['大于基准每多一手可增加减收'][n]
            else:
                a[['满足85%补产品仓位前仓数','补仓方案','此产品补完后剩余需补仓数','补仓所需金额(万元)','补仓所获增收额度']] = 0
            f = 1
            while 0 not in a['此产品补完后剩余需补仓数'].values.tolist():
                for t in range(f,len(mslist)):               
                    j = mslist[t]
                    k = mslist[f-1]
                    if a['达到基准需下一日开仓'][j] < a['此产品补完后剩余需补仓数'][k] and a['开仓权重'][j] > 0:              
                        a.loc[j,'满足85%补产品仓位前仓数'] = a['此产品补完后剩余需补仓数'][k]
                        if a['满足85%补产品仓位前仓数'][j] >= a['获得所有增量减收需下日开仓'][j]:
                            a.loc[j,'补仓方案'] = a['获得所有增量减收需下日开仓'][j]
                        else:
                            a.loc[j,'补仓方案'] =  a['满足85%补产品仓位前仓数'][j]
                        a.loc[j,'此产品补完后剩余需补仓数'] = a['满足85%补产品仓位前仓数'][j] - a['补仓方案'][j]
                        a.loc[j,'补仓所需金额(万元)'] = a['补仓方案'][j]*a['每手保证金'][j]/20000
                        a.loc[j,'补仓所获增收额度'] = (a['补仓方案'][j] - a['达到基准需下一日开仓'][j])*self.leftday/self.totalday*a['大于基准每多一手可增加减收'][j]
                        f = t + 1
                        break                  
                    else:
                        if t == len(mslist)-1:                         
                            if a[a['产品'] == '燃料油']['满足85%补产品仓位前仓数'][0] > 0 :
                                a.loc[0,'满足85%补产品仓位前仓数'] = a[a['产品'] == '燃料油']['满足85%补产品仓位前仓数'][0] + a['此产品补完后剩余需补仓数'][k]
                                a.loc[0,'补仓方案']  = a[a['产品'] == '燃料油']['补仓方案'][0] + a['此产品补完后剩余需补仓数'][k]
                                a.loc[0,'此产品补完后剩余需补仓数'] = 0
                                break
                            else:
                                a.loc[0,'满足85%补产品仓位前仓数'] = a['此产品补完后剩余需补仓数'][k]
                                a.loc[0,'补仓方案'] = a['此产品补完后剩余需补仓数'][k]
                                a.loc[0,'此产品补完后剩余需补仓数'] = 0
                                a.loc[0,'补仓所需金额(万元)'] = a['补仓方案'][0]*a['每手保证金'][0]/20000
                                a.loc[0,'补仓所获增收额度'] = 0
                                break
                        else:
                            continue          
            # =============================================================================
            if b['预计达标85%下一日需共开仓'][0] > 0:
                a['达标85%需保证金'] = b['预计达标85%下一日需共开仓'][0] * a['每手保证金'] / 20000
                for t in range(0,len(a)):
                    a.loc[t,'达标85%增加增量减收'] = min ((b['预计达标85%下一日需共开仓'][0] - a['达到基准需下一日开仓'][t]) * self.leftday / self.totalday  * a['大于基准每多一手可增加减收'][t] , a['增量减收可减最大值'][t] )
                # a['实施后可多返'] = a['达标85%增加增量减收'] + a['上交手续费'].sum() * (b['大于等于85%'][0] - b['当前达标减收比例'][0])
                a['实施后可多返'] = a['达标85%增加增量减收']
                a['达标85%年化收益率'] = 365*a['实施后可多返']/10000/(a['达标85%需保证金'] * 30)
            else:
                a[['达标85%需保证金','实施后可多返','达标85%年化收益率']] = 0
    
            self.base = a   
            shfelist = ['达标完成率','超标完成率','持仓超标增投年化收益','达标85%年化收益率','当前已持仓预计年化收益率','累积公司已投入持仓超标增投年化收益','预计年化']
            

            self.reout(th1,th2,shfelist)
            
 
  #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++                                
        if self.ineflag == 1:
            self.exchoose = '能源'
            self.exchange = '能源中心'
            th1 = '能源中心月末'
            th2 = '能源中心原始'
            self.shcom(self.exchoose,self.inezlwjt,self.inedbwjt)
            a = self.base
            b = self.dbwj
            b['预计当月日均持仓总值'] = self.totalnum
            b['预计满足60%达标总日均还差'] = b['60%'] - b['预计当月日均持仓总值']
            b['预计达标60%下一日需共开仓'] = b['预计满足60%达标总日均还差'] * self.totalday/ self.nleftday
            if b['预计当月日均持仓总值'][0] >= b['60%'][0]:
                b['达标情况'] = '已达标'
                b['当前达标减收比例'] = b['大于等于60%']
            else:
                b['达标情况'] = '未达标60%'
                b['当前达标减收比例'] = b['小于60%']
            b['公司日均占有总值'] = a['预计考核月内公司占有日均(双边)'].sum()
            b['去除公司占有满足60%日均还差'] = b['预计满足60%达标总日均还差'] + b['公司日均占有总值']
            self.dbwz = b      
            
            # =============================================================================
            for t in range(0,len(a)): 
                if a['达到基准需下一日开仓'][t] != 0:
                    a.loc[t,'开仓权重'] =  min ((b['预计达标60%下一日需共开仓'][0] - a['达到基准需下一日开仓'][t]) * self.leftday / self.totalday  * a['大于基准每多一手可增加减收'][t] , a['增量减收可减最大值'][t] ) / a['每手保证金'][t]
                else:
                    a.loc[t,'开仓权重'] = 0
            mslist = a.sort_values(by = '开仓权重',ascending = False).index
            n = mslist[0]
            if a['开仓权重'][n] > 0:
                a.loc[n,'满足60%补产品仓位前仓数'] = b['预计达标60%下一日需共开仓'][0]
                if a['满足60%补产品仓位前仓数'][n] >= a['获得所有增量减收需下日开仓'][n]:
                    a.loc[n,'补仓方案'] = a['获得所有增量减收需下日开仓'][n]
                else:
                    a.loc[n,'补仓方案'] =  a['满足60%补产品仓位前仓数'][n]
                a.loc[n,'此产品补完后剩余需补仓数'] = a['满足60%补产品仓位前仓数'][n] - a['补仓方案'][n]
                a.loc[n,'补仓所需金额(万元)'] = a['补仓方案'][n]*a['每手保证金'][n]/20000
                a.loc[n,'补仓所获增收额度'] = (a['补仓方案'][n] - a['达到基准需下一日开仓'][n])*self.leftday/self.totalday*a['大于基准每多一手可增加减收'][n]
            else:
                a[['满足60%补产品仓位前仓数','补仓方案','此产品补完后剩余需补仓数','补仓所需金额(万元)','补仓所获增收额度']] = 0
            f = 1
            while 0 not in a['此产品补完后剩余需补仓数'].values.tolist():
                for t in range(f,len(mslist)):               
                    j = mslist[t]
                    k = mslist[f-1]
                    if a['达到基准需下一日开仓'][j] < a['此产品补完后剩余需补仓数'][k] and a['开仓权重'][j] > 0:              
                        a.loc[j,'满足60%补产品仓位前仓数'] = a['此产品补完后剩余需补仓数'][k]
                        if a['满足60%补产品仓位前仓数'][j] >= a['获得所有增量减收需下日开仓'][j]:
                            a.loc[j,'补仓方案'] = a['获得所有增量减收需下日开仓'][j]
                        else:
                            a.loc[j,'补仓方案'] =  a['满足60%补产品仓位前仓数'][j]
                        a.loc[j,'此产品补完后剩余需补仓数'] = a['满足60%补产品仓位前仓数'][j] - a['补仓方案'][j]
                        a.loc[j,'补仓所需金额(万元)'] = a['补仓方案'][j]*a['每手保证金'][j]/20000
                        a.loc[j,'补仓所获增收额度'] = (a['补仓方案'][j] - a['达到基准需下一日开仓'][j])*self.leftday/self.totalday*a['大于基准每多一手可增加减收'][j]
                        f = t + 1
                        break                  
                    else:
                        if t == len(mslist)-1:                         
                            if a[a['产品'] == '低硫燃料油']['满足60%补产品仓位前仓数'][0] > 0 :
                                a.loc[0,'满足60%补产品仓位前仓数'] = a[a['产品'] == '燃料油']['满足60%补产品仓位前仓数'][0] + a['此产品补完后剩余需补仓数'][k]
                                a.loc[0,'补仓方案']  = a[a['产品'] == '燃料油']['补仓方案'][0] + a['此产品补完后剩余需补仓数'][k]
                                a.loc[0,'此产品补完后剩余需补仓数'] = 0
                                break
                            else:
                                a.loc[0,'满足60%补产品仓位前仓数'] = a['此产品补完后剩余需补仓数'][k]
                                a.loc[0,'补仓方案'] = a['此产品补完后剩余需补仓数'][k]
                                a.loc[0,'此产品补完后剩余需补仓数'] = 0
                                a.loc[0,'补仓所需金额(万元)'] = a['补仓方案'][0]*a['每手保证金'][0]/20000
                                a.loc[0,'补仓所获增收额度'] = 0
                                break
                        else:
                            continue          
            # =============================================================================
            if b['预计达标60%下一日需共开仓'][0] > 0:
                a['达标60%需保证金'] = b['预计达标60%下一日需共开仓'][0] * a['每手保证金'] / 20000
                for t in range(0,len(a)):
                    a.loc[t,'达标60%增加增量减收'] = min ((b['预计达标60%下一日需共开仓'][0] - a['达到基准需下一日开仓'][t]) * self.leftday / self.totalday  * a['大于基准每多一手可增加减收'][t] , a['增量减收可减最大值'][t] )
                # a['实施后可多返'] = a['达标60%增加增量减收'] + a['上交手续费'].sum() * (b['大于等于60%'][0] - b['当前达标减收比例'][0])
                a['实施后可多返'] = a['达标60%增加增量减收']
                a['达标60%年化收益率'] = 365*a['实施后可多返']/10000/(a['达标60%需保证金'] * 30)
            else:
                a[['达标60%需保证金','实施后可多返','达标60%年化收益率']] = 0
    
            self.base = a   
            inelist = ['达标完成率','超标完成率','持仓超标增投年化收益','达标60%年化收益率','当前已持仓预计年化收益率','累积公司已投入持仓超标增投年化收益','预计年化']
            

            self.reout(th1,th2,inelist)

    def anadce(self,basic):
        for n in ['A','B','C','D']:
            
            basic[n+'类达标日均']=basic['加权基准']+basic[n+'类减收']
            basic[n+'类加权相差日均']=basic[n+'类达标日均']-basic['加权预计日均']
            # basic[n+'类加权达标相差手数双边']=[(x*y/z/1.5) for (x,y,z) in zip (basic[n+'类加权相差日均'],basic['总交易天数'],basic['未交易天数'])]
            basic[n+'类加权达标相差手数双边']=[(x*y/self.nleftday/1.5) for (x,y) in zip (basic[n+'类加权相差日均'],basic['总交易天数'])]
            basic[n+'类达标相差手数双边']= [y if x < y else x for (x,y) in zip (basic[n+'类加权达标相差手数双边'],basic['非加权相差手数双边'])]
            basic[n+'类达标相差手数单边']=[math.ceil(x/2) for x in  basic[n+'类达标相差手数双边']]
            basic[n+'类资金预计占用']=[(x*y*z) for (x,y,z) in zip (basic['结算价'],basic['保证金占用'],basic[n+'类达标相差手数单边'])]
            basic = basic.set_index('产品')
            basic.loc['生猪',n+'类资金预计占用'] = basic.loc['生猪',n+'类资金预计占用']*2
            basic = basic.reset_index()
            basic[n+'类资金占用(万元)']=[(x+a)/10000 for (x,a) in zip (basic[n+'类资金预计占用'],basic['今日公司保证金占用'])]
            basic[n+'类完成百分比']=[(x/y) for (x,y) in zip (basic['加权预计日均'],basic[n+'类达标日均'])]
            basic[n+'类减收手续费']=[(x*y) for (x,y) in zip (basic['上交手续费'],basic[n+'类减收比例'])]
            basic[n+'类收益']=[(a/b*c*12) for (a,b,c) in zip(basic[n+'类减收手续费'],basic['已交易天数'],basic['总交易天数'])]
            basic[n+'类收益率']=[(x/y/10000) if y!=0 else 0 for (x,y) in zip(basic[n+'类收益'],basic[n+'类资金占用(万元)'])]
        self.atmp = basic        
    
    def reout(self,a,b,c):
        self.base = self.base.fillna('')
        
        def zh(d):
            self.base[d] = self.base[self.base[d]!=''][d].astype(float).apply(lambda x: format(x,'.2%'))
        for n in c:
            zh(n)
        self.base = self.base.round(2)
        n1 = pd.read_excel(self.tablehead,sheet_name = self.exchange)
        if self.monedflag == 1:
            n1 = pd.read_excel(self.tablehead,sheet_name = a)
            
        
        if self.chosflag == 1:
            n2 = pd.read_excel(self.tablehead,sheet_name = b)
            self.total = self.base.copy().reindex(columns = n2.columns.tolist())
            name = self.exchange+self.cnday+'各产品增量减收测算明细'+ self.moned+'.xlsx'
            self.total.to_excel(os.path.join(self.path,self.cnday,self.exchange,name),index = False,freeze_panes=(1,2))
        
        oricol = n1.columns.tolist()
        base = self.base.copy().reindex(columns = oricol).fillna(0)
        
        
        newcol = n1.loc[0]
        for n in range(0,len(newcol)):
            if newcol[n][:4] == 'self':
                exec(newcol[n])
                newcol.iloc[n] = self.tmp

        newdict = dict(zip(oricol,newcol))
        
        a = base.rename(columns = newdict)
        
        a.loc[len(a),a.columns[0]] = ''
        if self.exchoose == 'SHFE' or self.exchoose == '能源':
            a.loc[len(a)-1,newdict['增量实际减收']] = a[newdict['增量实际减收']].sum()
            a.loc[len(a)-1,newdict['增量减收可减最大值']] = a[newdict['增量减收可减最大值']].sum()
        # else:
            # a.loc[len(a)-1,newdict['持仓和法人占比日达标减收额度（元）']] = a[newdict['持仓和法人占比日达标减收额度（元）']].sum()
        a.loc[len(a),a.columns[0]] = '本月交易日'
        a.loc[len(a)-1,a.columns[1]] = str(self.totalday)+'天'
        a.loc[len(a),a.columns[0]] = '已交易'
        a.loc[len(a)-1,a.columns[1]] = str(self.traday)+'天'
        a.loc[len(a),a.columns[0]] = '剩余交易日'
        if self.monedflag == 1:
            self.leftday = 0
        a.loc[len(a)-1,a.columns[1]] = str(self.leftday)+'天'
        a=a.fillna('')
        if len(n1) > 1:          
            hlori = n1.loc[1][n1.loc[1] == 1].index.tolist()
            hlnew = list(itemgetter(*hlori)(newdict))
            self.outtb = a[a!=''].style.set_properties(subset=hlnew,**{'background-color':'#FF9900'}).highlight_null('none')
        else:
            self.outtb = a
        
   
        name = self.exchange+self.cnday+'各产品增量减收测算'+ self.moned+'.xlsx'
        # xx = self.outtb
        if self.exchoose == 'CZCE':

            self.outtb.to_excel(os.path.join(self.path,self.cnday,self.exchange,name),index = False,freeze_panes=(1,2))
        
        else:
            with pd.ExcelWriter(os.path.join(self.path,self.cnday,self.exchange,name)) as writer:
                self.toxlsx(self.outtb,writer,'增量及达标测算')
                self.toxlsx(self.dbwz,writer,'所有产品日均达标情况')
                self.toxlsx(self.qqpz,writer,'期权手数情况') 
   
        
    
    def czcefr(self):
        frzb = self.tzzcccxqh.groupby(['结算日','产品','客户性质'])['手数'].sum().unstack().fillna(0).reset_index()
        frzb['当前法人持仓占比总值'] = frzb['法人']/ (frzb['法人'] + frzb['自然人'])
        tdfrzb = frzb.copy()[frzb['结算日'] == int(self.td)][['产品','法人','当前法人持仓占比总值']].rename(columns={'当前法人持仓占比总值':'今日法人持仓占比','法人':'今日法人持仓(双边)'}) 
        frzb = frzb.groupby(['产品']).sum().reset_index()[['产品','自然人','法人','当前法人持仓占比总值']].rename(columns={'自然人':'当前自然人总持仓(双边)','法人':'当前法人总持仓(双边)'}) 
        frzb = pd.merge(frzb,tdfrzb,on='产品',how='left').fillna(0)
        frzb['当前法人持仓占比日均值'] = frzb['当前法人持仓占比总值']/self.traday
        frzb['预计法人持仓占比总值'] = frzb['当前法人持仓占比总值'] + frzb['今日法人持仓占比']*self.leftday
        frzb['预计法人持仓占比日均值'] = frzb['预计法人持仓占比总值']/self.totalday
        self.frzb = frzb  
        
    def czcedbjs(self,a):
        a['预计持仓增长手数'] = a['预计当月日均持仓'] - a['日均持仓基准']
        a['预计持仓增长比例'] = a['预计持仓增长手数'] / a['日均持仓基准']
        a['当前持仓达标减收比例'] = [e if a>=b or c>=d else 0 if  a<=0 else max(a/b,c/d)*e for (a,b,c,d,e) in zip(a['预计持仓增长手数'],a['持仓达标增长手数要求'],a['预计持仓增长比例'],a['持仓达标增长比例要求'],a['持仓达标减收比例'])]
        a['持仓增长达标情况'] = ['已达标' if a>=b or c>=d else '未达到基准' if  a<=0 else '达到基准未达标' for (a,b,c,d) in zip(a['预计持仓增长手数'],a['持仓达标增长手数要求'],a['预计持仓增长比例'],a['持仓达标增长比例要求'])]
        a['达标完成率'] = [max(a/b,c/d) for (a,b,c,d) in zip(a['预计持仓增长手数'],a['持仓达标增长手数要求'],a['预计持仓增长比例'],a['持仓达标增长比例要求'])]
        a['满足手数达标日均差额'] = a['持仓达标增长手数要求'] - a['预计持仓增长手数']
        a['满足比例达标日均差额'] = a['日均持仓基准'] * (1 + a['持仓达标增长比例要求']) - a['预计当月日均持仓']
        a['持仓达标所需最少日均差额'] = [x if x<=y else y for (x,y) in zip(a['满足比例达标日均差额'],a['满足手数达标日均差额'])]
        a['持仓达标需下日开仓'] = a['持仓达标所需最少日均差额'] * self.totalday / self.leftday   
        a['预计持仓达标需投资金（万元）'] = a['每手保证金']*a['持仓达标需下日开仓']/20000
        a['比例减收手续费'] = a['预计当月上交手续费'] * a['比例减收']
        a['预计当前持仓减收手续费（元）'] = a['预计当月上交手续费'] * a['当前持仓达标减收比例']
        a['持仓达标减收手续费（元）'] = a['预计当月上交手续费'] * a['持仓达标减收比例']
        
        
        a['上限达标年化收益率'] =  365*(a['持仓达标减收手续费（元）'] - a['预计当前持仓减收手续费（元）'])/10000 /(a['预计持仓达标需投资金（万元）']*30)
        a['持仓达标去除公司占有日均差额']  = a['预计考核月内公司占有日均(双边)'] + a['持仓达标所需最少日均差额']
        a['持仓达标去除公司占有需下日开仓']  = a['持仓达标去除公司占有日均差额'] * self.totalday / self.leftday  
        a['持仓达标去除公司占有需投资金（万元）'] = a['每手保证金']*a['持仓达标去除公司占有需下日开仓']/20000
        a['预计去除公司占有持仓增长手数'] = a['预计持仓增长手数'] - a['预计考核月内公司占有日均(双边)']
        a['预计去除公司占有持仓增长比例'] = a['预计去除公司占有持仓增长手数'] / a['日均持仓基准']
        a['去除公司占有后持仓达标减收比例'] = [e if a>=b or c>=d else 0 if  a<=0 else max(a/b,c/d)*e for (a,b,c,d,e) in zip(a['预计去除公司占有持仓增长手数'],a['持仓达标增长手数要求'],a['预计去除公司占有持仓增长比例'],a['持仓达标增长比例要求'],a['持仓达标减收比例'])]
        a['预计去除公司占有后持仓达标减收手续费（元）'] = a['预计当月上交手续费'] * a['去除公司占有后持仓达标减收比例']
        a['去除公司占有上限达标年化收益率'] =  365*(a['持仓达标减收手续费（元）'] - a['预计去除公司占有后持仓达标减收手续费（元）'])/10000 /(a['持仓达标去除公司占有需投资金（万元）']*30)
        
        self.czcefrjs(a)
        
        
        
        
    def czcefrjs(self,a):
        for t in range(0,len(a)):
            if a['法人达标值'][t] != '':
                a.loc[t,'当月法人持仓占比基准'] = a['法人最低基准'][t] + 0.01 * self.today.month 
                if a['当月法人持仓占比基准'][t] >=  a['法人达标值'][t]:
                    a.loc[t,'当月法人持仓占比基准'] = a['法人达标值'][t]
                if a['预计法人持仓占比日均值'][t] >= a['法人达标值'][t]:
                    a.loc[t,'法人达标情况'] = '已达标'
                    a.loc[t,'法人当前减收比例'] = a['法人达标乘积系数'][t]
                elif a['当月法人持仓占比基准'][t] != a['法人达标值'][t] and a['预计法人持仓占比日均值'][t] >= a['当月法人持仓占比基准'][t]:
                    a.loc[t,'法人达标情况'] = '已达标起减未达标满减'
                    a.loc[t,'法人当前减收比例'] = a['达到增幅要求后乘积系数'][t]
                else:
                    a.loc[t,'法人达标情况'] = '未达标'
                    a.loc[t,'法人当前减收比例'] = 0
                
                a.loc[t,'当前法人占比可增加减收（元）'] = a['预计当前持仓减收手续费（元）'][t]*(a['法人当前减收比例'][t]-1)
                a.loc[t,'法人占比达标可增加减收（万）'] =  a['持仓达标减收手续费（元）'][t]*(a['法人达标乘积系数'][t]-1)/10000
                a.loc[t,'持仓及法人占比均达标可减收(元)'] = a['持仓达标减收手续费（元）'][t]*a['法人达标乘积系数'][t]
                a.loc[t,'法人达标完成率'] = a['预计法人持仓占比日均值'][t]/a['法人达标值'][t]
                    
                a.loc[t,'法人达标日均差值'] = a['法人达标值'][t] - a['预计法人持仓占比日均值'][t]
                a.loc[t,'法人达标需日后占比'] = a['今日法人持仓占比'][t] + a['法人达标日均差值'][t] * self.totalday / self.nleftday
                a.loc[t,'单日法人达标需增仓'] = (a['今日持仓'][t]*a['法人达标需日后占比'][t] - a['今日法人持仓(双边)'][t])/(1-a['法人达标需日后占比'][t])
                if a['法人达标需日后占比'][t] >= 1 :
                    a.loc[t,'单日法人达标需增仓'] = '无法达标'
                    a.loc[t,'法人达比需增投（万元）'] = '无法达标'
                    a.loc[t,'法人达比增投年化收益'] =  0
                    a.loc[t,'法人达标情况'] = '无法达标'
                elif a['法人达标需日后占比'][t] > 0:
                    a.loc[t,'法人达比需增投（万元）'] = a['每手保证金'][t]*a['单日法人达标需增仓'][t]/20000
                    a.loc[t,'法人达比增投年化收益'] =  365*a['法人占比达标可增加减收（万）'][t]/(a['法人达比需增投（万元）'][t]*30)
                else:
                    a.loc[t,'法人达比需增投（万元）'] = a['每手保证金'][t]*a['单日法人达标需增仓'][t]/20000
                    a.loc[t,'法人达比增投年化收益'] =  365*a['法人占比达标可增加减收（万）'][t]/(a['法人达比需增投（万元）'][t]*30)
        
        a = a.fillna('')
        self.czcezdhy(a)
        
    def czcezdhy(self,a):
        a = pd.merge(a.copy(),self.zdhy,on='产品',how='left')
        a['预计重点合约增长手数'] = a['预计重点合约平均日均持仓(双边)'] - a['重点合约日均基准']
        a['预计重点合约增长比例'] = a['预计重点合约增长手数'] / a['重点合约日均基准']
        a['当前重点合约减收比例'] = [e if a>=b or c>=d else 0 if  a<=0 else max(a/b,c/d)*e for (a,b,c,d,e) in zip(a['预计重点合约增长手数'],a['重点合约日均基准'],a['预计重点合约增长比例'],a['重点合约较基数增长比例要求'],a['重点合约达标减免比例'])]

      
        a['重点合约增长达标情况'] = ['已达标' if a>=b or c>=d else '未达到基准' if  a<=0 else '' if pd.isnull(a)  else '达到基准未达标' for (a,b,c,d) in zip(a['预计重点合约增长手数'],a['重点合约日均基准'],a['预计重点合约增长比例'],a['重点合约较基数增长比例要求'])]
        a['重点合约达标完成率'] = [max(a/b,c/d) for (a,b,c,d) in zip(a['预计重点合约增长手数'],a['重点合约日均基准'],a['预计重点合约增长比例'],a['重点合约较基数增长比例要求'])]
        a['重点合约满足手数达标日均差额'] = a['重点合约较基数增长手数要求'] - a['预计重点合约增长手数']
        a['重点合约满足比例达标日均差额'] = a['重点合约日均基准'] * (1 + a['重点合约较基数增长比例要求']) - a['预计重点合约平均日均持仓(双边)']
        a['重点合约持仓达标所需最少日均差额'] = [x if x<=y else y for (x,y) in zip(a['重点合约满足手数达标日均差额'],a['重点合约满足比例达标日均差额'])]
        a['重点合约持仓达标需下日开仓'] = a['重点合约持仓达标所需最少日均差额'] * self.totalday *  a['重点合约数量'] / self.leftday   
        a['预计重点合约持仓达标需投资金（万元）'] = a['每手保证金']*a['重点合约持仓达标需下日开仓']/20000
        # a['比例减收手续费'] = a['预计当月上交手续费'] * a['比例减收']
        a['预计重点合约当前持仓减收手续费（元）'] = a['预计当月上交手续费'] * a['当前重点合约减收比例']
        a['持仓重点合约达标减收手续费（元）'] = a['预计当月上交手续费'] * a['重点合约达标减免比例']
        
        
        a['重点合约上限达标年化收益率'] =  365*(a['持仓重点合约达标减收手续费（元）'] - a['预计重点合约当前持仓减收手续费（元）'])/10000 /(a['预计重点合约持仓达标需投资金（万元）']*30)
        # a['持仓达标去除公司占有日均差额']  = a['预计考核月内公司占有日均(双边)'] + a['持仓达标所需最少日均差额']
        # a['持仓达标去除公司占有需下日开仓']  = a['持仓达标去除公司占有日均差额'] * self.totalday / self.leftday  
        # a['持仓达标去除公司占有需投资金（万元）'] = a['每手保证金']*a['持仓达标去除公司占有需下日开仓']/20000
        # a['预计去除公司占有持仓增长手数'] = a['预计持仓增长手数'] - a['预计考核月内公司占有日均(双边)']
        # a['预计去除公司占有持仓增长比例'] = a['预计去除公司占有持仓增长手数'] / a['日均持仓基准']
        # a['去除公司占有后持仓达标减收比例'] = [e if a>=b or c>=d else 0 if  a<=0 else max(a/b,c/d)*e for (a,b,c,d,e) in zip(a['预计去除公司占有持仓增长手数'],a['持仓达标增长手数要求'],a['预计去除公司占有持仓增长比例'],a['持仓达标增长比例要求'],a['持仓达标减收比例'])]
        # a['预计去除公司占有后持仓达标减收手续费（元）'] = a['预计当月上交手续费'] * a['去除公司占有后持仓达标减收比例']
        # a['去除公司占有上限达标年化收益率'] =  365*(a['持仓达标减收手续费（元）'] - a['预计去除公司占有后持仓达标减收手续费（元）'])/10000 /(a['持仓达标去除公司占有需投资金（万元）']*30)
        
        
        self.atmp = a
    
    def shcom(self,a,b,c):
        self.exchoose = a
        zlwj = pd.read_excel(self.exchangwj,sheet_name= b)
        self.dbwj = pd.read_excel(self.exchangwj,sheet_name= c)  
        self.cominit() 
        self.company()
        if self.imgflag == 1:
            self.drawimg()
                
        base = pd.merge(zlwj,self.qhpz,on='产品',how='left').fillna(0)
        base = pd.merge(base,self.zlhy,on='产品',how='left')
        self.comother(base)
        base = self.atmp 
        base = pd.merge(base,self.comqz,on='产品',how='left').fillna(0)
        base['保证金'] = base['保证金']/10000
        base = base.rename(columns={'保证金':'当前已投入(万元)'})
        base['达到基准日均还差'] = base['考核基数'] - base['预计当月日均持仓']
        base['达到基准总差额'] = base['达到基准日均还差']*self.totalday
        base['比例减收手续费30%'] = base['预计当月上交手续费'] * base['交易所比例减收']
        base['增量减收可减最大值'] = base['预计当月上交手续费'] *  base['增量减收不能超过上交额度']
        base['增量减收最大时超过基准日均'] = base['增量减收可减最大值']/base['大于基准每多一手可增加减收']
        base['超过基准可增持的最大手数']  = base['增量减收最大时超过基准日均'] * self.totalday
        base['获取所有增量减收所需日均'] =  base['增量减收最大时超过基准日均'] +  base['考核基数']
        a = base.copy()
        for t in range(0,len(a)):
            if a['达到基准总差额'][t] < 0:
                a.loc[t,'达到基准需下一日开仓'] = 0
                a.loc[t,'预计日均超过基准'] = -a['达到基准日均还差'][t]
                a.loc[t,'预计增量减收'] = a['预计日均超过基准'][t] * a['大于基准每多一手可增加减收'][t]
                if a['预计增量减收'][t] >= a['增量减收可减最大值'][t]:
                    a.loc[t,'达标情况'] = '已超标'
                    a.loc[t,'超过增量减收最大值'] = a['预计增量减收'][t] - a['增量减收可减最大值'][t]
                else:
                    a.loc[t,'达标情况'] = '已达标未超标'
                    a.loc[t,'超过增量减收最大值'] = 0
                a.loc[t,'增量实际减收'] = min(a['预计增量减收'][t],a['增量减收可减最大值'][t])
                a.loc[t,'未获得的增量减收'] = a['增量减收可减最大值'][t] - a['增量实际减收'][t]
                a.loc[t,'获得所有增量减收日均仓数差额'] = a['未获得的增量减收'][t] / a['大于基准每多一手可增加减收'][t]
                a.loc[t,'获得所有增量减收需下日开仓'] = a['获得所有增量减收日均仓数差额'][t] * self.totalday / self.nleftday
            
            else:
                a.loc[t,'达标情况'] = '未达标'
                a.loc[t,'达到基准需下一日开仓'] = a['达到基准总差额'][t] / self.nleftday
                a.loc[t,'超过增量减收最大值'] = 0
                a.loc[t,'预计日均超过基准'] = 0
                a.loc[t,'预计增量减收'] = 0
                a.loc[t,'增量实际减收'] = 0
                a.loc[t,'未获得的增量减收'] = a['增量减收可减最大值'][t]
                a.loc[t,'获得所有增量减收日均仓数差额'] = a['未获得的增量减收'][t] / a['大于基准每多一手可增加减收'][t] + a['达到基准日均还差'][t]                
                a.loc[t,'获得所有增量减收需下日开仓'] = a['获得所有增量减收日均仓数差额'][t] * self.totalday/ self.nleftday
        a['获取所有增量减收去除公司占有差额'] = (a['获取所有增量减收所需日均'] - a['预计当月日均持仓'] )*self.totalday +a['预计考核月内公司共占有(双边)']
        a['达标完成率'] = a['预计当月日均持仓'] /  a['考核基数'] 
        a['超标完成率'] = a['预计当月日均持仓'] /  a['获取所有增量减收所需日均']
        a['预计持仓达标需投资金（万元）'] = a['达到基准需下一日开仓'] * a['每手保证金'] / 20000
        a['预计持仓超标需投资金（万元）'] = a['获得所有增量减收需下日开仓'] * a['每手保证金'] / 20000
        a['预计超标下一交易日可释放仓位'] = a['超过增量减收最大值'] / a['大于基准每多一手可增加减收'] * self.totalday/ self.nleftday
        a['预计超标可释放资金（万元）'] =  a['预计超标下一交易日可释放仓位'] *  a['每手保证金'] / 20000
        for t in range(0,len(a)):
            if a['预计持仓超标需投资金（万元）'][t] == 0:
                a.loc[t,'持仓超标增投年化收益'] = 0     
                a.loc[t,'累积公司已投入持仓超标增投年化收益'] = 0
              
            else:
                a.loc[t,'持仓超标增投年化收益'] =  365*a['未获得的增量减收'][t]/10000/(a['预计持仓超标需投资金（万元）'][t] * 30)
                a.loc[t,'累积公司已投入持仓超标增投年化收益'] = 365*a['增量减收可减最大值'][t]/10000/( (a['预计持仓超标需投资金（万元）'][t] + a['当前已投入(万元)'][t]) * 30)
            if  a['当前已投入(万元)'][t] == 0:
                a.loc[t,'当前已持仓预计年化收益率'] = 0
            else:
                a.loc[t,'当前已持仓预计年化收益率'] = 365*a['增量实际减收'][t]/10000/(a['当前已投入(万元)'][t]*30)
        a['去除公司占有基准差额日均'] =  a['达到基准日均还差']+ a['预计考核月内公司占有日均(双边)']
        a['去除公司占有获取所有增量减收差额日均'] = a['获取所有增量减收所需日均'] - a['预计当月日均持仓']+ a['预计考核月内公司占有日均(双边)']
        a['达标满减需投入(万元)'] = a['去除公司占有获取所有增量减收差额日均'] * a['每手保证金'] / 20000
        a['预计年化'] = 365*a['增量减收可减最大值']/10000/(a['达标满减需投入(万元)'] * 30)
        for t in range(0,len(a)):
            if a['去除公司占有获取所有增量减收差额日均'][t] <= 0 :
                a.loc[t,'去除公司占有后达标情况'] = '已超标'
            elif a['去除公司占有获取所有增量减收差额日均'][t] > 0  and a['去除公司占有基准差额日均'][t] <=0:
                a.loc[t,'去除公司占有后达标情况'] = '已达标未超标'
            else:
                a.loc[t,'去除公司占有后达标情况'] = '未达标'
        
        
        
        self.base = a.fillna(0)
        
# =============================================================================

            
            
    def drawimg(self):
        
        a = self.jysjsmx
        for p in a['产品'].unique():
            buy = a[a['产品'] == p][['结算日','买持仓','卖持仓','手数']].rename(columns={'买持仓':'买','卖持仓':'卖','手数':'总持仓'})
            b = self.comimg[self.comimg['产品'] == p].groupby('结算日').sum('手数')['手数'].reset_index()
            buy = pd.merge(buy,b,how='left',on='结算日').rename(columns={'手数':'公司持仓'})
            buy = buy.fillna(0)
            buy['结算日'] = buy['结算日'].apply(lambda x:datetime.datetime.strptime(str(x),'%Y%m%d'))
            buy = buy.set_index('结算日')
            img = buy.plot(style=['red','green','yellow','purple'],title=p,figsize=(24,16),fontsize=20)
            img.axes.title.set_size(40)
            img.legend(fontsize=30) 
            img.set_xlabel('结算日',fontsize=24)
            img.set_ylabel('手数',fontsize=24)
            try:
                img.get_figure().savefig(os.path.join(self.path,self.cnday,self.exchange,self.cnday+p+"持仓情况.png"))
                img.figure.close()
            except:
                pass
            
    def company(self):
        a = self.tzzcccxqh[self.tzzcccxqh['投资者代码'].isin(self.comuserlist)][['结算日','产品','买持量','卖持量','手数','合约','保证金']]
        a['卖持量'] = -a['卖持量']
        a['flag'] = a['买持量']+a['卖持量']
        # a = 
        # a['tflag'] = a['买卖'].apply(lambda x:1 if x =='买' else -1)
        # b = a.groupby('flag').size()
        # b = b[b>1].reset_index()
        # pp = pd.merge(a,b,on='flag')[['flag','tflag']]
        # ppqk = pp.groupby('flag').sum('tflag')
        # if len(ppqk) != 0:
        #     ppqk = ppqk[ppqk['tflag'] == 0]
        self.comimg = a[a['flag'] == 0][['结算日','产品','手数','合约','保证金']]
        
        tdtr = self.comimg[self.comimg['结算日'] == int(self.td)][['产品','保证金']]
        td = self.comimg[self.comimg['结算日'] == int(self.td)][['产品','合约','手数']]
        comqz = self.comimg.groupby(['产品','合约']).sum('手数')['手数'].reset_index()
        comqz = pd.merge(comqz,td.set_index('产品'),on='合约',how='left').rename(columns={'手数_x':'公司占有总持仓(双边)','手数_y':'公司今日占有(双边)'}).fillna(0)
        comqz['预计考核月内公司共占有(双边)'] = comqz['公司占有总持仓(双边)'] + comqz['公司今日占有(双边)'] * self.leftday
        # comqz = pd.merge(comqz,self.lsjsjcx.set_index('产品'),on='合约',how='left')
        comqz['预计考核月内公司占有日均(双边)'] = comqz['预计考核月内公司共占有(双边)']/ self.totalday
        comqz['当前日均(双边)'] = comqz['公司占有总持仓(双边)']/ self.traday
       
     
        self.comqz = comqz[['产品','公司占有总持仓(双边)','公司今日占有(双边)','预计考核月内公司共占有(双边)','预计考核月内公司占有日均(双边)','当前日均(双边)']].groupby('产品').sum().reset_index()
        self.comqz = pd.merge(self.comqz,tdtr,on='产品',how ='left' )
   


    def comother(self,a):
        sjssf = self.jysjsmxqh.groupby('产品').sum()['上交手续费'].reset_index()
        a = pd.merge(a.copy(),sjssf,on = '产品',how='left')
        a = a.copy().fillna(0)
        a['每手保证金'] = a['保证金率'] * a['结算价'] * a['每手数量']
        a['预计当月上交手续费'] = a['上交手续费'] / self.traday * self.totalday
        
        self.atmp = a.fillna(0)
        
    def cominit(self):
        #############################建立目录###################
        
        
        ################################初始化日期#####################
        if self.exchoose == 'CZCE' or self.exchoose == 'DCE':
            monthbeg = self.today.replace(day=1)                                                         #获得当月月初日期
            monthend = self.today.replace(day=1) + relativedelta(months=1) - relativedelta(days=1)        #获得当月月末日期
            nextmonthbeg = self.today.replace(day=1) + relativedelta(months=1)
            nextmonthend = nextmonthbeg + relativedelta(months=1) - relativedelta(days=1) 
        else:
            if int(self.td) <= 20211025:
                monthbeg = datetime.datetime.strptime('20211008', '%Y%m%d') 
                monthend = datetime.datetime.strptime('20211025', '%Y%m%d') 
        #10月26号及以后 
            elif self.today >= self.today.replace(day=26):
                monthbeg = self.today.replace(day=26)
                monthend = self.today.replace(day=25) + relativedelta(months=1)

               
        
            else:
                monthbeg = self.today.replace(day=26) - relativedelta(months=1)
                monthend = self.today.replace(day=25)

            nextmonthbeg = monthbeg + relativedelta(months=1)
            nextmonthend = monthend + relativedelta(months=1)
        
        mwd = pd.date_range(monthbeg,monthend,freq='B')                                              #生成当月所有工作日的数组   
        nextmwd = pd.date_range(nextmonthbeg,nextmonthend,freq='B')
        for n in self.jjr:
            m = datetime.datetime.strptime(n,'%Y%m%d')
            if m in mwd:
                mwd = mwd.drop(m)                                                                          #从当月所有工作日的数组中去掉当月节日日期
            if m in nextmwd:
                nextmwd = nextmwd.drop(m)
        self.tracalen = pd.DataFrame(mwd[mwd<=self.today].strftime('%Y%m%d')).rename(columns={0:'结算日'}).astype(int)
        self.totalday = len(mwd)                                                                     #数组总数即为当月的交易日数量
        self.leftday = len(mwd[mwd>self.today])                                                      #数组中日期大于今天日期的数量，即为剩余交易日的数量
        self.traday = len(mwd[mwd<=self.today]) 
        
        self.nextotalday = len(nextmwd)
        
        
        
        self.moned = ''
        self.monedflag = 0
        self.nleftday = self.leftday
        if self.leftday == 0:
            self.monedflag = 1
            self.nleftday = 1
            # self.leftday = 1
            self.moned = '(部分计算按剩1天计算)'
        if self.leftday > 5:
            self.tmpday = self.leftday - 5
        else:
            self.tmpday = 1
        #生成下一交易日前判断当天交易日是否为最后一个交易日
        if self.today != mwd[-1]:
            self.ntday = mwd[mwd > self.today][0]                                   #下一交易日
        else:
            self.ntday = self.today
            
        #  下一交易日 x月n日 
        self.ntcnday = self.ntday.__format__('%m月%d日')
                    
        #  当前交易日 x月n日    
        self.cnday = self.today.__format__('%m月%d日')
        
        #x月n日-y月m日
        self.bwday = monthbeg.__format__('%m月%d日-') + self.today.__format__('%m月%d日')
        # 建立文件夹
        try:
            os.makedirs(os.path.join(self.path,self.cnday,self.exchange))
        except:
            pass
        ##############################################基本计算########################################
        def fu(b):
            b['预计当月持仓总手数'] = b['手数'] + b['今日持仓'] * self.leftday
            b['预计当月日均持仓'] = b['预计当月持仓总手数'] / self.totalday   
        self.bzjl = self.bzjl[self.bzjl['交易所'] == self.exchange]
        self.hylist = self.bzjl['合约'].tolist()    
        
        self.jysjsmx = self.jysjsmx[self.jysjsmx['交易所'] == self.exchange]
        self.jysjsmx = self.jysjsmx[self.jysjsmx['结算日'] >= int(monthbeg.__format__('%Y%m%d'))]
        self.pzlist = list(set(self.cpsz['产品'].tolist())&set(self.jysjsmx['产品'].unique().tolist()))
        
        self.tzzcccxtmp = self.tzzcccx[self.tzzcccx['交易所'] == self.exchange]
        self.tzzcccxtmp = self.tzzcccxtmp[self.tzzcccxtmp['结算日'] >= int(monthbeg.__format__('%Y%m%d'))]
        
        self.tzzcccxqh = self.tzzcccxtmp[self.tzzcccxtmp['产品'].isin(self.pzlist)]
        self.tzzcccxqq = self.tzzcccxtmp.drop(self.tzzcccxqh.index)
        self.tdtzzcccxqh = self.tzzcccxqh[self.tzzcccxqh['结算日'] == int(self.td) ]
        self.tdtzzcccxqq = self.tzzcccxqq[self.tzzcccxqq['结算日'] == int(self.td) ]
                     
        # self.lsjsjcxtmp = self.lsjsjcx[self.lsjsjcx['交易所'] == a]        
        # self.bzjltmp = self.lsjsjcx[self.lsjsjcx['交易所'] == a]      
        # self.jysjsmxtmp = self.jysjsmx[self.jysjsmx['交易所'] == a]
        
        
        
        self.jysjsmxqh = self.jysjsmx[self.jysjsmx['产品'].isin(self.pzlist)]
        self.jysjsmxqq = self.jysjsmx.drop(self.jysjsmxqh.index)
        
        
        
        self.qhpz = self.jysjsmxqh.groupby('产品').sum('手数')['手数'].reset_index().set_index('产品')
        self.tdqhpz = self.jysjsmxqh[self.jysjsmxqh['结算日']==int(self.td)][['产品','手数']].set_index('产品')
        self.qhpz['今日持仓'] = self.tdqhpz['手数']
        self.qhpz = self.qhpz.fillna(0)
        fu(self.qhpz)
        self.qhpz = self.qhpz.reset_index()
        
        self.qqpz = self.jysjsmxqq.groupby('产品').sum('手数')['手数'].reset_index().set_index('产品')
        self.tdqqpz = self.jysjsmxqq[self.jysjsmxqq['结算日']==int(self.td)][['产品','手数']].set_index('产品')
        self.qqpz['今日持仓'] = self.tdqqpz['手数']
        self.qqpz = self.qqpz.fillna(0)
        fu(self.qqpz)
        self.qqpz = self.qqpz.reset_index()
        
        self.total = self.jysjsmx['手数'].sum()
        self.tdtotal = self.jysjsmx[self.jysjsmx['结算日'] == int(self.td)]['手数'].sum()
        self.totalnum = self.total + self.tdtotal * self.leftday
       
        
        ###########################################主力合约#########################################
        hy = self.tzzcccxqh.groupby(['产品','合约']).sum('手数').reset_index()[['产品','合约','手数']]
        
        self.cczlhy = hy.loc[hy.groupby('产品')['手数'].idxmax()][['产品','合约']].set_index('产品')
        
        jsjminhy = self.bzjl.loc[self.bzjl.groupby('产品')['结算价'].idxmin()][['合约','产品']].set_index('产品')     
        jsjminhy.update(self.cczlhy)
        self.zlhy = pd.merge(jsjminhy,self.bzjl,on='合约',how='left')[['合约','结算价','产品','每手数量','保证金率','产品代码']]
        if self.exchoose == 'CZCE':
            # jsjminhy['产品代码'] = jsjminhy['合约'].str[:2]
            self.zdhyfr = zdhyfr = pd.read_excel(self.otherinfo,sheet_name = self.czcezdhyt)                                   #重点合约生成fr
            zdhyss = pd.merge(zdhyfr,hy,on='合约',how='left')[['合约','手数']]
            zdhyss = zdhyss.fillna(0)
            zdhyjsj = pd.merge(zdhyss,self.bzjl,on='合约',how='left')
            zdhyjsj.loc[zdhyjsj[zdhyjsj['手数'] == 0].index,'手数'] = -zdhyjsj['结算价']
            zdhy = zdhyjsj.loc[zdhyjsj.groupby(['产品'])['手数'].idxmax()].set_index('产品')[['合约','结算价','每手数量','保证金率','产品代码']]
            self.zlhy = self.zlhy.set_index('产品')
            self.zlhy.update(zdhy)
            self.zlhy = self.zlhy.reset_index()
            # jsjminhy.update(zdhy)
            # jsjminhy['产品代码'] = jsjminhy['合约'].str[:2]
            # zlhy = pd.merge(jsjminhy,self.lsjsjcx,on='合约',how='left')[['合约','结算价','产品','产品代码']]
            zdhyfr['产品代码'] = zdhyfr['合约'].str[:2]                                     #添加产品代码列
            zdhyfrdata = zdhyfr.groupby('产品代码').apply(lambda x: ','.join(x['合约'])).reset_index().rename(columns={0:'重点合约'})      #按产品分类合并重点合约为同一行
            zdhyfrdata = pd.merge(zdhyfrdata,zdhyfr.groupby('产品代码').count().reset_index().rename(columns= {'合约':'重点合约数量'}),how='left',on='产品代码')
            self.zlhy = pd.merge(self.zlhy,zdhyfrdata,on='产品代码',how='left').fillna('')[['产品','合约','结算价','重点合约','重点合约数量','每手数量','保证金率']]
            self.zdhynum = self.zlhy[self.zlhy['重点合约']!=''][['产品','重点合约数量']]
            self.zlhy = self.zlhy[['产品','合约','结算价','重点合约','每手数量','保证金率']]
            
    def cinit(self):
        self.path = path = os.getcwd()
        cpath = os.path.join(path,'config')
        
        #config文件设置 
        comuser = '公司账户'
        jjrfile = '节假日'
        self.tablehead = os.path.join(cpath,'表头.xlsx')
        self.exchangwj = os.path.join(cpath,'交易所文件.xlsx')
        self.otherinfo = os.path.join(cpath,'其它信息.xlsx')
        #上期文件
        self.shfezlwjt = '上期增量文件'
        self.shfedbwjt = '上期达标文件'
        
        #能源文件
        self.inezlwjt = '能源增量文件'
        self.inedbwjt = '能源达标文件'
        
        #郑商文件
        self.czcewjt = '郑商文件'                                                
        self.czcezdhyt = '郑商所重点合约'    
        
        #获取当前交易日，
        if self.td != '':
            self.today = datetime.datetime.strptime(self.td, '%Y%m%d') 
        else:            
            self.today = datetime.datetime.today()   #获得当天日期
            self.td = self.today.__format__('%Y%m%d')
            
        self.framework(self.tzzcccxt)
        self.tzzcccx = self.atmp[self.atmp['结算日'] != '总计'][['结算日','投资者代码','交易所','产品名称','合约','买持量','买均价','卖持量','卖均价','保证金']]
        self.tzzcccx['结算日'] = self.tzzcccx['结算日'].astype(int)
        self.tzzcccx['投资者代码'] = self.tzzcccx['投资者代码'].astype(str).str[:-2]
        self.tzzcccx = self.tzzcccx[self.tzzcccx['结算日'] <= int(self.td)].rename(columns={'产品名称':'产品'})
        self.tzzcccx['手数'] = self.tzzcccx['买持量'] + self.tzzcccx['卖持量']
        
        if self.frflag == 1 :
            self.framework(self.tzzzlcxt)                                                                        
            self.tzzzlcx = self.atmp['投资者代码'].tolist()                                                        
            self.tzzcccx['客户性质'] = self.tzzcccx['投资者代码'].apply(lambda x:'法人' if x in self.tzzzlcx else '自然人' )
        
        self.framework(self.bzjlt)
        self.bzjl = self.atmp[['合约','投机多头(按金额)']].rename(columns={'投机多头(按金额)':'保证金率'})
        
        
        self.framework(self.cpszt)
        self.cpsz = self.atmp[self.atmp['产品类型'] == '期货']
        self.cpsz = self.cpsz.copy()[self.cpsz['产品生命周期状态'] == '活跃'][['产品名称','产品代码','合约乘数']].rename(columns={'产品名称':'产品','合约乘数':'每手数量'})
        
        
        self.framework(self.jysjsmxt)
        self.jysjsmx = self.atmp[self.atmp['结算日'] != '总计'][['结算日','交易所','产品','上交手续费','买持仓','卖持仓']]
        self.jysjsmx['结算日'] = self.jysjsmx['结算日'].astype(int)
        self.jysjsmx['手数'] = self.jysjsmx['买持仓'] + self.jysjsmx['卖持仓']
        
        

        
        
        self.framework(self.lsjsjcxt)
        self.lsjsjcx = self.atmp[['产品/合约','结算价','交易所','产品']].rename(columns={'产品/合约':'合约'})
        
        self.bzjl = pd.merge(self.bzjl,self.lsjsjcx,how='left',on='合约')
        self.bzjl = pd.merge(self.bzjl,self.cpsz,how='left',on='产品')
        
        self.comuserlist = pd.read_excel(self.otherinfo,sheet_name = comuser)['公司账户'].map(str).tolist()
        
        self.jjr = pd.read_excel(self.otherinfo,sheet_name = jjrfile)['节假日'].map(str).tolist()
        
        

    
        
    def __init__(self,a):                                                                              #主程序
        self.wzsz(a)    
       
        
    def framework(self,a):                                                                          #--------DataFrame格式转换模块--------#
        self.atmp = pd.read_csv(a,encoding='gbk',low_memory=False,thousands=',')      #以gbk编码读取csv文件，转为DataFrame格式存在atmp临时变量中    
    

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Wintool()
    sys.exit(app.exec_()) 



    
    
    
    
    
    
    
    
    
    
    
    
    
    