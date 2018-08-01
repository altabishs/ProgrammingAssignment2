# -*- coding: utf-8 -*-
import re;import sys;from datetime import datetime,date
from datetime import timedelta,datetime
import datetime
import string
import csv
import os;from textblob import TextBlob;#from langdetect import detect;from textblob import TextBlob;
from PyQt4.QtCore import*;from PyQt4.QtGui import*
import xlrd ;import xlsxwriter ;import ctypes
from BeautifulSoup import BeautifulSoup;import json
import traceback
import requests
import pandas

#=============================PyQt4 interface ===========================================

# -*- coding: utf-8 -*-
from PyQt4.QtCore import*;from PyQt4.QtGui import*
import xlsxwriter;import sys
reload(sys);sys.setdefaultencoding('utf8')
class QcWindow(QDialog):
    def __init__(self,EntityLst,msgList,toolTip,exePath,maxRows,duration,parent=None):
        super(QcWindow, self).__init__(parent)
        self.setGeometry(120, 100, 1200, 600)
        self.setWindowTitle("Evalueserve-LEI Data Quality Assessment Tool")
        self.setWindowIcon(QIcon(exePath+'\evs.png'))
        self.EntityLst=EntityLst;self.msgList=msgList;self.toolTip=toolTip ;self.exePath=exePath;self.maxRows=maxRows
        tablemodel = MyTableModel(EntityLst, self)
        tableview = QTableView()
        tableview.setModel(tablemodel)
        tableview.setColumnWidth(0,200);tableview.setColumnWidth(1,100);tableview.setColumnWidth(2,600)
        tableview.horizontalHeader().setStretchLastSection(True)
        tableview.setSelectionBehavior(QTableView.SelectRows)
        tableview.clicked.connect(self.viewClicked)
        tableview.horizontalHeader().setDefaultAlignment(Qt.AlignLeft)
        # ---- bottom buttons
        OkBtn = QPushButton('OK')
        CancelBtn = QPushButton('Cancel')
        n=0;
        for err in msgList: n=n+len(err)
        lbl=QLabel('Total records checked = '+ str(self.maxRows) +' ; Records with errors = '+ str(len(EntityLst)) +' ; No of errors found = ' + str(n) + '     Total Duration: {}'.format(duration))
        font = QFont();font.setPointSize(10);font.setBold(True);lbl.setFont(font)
        lbl.setStyleSheet('QLabel {color:#EE2653;}')  #668DB9
        ExportBtn = QPushButton('Export To Excel')
        ExportBtn.setFixedWidth(100)
        ExportBtn.setFont(QFont())
        OkBtn.setDefault(True)
        buttonLayout = QHBoxLayout()
        buttonLayout.setContentsMargins(0, 0, 0, 0)
        buttonLayout.setSpacing(12)
        buttonLayout.addWidget(OkBtn)
        buttonLayout.addWidget(CancelBtn)
        buttonLayout.addStretch()
        buttonLayout.addWidget(lbl)
        buttonLayout.addStretch()
        ExportBtn.colorCount()
        buttonLayout.addWidget(ExportBtn)
        
        layout = QVBoxLayout(self)        
        layout.addWidget(tableview)
        layout.addLayout(buttonLayout)
        self.setLayout(layout)
        OkBtn.clicked.connect(self.accept)
        CancelBtn.clicked.connect(QCoreApplication.quit)
        ExportBtn.clicked.connect(self.ExportToExcel)
    #----------------To export all the data along with error description into excel file    
    def ExportToExcel(self):
        try:
            WB=xlsxwriter.Workbook(self.exePath+'\Errors.xlsx')
            wSht=WB.add_worksheet("Errors") #.add_sheet('Errors') 
            hdrFormat = WB.add_format({'bold': True,'font_name':'Arial','font_size':'10','font_color':'white','bg_color': '#315EA5'})
            dataformat=WB.add_format({'font_name':'Arial','font_size':'9'})
            dataformat.set_text_wrap()
            dataformat.set_align('left')
            cols=['S.No.','Key Field ID', 'O', 'OfficialEntityName','Error Count','Field Name', 'Field Value', 'Error Message'] 
            for index,c in enumerate(cols):wSht.write(0,index,c,hdrFormat)
            r=1; counter=1
            for k, ent in enumerate(self.EntityLst):
                c=1
                for e in ent:
                    wSht.write(r,c,e,dataformat)
                    #wSht.write(r,0,counter)
                    c+=1
                for m in self.msgList[k]:
                    if c<5:
                        for e in ent:wSht.write(r,c,e,dataformat);c+=1
                    for msg in m:wSht.write(r,c,msg,dataformat);c+=1
                    wSht.write(r,0,counter,dataformat)
                    r+=1 ;c=1;counter+=1      
            wSht.set_column(0,0,5);wSht.set_column(1,2,10) ;wSht.set_column(3,3,30);wSht.set_column(5, 6, 25);wSht.set_column(7,7, 40)  
            WB.close()
            QMessageBox.about(self, "Evalueserve-LEI Data Quality Assessment", "%s" % ('Data Extraction Complete..!!'+ '\n' + 'Error log file Errors.xlsx is saved at "' + self.exePath+ '".'))    
        except IOError:QMessageBox.about(self, "Evalueserve-LEI Data Quality Assessment", "%s" % ('Could not open file! Please close "Errors.xlsx" file first.'+'\n'+'Now try again.....!!'))   
    #-------------------------Child dialogue window on the main window row click event------------------       
    def viewClicked(self,clickedIndex):
            row=clickedIndex.row(); msgList1=self.msgList[row];toolTipText=self.toolTip[row];pPath=self.exePath;entList=self.EntityLst
            class ErrorChWindow(QDialog):
                def __init__(self, parent=None):
                    super(ErrorChWindow, self).__init__(parent)
                    self.setWindowTitle("Evalueserve-LEI Data Quality Assessment Errors")
                    tablemodel = ErrorTableModel(msgList1, toolTipText,self)
                    lbl=QLabel('For Official Entity : '+ entList[row][2] )
                    font = QFont();font.setPointSize(10);font.setBold(True);lbl.setFont(font)
                    lbl.setStyleSheet('QLabel {color:#315EA5;}')  #668DB9
                    tableview = QTableView()
                    tableview.setModel(tablemodel)
                    self.setGeometry(250, 200, 1050,450)
                    tableview.setColumnWidth(0,200);tableview.setColumnWidth(1,200)
                    tableview.horizontalHeader().setStretchLastSection(True)
                    tableview.horizontalHeader().setDefaultAlignment(Qt.AlignLeft)
                    self.setWindowIcon(QIcon(pPath+'\evs.png'))             
                    tableview.setSelectionBehavior(QTableView.SelectRows)
                    CancelBtn = QPushButton('Cancel')
                    buttonLayout = QHBoxLayout()
                    buttonLayout.setContentsMargins(0, 0, 0, 0)
                    buttonLayout.addStretch()
                    buttonLayout.addWidget(CancelBtn)                
                    self.verticalLayout = QVBoxLayout(self)
                    self.verticalLayout.addWidget(lbl)
                    self.verticalLayout.addWidget(tableview)
                    self.verticalLayout.addLayout(buttonLayout)
                    self.setLayout(self.verticalLayout)
                    CancelBtn.clicked.connect(self.closeEvent)
                def closeEvent(self,event):
                    self.deleteLater()
                    
            class ErrorTableModel(QAbstractTableModel):
                header_labels = ['Field Name', 'Field Value', 'Error Message']
                def __init__(self, datain,data2, parent=None,*args):
                    QAbstractTableModel.__init__(self, parent, *args)
                    self.msgList1=datain
                    self.toolTipText = data2
                    
                def headerData(self, section, orientation, role=Qt.DisplayRole):
                    if role == Qt.DisplayRole and orientation == Qt.Horizontal:
                        return self.header_labels[section]
                    if role==Qt.FontRole and orientation == Qt.Horizontal:
                        font=QFont();font.setBold(True);font.setPointSize(10)
                        return font
                    if role == Qt.ForegroundRole and orientation == Qt.Horizontal:
                        return QBrush(QColor('#EE2653'))
                    return QAbstractTableModel.headerData(self, section, orientation, role)
                
                def rowCount(self, parent):
                    return len(self.msgList1)
            
                def columnCount(self, parent):
                    return len(self.msgList1[0])
            
                def data(self, index, role):
                    val=self.toolTipText[index.row()]
                    if role==Qt.ToolTipRole: return val
                    if not index.isValid():
                        return QVariant()
                    elif role == Qt.BackgroundRole:
                        if index.row() % 2 == 0:
                            return QBrush(QColor('#DDEBF7')) 
                        else:
                            return QBrush(Qt.white)
                    elif role != Qt.DisplayRole:
                        return QVariant()
                    return QVariant(self.msgList1[index.row()][index.column()])
            if msgList1!=[]:
                self.dialogTextBrowser = ErrorChWindow()
                self.dialogTextBrowser.exec_()
            else: 
                QMessageBox.about(self, "Evalueserve-LEI Data Quality Assessment", "%s" % ('No Errors Found!!')) 
                '''msg=QMessageBox();msg.setText('There is no error for this official entity!!') ; msg.setWindowTitle("Evalueserve-LEI Data Quality Check")
                msg.setIcon(QMessageBox.Information);msg.setWindowIcon(QIcon('logo538.png'));msg.exec_()'''
                          
#-----------------Parent class table model to load and display the data---------------------------
class MyTableModel(QAbstractTableModel):
    header_labels = ['Key Field ID', 'O', 'OfficialEntityName','Error Count']
    def __init__(self, datain, parent=None, *args):        
        QAbstractTableModel.__init__(self, parent, *args)
        self.EntityLst = datain
   
    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self.header_labels[section]
        if role==Qt.FontRole and orientation == Qt.Horizontal:
            font=QFont();font.setBold(True);font.setPointSize(10)
            return font
        if role==Qt.BackgroundRole and orientation == Qt.Horizontal:
            return QBrush(QColor('#DDEBF7'))
        if role == Qt.ForegroundRole and orientation == Qt.Horizontal:
            return QBrush(QColor('#315EA5'))
        return QAbstractTableModel.headerData(self, section, orientation, role)
    
    def rowCount(self, parent):
        return len(self.EntityLst)

    def columnCount(self, parent):
        return len(self.EntityLst[0])

    def data(self, index, role):
        if not index.isValid():
            return QVariant()
        elif role == Qt.BackgroundRole:
            if index.row() % 2 == 0:
                return QBrush(QColor('#DDEBF7')) 
            else:
                return QBrush(Qt.white)
        elif role != Qt.DisplayRole:
            return QVariant()
        return QVariant(self.EntityLst[index.row()][index.column()])
        


#=====================Qualitycheck tool===============================

requests.utils.DEFAULT_CA_BUNDLE_PATH = os.path.join(os.path.dirname(os.path.realpath(sys.executable)), 'cacert.pem')
sheetdict = {};EntityLst=[];msgList=[];toolTip=[];global fName ;global val

def ValidateLEI(fundMg):
    try:
        url='https://www.gleif.org/api/v1/lei/search' 
        data = {"query": fundMg,"filters":[], 'rows': 25}
        res = requests.post(url,json.dumps(data),verify=requests.utils.DEFAULT_CA_BUNDLE_PATH)
        dict=json.loads(res.content)
        if dict["results"]!=[]:
            d=json.loads(dict["results"][0]["LeiJsonSnippet"])
            legalNm= d["lei:LEIRecord"]["lei:Entity"]["lei:LegalName"]
            if type(legalNm)==unicode:legalNm=legalNm.encode('utf-8')
            if type(fundMg)==unicode:fundMg=fundMg.encode('utf-8')
            if legalNm==fundMg:return lei
    except:pass     

#Validate ISIN from http://www.isincodes.net Website
def ValidateISIN(officialEnNm):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:28.0) Gecko/20100101 Firefox/28.0','Accept': 'application/json, text/javascript, */*; q=0.01',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8','X-Requested-With': 'XMLHttpRequest'}
    url = 'http://www.isincodes.net/action/s.php'
    values = {"search": officialEnNm}
    resp = requests.post(url,headers=headers,data=values)
    soup=BeautifulSoup(resp.content)
    if soup.find('table',{'class':'table table-striped table-hover table-condensed'})==None:return None
    else:        
        table=soup.find('table',{'class':'table table-striped table-hover table-condensed'})
        columns=table.findAll('tr')[0].findAll('td') #Get first row in table and then it's columns
        isinTxt= columns[0].text + '(' + columns[1].text +')'
        return isinTxt 

def GetColumnNo(sht,colName): 
 
    for c in range(sht.ncols):
        if sht.cell(0,c).value==colName:
            if appndCol==False:
                colLst.append(c)
            return c
        
def GetCellValue(sht,rw,colNo,isText): 
    
    try:
        if isText:val=sht.cell(rw,colNo).value.strip()
        else: val=sht.cell(rw,colNo).value
        return val
    except:
      try:  
        val=sht.cell(rw,colNo).value
        return val
      except:
          pass
#         val="" 
#    return val
#----------------------Read the data file selected by User--------------------------    
def GetUnmatchedInvoiceColNo(sht): 
    indx=0  ;filterDict={}  ;ReqDict={}
    global appndCol
    
    for rw in range(1,sht.nrows):
        appndCol=True
#       get column no using sample data file 
        payCol=GetColumnNo(sht,"Payment Method")
        reqCol=GetColumnNo(sht,"Requestor")
        invCol=GetColumnNo(sht,"Invoice")
        
        #To filter all the Payment Method that have Invoice as first word
        if sht.cell(rw,payCol).value.split(' ')[0]=="Invoice":
            requesterId=sht.cell(rw,reqCol).value
            if filterDict.has_key(requesterId):
                filterDict[requesterId].append({sht.cell(rw,invCol).value:rw})
            else:
                filterDict[requesterId]=[{sht.cell(rw,invCol).value:rw}]  
    appndCol=False            
    #filterDict is a list of dictionaries contains invoices for the same Email id('Requester':[{Invoice:RowNumber}]
    #To get all the identified requests with their row no that do not have same invoice number for single EMail ID
    uIdList=filterDict.keys() 
#    filterList may contain 1 invoice or many invoices with similar no, hence a loop iterating with i
    for id in uIdList:
        pkey=''
        for i in range(0,len(filterDict[id])):
            pKey=filterDict[id][i].keys()
            if len(filterDict[id])>1:
#                If more than one value for a single requestor,
                for k in range(0,len(filterDict[id])):
                    if i!=k:
                        if pKey==filterDict[id][k].keys(): 
#                            find if different invoice numbers are there 
                            isDup=False
                            indx+=1   
                if indx==0:
                    if ReqDict.has_key(id):ReqDict[id].append(filterDict[id][i].values()[0])
                        #colLst.append({id:filterDict[id][i].values()[0]}) #print "id=",id,"values=",filterDict[id][i].values();colLst.append(filterDict[id][i].values())
                    else:ReqDict[id]=[filterDict[id][i].values()[0],]
                else:indx=0
    return ReqDict  

def getOpenDiologTip():
    return 'Open an Input LSE Excel Workbook' 

def isEnglish(s):
    try:
        s.encode('ascii')
    except UnicodeEncodeError:
        return False
    else:
        return True
def SpaceBetweenWords(txt):
    lst=re.findall(r'\s{2,}', txt.strip())
    if lst!=[]: return True
    else: return False 
def CheckSpecialChars(txt):
    lst=re.findall(r'^[0-9a-zA-Z ()/\-\.,]+$',txt)
    if lst==[]:return true   
def look4_BRegistry(search_value, sheet) :
    
    for i in range (1,sheet.nrows) :     
        if search_value.replace(" ","") == sheet.cell(i,1).value.replace(" ",""):return sheet.cell(i,1).value
def look4_LegalForms(search_value, sheet) :
    for i in range (0,sheet.nrows) :
        if search_value in sheet.cell(i,0).value: 
#             return i
            return sheet.cell(i,0).value
            
#***Function to check logic for Legal Form Country Wise and get matched appbr if any
def look4_LFlogic(search_value,cntry, logicSht): 
    cList=[];LfrmList=[] ;filterList=[];abbrFF='';cIndex=0;iGet=''
    for rw in range(1,logicSht.nrows):  #to get the list of county,their Abbreviations and legalForm
        cList=[]
        for cl in range(0,3):
            if logicSht.cell(rw,cl).value is not None:cList.append(logicSht.cell(rw,cl).value.strip())
        LfrmList.append(cList)
    #To get the filtered list for particular country
    for inx,i in enumerate(LfrmList):        
        if str(i[0]).upper()==str(cntry).upper():cIndex=inx;filterList.append(i)
    search_list= search_value.split(' ')  #List of official entity split by space to find out abbreviation
    langCheck=logicSht.cell(cIndex+1,4).value
    #finding matched abbr in filterlist and it's corresponding value
    for fList in filterList: 
        spacCnt=len(re.findall(r' |/',fList[1]))
        if spacCnt>0:
            if fList[1] == search_value or str(fList[1]).upper().replace('.','') == str(search_value).upper().replace('.',''):
                iGet=fList[1];abbrFF=fList[2];break
        else:
            for abbr in search_list:
                if str(abbr).upper() ==str(fList[1]).upper() or str(abbr).upper().replace('.','')==str(fList[1]).upper().replace('.',''): iGet=abbr;abbrFF=fList[2];break
    return iGet,abbrFF,langCheck 

def IsfieldExist(fieldList,fldName): 
    if fldName in fieldList:return True
    else:return False     
#-----------------------Initializing the progress bar and calling the quality check set the progress value--------------------#
def SuperMainFn():
    start_time = datetime.datetime.now()
    app = QApplication(sys.argv)   
    fPath=sys.argv[1]
#    change of code
    '''
    make sht=entire data for csv
    '''
    wb =xlrd.open_workbook(fPath)    
    sht=wb.sheet_by_index(0) 
    maxRows=sht.nrows -1   
    if getattr(sys,'frozen',False):
        exePath=os.path.dirname(os.path.realpath(sys.executable))
    elif __file__:
        exePath=os.path.dirname(os.path.abspath(__file__))
    
#    exePath=os.path.dirname(os.path.realpath(sys.executable))
    bar = ProgressBar(maxRows,exePath)
    bar.show()
   
    QualityLogicCheck(wb,sht,bar,exePath)
    bar.close()
    main(app,maxRows,start_time)
    
class ProgressBar(QWidget):
    def __init__(self, total,exePath,parent=None):
        super(ProgressBar, self).__init__(parent)
        self.name_line = QLineEdit()
        self.resize(450, 100)
        self.progressbar = QProgressBar()
        self.progressbar.setGeometry(10, 35, 270, 21)
        self.setWindowIcon(QIcon(exePath +'\python.png'))
        self.progressbar.setMinimum(1)
        self.progressbar.setMaximum(total)
        self.label2 = QLabel('Initializing...........')
        self.label2.setGeometry(QRect(20, 50, 151, 16))
        self.label2.setMinimumSize(400,50)
        self.label2.setWordWrap(True)   # Added on 23-05-2016
        main_layout = QGridLayout()
        main_layout.addWidget(self.progressbar, 0, 0)
        main_layout.addWidget(self.label2)
        self.setLayout(main_layout)
        self.setWindowTitle('Execution Progress')       
        self.setMinimumSize(self.sizeHint()) # Added on 23-05-2016
        
    def startLoop(self,val):
        if val <= self.progressbar.maximum():
            self.progressbar.setValue(val)
            qApp.processEvents()                                 
                                             
def QualityLogicCheck(wb,sht,bar,exePath):
  try:  
      
    fPath1=exePath + '\DQ ENGINE 2017.xlsm'

    global EntityLst
    global colLst
    colLst=[]
    logicWB=xlrd.open_workbook(fPath1)
    businessSht=logicWB.sheet_by_name("Business Registries in UnaVista")
    LegalFrmSht=logicWB.sheet_by_name("Existing Legal Forms - UnaVista")
    logicSht=logicWB.sheet_by_name('Legal Form Country Wise Logics')
    elfSht=logicWB.sheet_by_name('ELF')
    cntrySht=logicWB.sheet_by_name('Country')
    
    dataSht=logicWB.sheet_by_name('Data');fieldList=[];InvoiceDict={}
#    invoicedict gives the different invoice numbers and tehir row nos for same requestor email id
    InvoiceDict=GetUnmatchedInvoiceColNo(sht)
#    get selected fields by user
    for r in range(1,dataSht.nrows):
        if dataSht.cell(r,4).value!='':fieldList.append(dataSht.cell(r,4).value)
           
    val=1;
    EnableCrawl=dataSht.cell(0,6).value  #print EnableCrawl,"fields=",fieldList     #\Get the crawler status 
    #Get the column no for required fields from the sample data sheet
    keyIdCol=GetColumnNo(sht,dataSht.cell(1,0).value);           oCol=GetColumnNo(sht,dataSht.cell(2,0).value)
    reqCol=GetColumnNo(sht,dataSht.cell(3,0).value);             PayMthdCol=GetColumnNo(sht,dataSht.cell(4,0).value)
    icol=GetColumnNo(sht,dataSht.cell(5,0).value);               leiStatusCol=GetColumnNo(sht,dataSht.cell(6,0).value)
    leiEvntCol=GetColumnNo(sht,dataSht.cell(7,0).value);         EntStatusCol=GetColumnNo(sht,dataSht.cell(8,0).value)
    EntEvntCol=GetColumnNo(sht,dataSht.cell(9,0).value);         payStatusCol=GetColumnNo(sht,dataSht.cell(10,0).value)
    preLouCol=GetColumnNo(sht,dataSht.cell(11,0).value);         leiCol=GetColumnNo(sht,dataSht.cell(12,0).value)
    officialEnNmCol=GetColumnNo(sht,dataSht.cell(13,0).value);   angEnNameCol=GetColumnNo(sht,dataSht.cell(14,0).value)
    cntryCol=GetColumnNo(sht,dataSht.cell(15,0).value);          LegalFormCol=GetColumnNo(sht,dataSht.cell(16,0).value)
    LegalFormTxtCol=GetColumnNo(sht,dataSht.cell(17,0).value);   fundMgCol=GetColumnNo(sht,dataSht.cell(18,0).value)
    HQAddressLine1Col=GetColumnNo(sht,dataSht.cell(19,0).value); HQAddressLine2Col=GetColumnNo(sht,dataSht.cell(20,0).value)
#    newly added-agnisha
    HQAddressLine3Col=GetColumnNo(sht,dataSht.cell(48,0).value); HQAddressLine4Col=GetColumnNo(sht,dataSht.cell(49,0).value)
    HQCityCol=GetColumnNo(sht,dataSht.cell(21,0).value);         HQCountyCol=GetColumnNo(sht,dataSht.cell(22,0).value)
    HQCntryCol=GetColumnNo(sht,dataSht.cell(23,0).value);        HQPostCodeCol=GetColumnNo(sht,dataSht.cell(24,0).value)
#    new add-agnisha
    HQAddRegCol=GetColumnNo(sht,dataSht.cell(50,0).value)
    LFaddress1Col=GetColumnNo(sht,dataSht.cell(25,0).value);     LFaddress2Col=GetColumnNo(sht,dataSht.cell(26,0).value)
    LFCityCol=GetColumnNo(sht,dataSht.cell(27,0).value);         LFCountyCol=GetColumnNo(sht,dataSht.cell(28,0).value)
    LFCountry1Col=GetColumnNo(sht,dataSht.cell(29,0).value);     LFPostCodeCol=GetColumnNo(sht,dataSht.cell(30,0).value)
    BRCountryCol=GetColumnNo(sht,dataSht.cell(31,0).value);      OBRegCol=GetColumnNo(sht,dataSht.cell(32,0).value)
    OBRegTxtCol=GetColumnNo(sht,dataSht.cell(33,0).value);       RegAuthEntIdCol=GetColumnNo(sht,dataSht.cell(34,0).value)
    bicCol=GetColumnNo(sht,dataSht.cell(35,0).value);            FRNCol=GetColumnNo(sht,dataSht.cell(36,0).value)
    ISINCol=GetColumnNo(sht,dataSht.cell(37,0).value);           LinkedIssuerCol=GetColumnNo(sht,dataSht.cell(38,0).value)
    LOUCol=GetColumnNo(sht,dataSht.cell(11,0).value) ;           #ICBSectorCol=GetColumnNo(sht,dataSht.cell(40,0).value)
    AssociatedLEICol=GetColumnNo(sht,dataSht.cell(41,0).value)     
    infoLUDtCol=GetColumnNo(sht,dataSht.cell(42,0).value);       ExpirtDtCol=GetColumnNo(sht,dataSht.cell(43,0).value)
    ARDateCol=GetColumnNo(sht,dataSht.cell(44,0).value);         SourcesCol=GetColumnNo(sht,dataSht.cell(45,0).value) 
    PaySet_DtCol=GetColumnNo(sht,dataSht.cell(46,0).value);      WebsiteCol=GetColumnNo(sht,dataSht.cell(47,0).value) 
#    new add
    DirPrntCol=GetColumnNo(sht,dataSht.cell(51,0).value);        DirExcptnRsnCol=GetColumnNo(sht,dataSht.cell(52,0).value)
    DirRelTypeCol=GetColumnNo(sht,dataSht.cell(53,0).value);     DirAccPrdStartCol=GetColumnNo(sht,dataSht.cell(54,0).value)
    DirFilngPrdStrtCol=GetColumnNo(sht,dataSht.cell(55,0).value);DirRelshpStatusCol=GetColumnNo(sht,dataSht.cell(58,0).value) 

    PNIbnsRegEnIDCol=GetColumnNo(sht,dataSht.cell(56,0).value);  DirLEICol=GetColumnNo(sht,dataSht.cell(57,0).value)
    DirValdnSrcsCol=GetColumnNo(sht,dataSht.cell(59,0).value);   DirVldnDocsCol=GetColumnNo(sht,dataSht.cell(60,0).value)
    DirRlnshpPrdStrtCol=GetColumnNo(sht,dataSht.cell(61,0).value)
    DirRlnQlfrCol=GetColumnNo(sht,dataSht.cell(62,0).value);     DirAccPrdEndCol=GetColumnNo(sht,dataSht.cell(63,0).value)
    DirRlnshpPrdEndCol=GetColumnNo(sht,dataSht.cell(64,0).value)
    PNIRegAuthIDCol=GetColumnNo(sht,dataSht.cell(65,0).value)
    DirFlngPrdEndCol=GetColumnNo(sht,dataSht.cell(66,0).value);  PNILglFrmnAddL1Col=GetColumnNo(sht,dataSht.cell(67,0).value)
    PNILglFormnAddCityCol=GetColumnNo(sht,dataSht.cell(68,0).value);
    PNILglFormnAddRgnCol=GetColumnNo(sht,dataSht.cell(69,0).value); PNILglFormnAddCntryCol=GetColumnNo(sht,dataSht.cell(70,0).value)
    PNILglFormnAddPstCodeCol=GetColumnNo(sht,dataSht.cell(71,0).value)
    PNIHQaddL1Col=GetColumnNo(sht,dataSht.cell(72,0).value);     PNIHQaddCityCol=GetColumnNo(sht,dataSht.cell(73,0).value)
    PNIHQaddRegCol=GetColumnNo(sht,dataSht.cell(74,0).value);    PNIHQaddPstCodeCol=GetColumnNo(sht,dataSht.cell(75,0).value)
    PNIHQaddCntryCol=GetColumnNo(sht,dataSht.cell(76,0).value);  UlExcptRsnCol=GetColumnNo(sht,dataSht.cell(77,0).value)   
    UlPrntCol=GetColumnNo(sht,dataSht.cell(78,0).value);         UlLEICol=GetColumnNo(sht,dataSht.cell(79,0).value)
    UlRlshpTypeCol=GetColumnNo(sht,dataSht.cell(80,0).value);    UlRlshpStatusCol=GetColumnNo(sht,dataSht.cell(81,0).value)
    UlVldnSrcsCol=GetColumnNo(sht,dataSht.cell(82,0).value);     UlVldnDocsCol=GetColumnNo(sht,dataSht.cell(83,0).value)
    UlRlshpQlfrCatCol=GetColumnNo(sht,dataSht.cell(84,0).value); UlAccPrdStrtCol=GetColumnNo(sht,dataSht.cell(85,0).value)
    UlAccPrdEndCol=GetColumnNo(sht,dataSht.cell(86,0).value);    UlRlshpPrdStrtCol=GetColumnNo(sht,dataSht.cell(87,0).value)
    UlRlshpPrdEndCol=GetColumnNo(sht,dataSht.cell(88,0).value);  UlFlngPrdStrtCol=GetColumnNo(sht,dataSht.cell(89,0).value)
    UlFlngPrdEndCol=GetColumnNo(sht,dataSht.cell(90,0).value);   PNI2LglFrmnAddL1Col=GetColumnNo(sht,dataSht.cell(91,0).value)
    PNI2LglFrmnAddCityCol=GetColumnNo(sht,dataSht.cell(92,0).value); PNI2LglFrmnAddRgnCol=GetColumnNo(sht,dataSht.cell(93,0).value)
    PNI2LglFrmnAddCntryCol=GetColumnNo(sht,dataSht.cell(94,0).value)
    PNI2LglFrmnAddPstCodeCol=GetColumnNo(sht,dataSht.cell(95,0).value)
    PNI2HQAddL1Col=GetColumnNo(sht,dataSht.cell(96,0).value);       PNI2HQAddCityCol=GetColumnNo(sht,dataSht.cell(97,0).value)
    PNI2HQAddRgnCol=GetColumnNo(sht,dataSht.cell(98,0).value);      PNI2HQAddPstCdCol=GetColumnNo(sht,dataSht.cell(99,0).value)
    PNI2HQAddCntryCol=GetColumnNo(sht,dataSht.cell(100,0).value);   PNI2RegAuthIDCol=GetColumnNo(sht,dataSht.cell(101,0).value)
    PNI2bsnsRegEnIDCol= GetColumnNo(sht,dataSht.cell(102,0).value); prevEnNamCol= GetColumnNo(sht,dataSht.cell(103,0).value) 
    altEnNameCol=GetColumnNo(sht,dataSht.cell(104,0).value);     EntCatCol=GetColumnNo(sht,dataSht.cell(105,0).value)
    HQAddAddNoCol=GetColumnNo(sht,dataSht.cell(106,0).value);    HQAddAddNoBldgCol=GetColumnNo(sht,dataSht.cell(107,0).value)
    LglFrmnAddAddNoCol=GetColumnNo(sht,dataSht.cell(108,0).value)
    VldnAuthIDcntryCol=GetColumnNo(sht,dataSht.cell(109,0).value)
    VlnAuthIdCol=GetColumnNo(sht,dataSht.cell(110,0).value);     OthrValAuthIdCol= GetColumnNo(sht,dataSht.cell(111,0).value)
    VlnAuthEnIdCol= GetColumnNo(sht,dataSht.cell(112,0).value);  CmntCol= GetColumnNo(sht,dataSht.cell(113,0).value)   
    CrtDateCol=GetColumnNo(sht,dataSht.cell(114,0).value);       lglFormnAddAddNoBldgCol= GetColumnNo(sht,dataSht.cell(115,0).value) 
    lglFrmnAddRgnCol=GetColumnNo(sht,dataSht.cell(116,0).value); dupChkResCol=  GetColumnNo(sht,dataSht.cell(117,0).value)   
    entLglFormCol=GetColumnNo(sht,dataSht.cell(118,0).value);    firstAssCol=GetColumnNo(sht,dataSht.cell(120,0).value)
    offBusRegRefCol=GetColumnNo(sht,dataSht.cell(119,0).value);  dirVldnRefCol=GetColumnNo(sht,dataSht.cell(121,0).value)
    ulVldnRefCol= GetColumnNo(sht,dataSht.cell(122,0).value);    othrAdd1TypeCol=GetColumnNo(sht,dataSht.cell(123,0).value)   
    othrAdd1AddAddNoCol=GetColumnNo(sht,dataSht.cell(124,0).value);othrAdd1AddAddNoWthnBldgCol=GetColumnNo(sht,dataSht.cell(125,0).value)
    othrAdd1AddLine1Col=GetColumnNo(sht,dataSht.cell(126,0).value);othrAdd1AddLine2Col=GetColumnNo(sht,dataSht.cell(127,0).value)
    othrAdd1AddLine3Col=GetColumnNo(sht,dataSht.cell(128,0).value);othrAdd1AddLine4Col=GetColumnNo(sht,dataSht.cell(129,0).value)
    othrAdd1AddCityCol=GetColumnNo(sht,dataSht.cell(130,0).value)
    othrAdd1AddCountyCol=GetColumnNo(sht,dataSht.cell(131,0).value);othrAdd1AddCntryCol=GetColumnNo(sht,dataSht.cell(132,0).value)
    othrAdd1AddRgnCol=GetColumnNo(sht,dataSht.cell(133,0).value);othrAdd1AddPostcodeCol=GetColumnNo(sht,dataSht.cell(134,0).value)
    othrAdd1AddMailRtngOnAddCol=GetColumnNo(sht,dataSht.cell(135,0).value);othrAdd2TypeCol=GetColumnNo(sht,dataSht.cell(136,0).value)
    othrAdd2AddAddNoCol=GetColumnNo(sht,dataSht.cell(137,0).value);othrAdd2AddAddNoWthnBldgCol=GetColumnNo(sht,dataSht.cell(138,0).value)
    othrAdd2AddLine1Col=GetColumnNo(sht,dataSht.cell(139,0).value);othrAdd2AddLine2Col=GetColumnNo(sht,dataSht.cell(140,0).value)
    othrAdd2AddLine3Col=GetColumnNo(sht,dataSht.cell(141,0).value);othrAdd2AddLine4Col=GetColumnNo(sht,dataSht.cell(142,0).value)
    othrAdd2AddCityCol=GetColumnNo(sht,dataSht.cell(143,0).value);othrAdd2AddCountyCol=GetColumnNo(sht,dataSht.cell(144,0).value)
    othrAdd2AddCountryCol=GetColumnNo(sht,dataSht.cell(145,0).value);othrAdd2AddRgnCol=GetColumnNo(sht,dataSht.cell(146,0).value)
    othrAdd2AddPostcodeCol=GetColumnNo(sht,dataSht.cell(147,0).value);othrAdd2AddMailRtngOnAddCol=GetColumnNo(sht,dataSht.cell(148,0).value)

   #Lists used to Check logic
    evtLst=['DISSOLVED','MERGER/ACQUISITION'];leiStatusLst=['MERGED','RETIRED']
    LegalLst=['FUND','SICAV SPECIALISED INVESTMENT FUND','UCITS - UNDERTAKINGS FOR COLLECTIVE INVESTMENT IN TRANSFERABLE SECURITIES','FONDS COMMUN DE PLACEMENT','ICVC - INVESTMENT COMPANY WITH VARIABLE CAPITAL']
    LglLst=['FUND', 'FONDO', 'FONDS', 'FONDOS', 'INVERSION', 'UCITS', 'ICVC', 'SICAV']
    lst= ['PENSION FUND','FUND','UCITS','TRUST', 'OTHER - PLEASE SPECIFY', 'PLEASE SPECIFY OTHER']
    LegalFormTxtLst=['SUB-FUND','UNIT TRUST','SICAV SPECIALISED INVESTMENT FUND','UCITS - UNDERTAKINGS FOR COLLECTIVE INVESTMENT IN TRANSFERABLE SECURITIES','FONDS COMMUN DE PLACEMENT','ICVC - INVESTMENT COMPANY WITH VARIABLE CAPITAL']
    
    #--------------------------loop through each excel row and get the cell value and associated error ---------------

    
    for rownum in range(1,sht.nrows):
        
        errMsg=[];wdLst='';wdLst1='';wdLst2='';toolTipLst=[];fd='';fd1='';fd2=''
        
#       --------------------------- spaces---------------------------
#        try:
#        for col in range(13,WebsiteCol):
        for col in colLst:
            if sht.cell(rownum,col).value!='' or sht.cell(rownum,col).value!=None : 
                feld=sht.cell(0,col).value #col name
               
                if type(sht.cell(rownum,col).value)==unicode:
                    txt=sht.cell(rownum,col).value.encode('utf-8')
                else:
                    txt=str(sht.cell(rownum,col).value)                
                if SpaceBetweenWords(txt)==True:
                    if wdLst==''and fd=='':
                        wdLst=wdLst+txt
                        fd=fd+feld
                    else:
                        wdLst=wdLst+','+txt
                        fd=fd+','+feld
                if txt[:1] ==' ':                    
                    if wdLst1=='':
                        wdLst1=wdLst1+txt
                        fd1=fd1+feld
                    else:
                        wdLst1=wdLst1+','+txt
                        fd1=fd1+','+feld
                if txt[-1:] ==' ': 
                    if wdLst2=='':
                        wdLst2=wdLst2+txt
                        fd2=fd2+feld
                    else:
                        wdLst2=wdLst2+','+txt
                        fd2=fd2+','+feld
                
        if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,'Extra space--')==True:
            if wdLst!='':errMsg.append([fd,wdLst,'Unnecessary space between words']);toolTipLst.append('Error Logic:'+'\n'+'Space between words should not be more than one')
            if wdLst1!='':errMsg.append([fd1,wdLst1,'Leading space in the field text']);toolTipLst.append('Error Logic:'+'\n'+'Fields should not start with space')
            if wdLst2!='':errMsg.append([fd2,wdLst2,'Trailing space in the field text']) ;toolTipLst.append('Error Logic:'+'\n'+'Fields should not end with space')
        
#       pick up values of each row , each field from sample data sheet 
        try:          
            keyId=int(GetCellValue(sht,rownum,keyIdCol,False)) 
        except:
            keyId=str(GetCellValue(sht,rownum,keyIdCol,False))
            
        o=GetCellValue(sht,rownum,oCol,True)
        invoice=GetCellValue(sht,rownum,icol,True);                     req=GetCellValue(sht,rownum,reqCol,True)  
        leiStatus=GetCellValue(sht,rownum,leiStatusCol,True);           leiEvnt=GetCellValue(sht,rownum,leiEvntCol,True)
        EntStatus=GetCellValue(sht,rownum,EntStatusCol,True);           EntEvnt=GetCellValue(sht,rownum,EntEvntCol,True)   
        payStatus=GetCellValue(sht,rownum,payStatusCol,True);           preLou=GetCellValue(sht,rownum,preLouCol,True)
        lei=GetCellValue(sht,rownum,leiCol,True);  
        officialEnNm=GetCellValue(sht,rownum,officialEnNmCol,True);     angEnName=GetCellValue(sht,rownum,angEnNameCol,True)
        cntry=GetCellValue(sht,rownum,cntryCol,True);                   LegalForm=GetCellValue(sht,rownum,LegalFormCol,True);  
        LegalFormTxt=GetCellValue(sht,rownum,LegalFormTxtCol,True);     fundMg=GetCellValue(sht,rownum,fundMgCol,True)
        HQAddressLine1=GetCellValue(sht,rownum,HQAddressLine1Col,True); HQAddressLine2=GetCellValue(sht,rownum,HQAddressLine2Col,True)
        HQAddressLine3=GetCellValue(sht,rownum,HQAddressLine3Col,True); HQAddressLine4=GetCellValue(sht,rownum,HQAddressLine4Col,True)
        HQCity=GetCellValue(sht,rownum,HQCityCol,True);                 HQCounty=GetCellValue(sht,rownum,HQCountyCol,True)
        HQCntry=GetCellValue(sht,rownum,HQCntryCol,True);               HQPostCode=GetCellValue(sht,rownum,HQPostCodeCol,True)
        HQAddReg=GetCellValue(sht,rownum,HQAddRegCol,True)
        LFaddress1=GetCellValue(sht,rownum,LFaddress1Col,True);         LFaddress2=GetCellValue(sht,rownum,LFaddress2Col,True)
        LFCity=GetCellValue(sht,rownum,LFCityCol,True);                 LFCounty=GetCellValue(sht,rownum,LFCountyCol,True)
        LFCountry1=GetCellValue(sht,rownum,LFCountry1Col,True);         LFPostCode=GetCellValue(sht,rownum,LFPostCodeCol,True)
        BRCountry=GetCellValue(sht,rownum,BRCountryCol,True);           OBReg=GetCellValue(sht,rownum,OBRegCol,True)
        OBRegTxt=GetCellValue(sht,rownum,OBRegTxtCol,True);             RegAuthEntId=GetCellValue(sht,rownum,RegAuthEntIdCol,True)
        bic=GetCellValue(sht,rownum,bicCol,True);                       FRN=GetCellValue(sht,rownum,FRNCol,True)
        ISIN=GetCellValue(sht,rownum,ISINCol,True);                     LinkedIssuer=GetCellValue(sht,rownum,LinkedIssuerCol,True) 
        LOU=GetCellValue(sht,rownum,LOUCol,True);                       #ICBSector=GetCellValue(sht,rownum,ICBSectorCol,True)  
        AssociatedLEI=GetCellValue(sht,rownum,AssociatedLEICol,True);   ExpiryDate=GetCellValue(sht,rownum,ExpirtDtCol,True)  
        ARDate=GetCellValue(sht,rownum,ARDateCol,True);                 PaySet_Dt=GetCellValue(sht,rownum,PaySet_DtCol,True) 
        Sources=GetCellValue(sht,rownum,SourcesCol,True);               Website=GetCellValue(sht,rownum,WebsiteCol,True) 
        infoLUDt=GetCellValue(sht,rownum,infoLUDtCol,True);        
        cntry=str(cntry).upper();uLegalFormTxt=str(LegalFormTxt).upper()
        DirPrnt=GetCellValue(sht,rownum,DirPrntCol,True);               DirExcptnRsn=GetCellValue(sht,rownum,DirExcptnRsnCol,True) 
        DirRelType=GetCellValue(sht,rownum,DirRelTypeCol,True);         DirAccPrdStart= GetCellValue(sht,rownum,DirAccPrdStartCol,True)
        DirFilngPrdStrt=GetCellValue(sht,rownum,DirFilngPrdStrtCol,True) 
        DirRelshpStatus=GetCellValue(sht,rownum,DirRelshpStatusCol,True); PNIbnsRegEnID=GetCellValue(sht,rownum,PNIbnsRegEnIDCol,True)
        DirRlnshpPrdStrt=GetCellValue(sht,rownum,DirRlnshpPrdStrtCol,True)
        DirLEI=GetCellValue(sht,rownum,DirLEICol,True)
        DirValdnSrcs=GetCellValue(sht,rownum,DirValdnSrcsCol,True);     DirVldnDocs=GetCellValue(sht,rownum,DirVldnDocsCol,True)
        DirRlnQlfr=GetCellValue(sht,rownum,DirRlnQlfrCol,True);         DirAccPrdEnd=GetCellValue(sht,rownum, DirAccPrdEndCol,True)
        DirRlnshpPrdEnd=GetCellValue(sht,rownum,DirRlnshpPrdEndCol,True)
        PNIRegAuthID=GetCellValue(sht,rownum,PNIRegAuthIDCol,True);     DirFlngPrdEnd=GetCellValue(sht,rownum,DirFlngPrdEndCol,True)
        PNILglFrmnAddL1=GetCellValue(sht,rownum,PNILglFrmnAddL1Col,True);
        PNILglFormnAddCity=GetCellValue(sht,rownum,PNILglFormnAddCityCol,True)
        PNILglFormnAddRgn=GetCellValue(sht,rownum,PNILglFormnAddRgnCol,True)
        PNILglFormnAddCntry=GetCellValue(sht,rownum,PNILglFormnAddCntryCol,True)
        PNILglFormnAddPstCode=GetCellValue(sht,rownum,PNILglFormnAddPstCodeCol,True)
        PNIHQaddL1=GetCellValue(sht,rownum,PNIHQaddL1Col,True);         PNIHQaddCity=GetCellValue(sht,rownum,PNIHQaddCityCol,True)
        PNIHQaddReg=GetCellValue(sht,rownum,PNIHQaddRegCol,True);       PNIHQaddPstCode=GetCellValue(sht,rownum,PNIHQaddPstCodeCol,True)
        PNIHQaddCntry=GetCellValue(sht,rownum,PNIHQaddCntryCol,True);   DirVldnDocs= GetCellValue(sht,rownum,DirVldnDocsCol,True) 
        UlExcptRsn=GetCellValue(sht,rownum,UlExcptRsnCol,True);         UlPrnt=GetCellValue(sht,rownum,UlPrntCol,True)        
        UlLEI=GetCellValue(sht,rownum,UlLEICol,True);                   UlRlshpType=GetCellValue(sht,rownum,UlRlshpTypeCol,True)
        UlRlshpStatus=GetCellValue(sht,rownum,UlRlshpStatusCol,True);   UlVldnSrcs= GetCellValue(sht,rownum,UlVldnSrcsCol,True)
        UlVldnDocs=GetCellValue(sht,rownum,UlVldnDocsCol,True);         UlRlshpQlfrCat=GetCellValue(sht,rownum, UlRlshpQlfrCatCol,True)  
        UlAccPrdStrt=GetCellValue(sht,rownum,UlAccPrdStrtCol,True);     UlAccPrdEnd=GetCellValue(sht,rownum,UlAccPrdEndCol,True)
        UlRlshpPrdStrt=GetCellValue(sht,rownum,UlRlshpPrdStrtCol,True); UlRlshpPrdEnd=GetCellValue(sht,rownum,UlRlshpPrdEndCol,True)
        UlFlngPrdStrt=GetCellValue(sht,rownum,UlFlngPrdStrtCol,True);   PNI2bsnsRegEnID=GetCellValue(sht,rownum,PNI2bsnsRegEnIDCol,True)
        UlFlngPrdEnd=GetCellValue(sht,rownum,UlFlngPrdEndCol,True);     PNI2LglFrmnAddL1=GetCellValue(sht,rownum,PNI2LglFrmnAddL1Col,True);
        PNI2LglFrmnAddCity=GetCellValue(sht,rownum,PNI2LglFrmnAddCityCol,True); PNI2LglFrmnAddRgn=GetCellValue(sht,rownum,PNI2LglFrmnAddRgnCol,True)
        PNI2LglFrmnAddCntry=GetCellValue(sht,rownum,PNI2LglFrmnAddCntryCol,True)
        PNI2LglFrmnAddPstCode=GetCellValue(sht,rownum,PNI2LglFrmnAddPstCodeCol,True)
        PNI2HQAddL1=GetCellValue(sht,rownum,PNI2HQAddL1Col,True);       PNI2HQAddCity=GetCellValue(sht,rownum,PNI2HQAddCityCol,True);
        PNI2HQAddRgn=GetCellValue(sht,rownum,PNI2HQAddRgnCol,True);     PNI2HQAddPstCd=GetCellValue(sht,rownum,PNI2HQAddPstCdCol,True)
        PNI2HQAddCntry=GetCellValue(sht,rownum,PNI2HQAddCntryCol,True); PNI2RegAuthID=GetCellValue(sht,rownum,PNI2RegAuthIDCol,True)
        prevEnNam=GetCellValue(sht,rownum,prevEnNamCol,True);           altEnName= GetCellValue(sht,rownum,altEnNameCol,True)
        EntCat= GetCellValue(sht,rownum,EntCatCol,True);                HQAddAddNo= GetCellValue(sht,rownum,HQAddAddNoCol,True)      
        HQAddAddNoBldg= GetCellValue(sht,rownum,HQAddAddNoBldgCol,True);VldnAuthIDcntry= GetCellValue(sht,rownum,VldnAuthIDcntryCol,True) 
        LglFrmnAddAddNo=GetCellValue(sht,rownum,LglFrmnAddAddNoCol,True)
        VlnAuthId=GetCellValue(sht,rownum,VlnAuthIdCol,True);           OthrValAuthId=GetCellValue(sht,rownum,OthrValAuthIdCol,True) 
        VlnAuthEnId=GetCellValue(sht,rownum,VlnAuthEnIdCol,True);       Cmnt= GetCellValue(sht,rownum,CmntCol,True)  
        CrtDate=GetCellValue(sht,rownum,CrtDateCol,True) ;              lglFormnAddAddNoBldg= GetCellValue(sht,rownum,lglFormnAddAddNoBldgCol,True)            
        lglFrmnAddRgn=GetCellValue(sht,rownum,lglFrmnAddRgnCol,True);   dupChkRes=GetCellValue(sht,rownum,dupChkResCol,True)
        entLglForm=GetCellValue(sht,rownum,entLglFormCol,True);         firstAss=GetCellValue(sht,rownum,firstAssCol,True)
        offBusRegRef=GetCellValue(sht,rownum,offBusRegRefCol,True);     dirVldnRef=GetCellValue(sht,rownum,dirVldnRefCol,True)
        ulVldnRef=GetCellValue(sht,rownum,ulVldnRefCol,True);           othrAdd1Type= GetCellValue(sht,rownum,othrAdd1TypeCol,True)
        othrAdd1AddAddNo=GetCellValue(sht,rownum,othrAdd1AddAddNoCol,True); othrAdd1AddAddNoWthnBldg=GetCellValue(sht,rownum,othrAdd1AddAddNoWthnBldgCol,True)
        othrAdd1AddLine1=GetCellValue(sht,rownum,othrAdd1AddLine1Col,True);othrAdd1AddLine2=GetCellValue(sht,rownum,othrAdd1AddLine2Col,True)
        othrAdd1AddLine3=GetCellValue(sht,rownum,othrAdd1AddLine3Col,True);othrAdd1AddLine4=GetCellValue(sht,rownum,othrAdd1AddLine4Col,True)
        othrAdd1AddCity=GetCellValue(sht,rownum,othrAdd1AddCityCol,True);othrAdd1AddCounty=GetCellValue(sht,rownum,othrAdd1AddCountyCol,True)
        othrAdd1AddCntry=GetCellValue(sht,rownum,othrAdd1AddCntryCol,True);othrAdd1AddRgn=GetCellValue(sht,rownum,othrAdd1AddRgnCol,True)
        othrAdd1AddPostcode=GetCellValue(sht,rownum,othrAdd1AddPostcodeCol,True);othrAdd1AddMailRtngOnAdd=GetCellValue(sht,rownum,othrAdd1AddMailRtngOnAddCol,True)
        othrAdd2Type=GetCellValue(sht,rownum,othrAdd2TypeCol,True);othrAdd2AddAddNo=GetCellValue(sht,rownum,othrAdd2AddAddNoCol,True)
        othrAdd2AddAddNoWthnBldg=GetCellValue(sht,rownum,othrAdd2AddAddNoWthnBldgCol,True);othrAdd2AddLine1=GetCellValue(sht,rownum,othrAdd2AddLine1Col,True)
        othrAdd2AddLine2=GetCellValue(sht,rownum,othrAdd2AddLine2Col,True);othrAdd2AddLine3=GetCellValue(sht,rownum,othrAdd2AddLine3Col,True)
        othrAdd2AddLine4=GetCellValue(sht,rownum,othrAdd2AddLine4Col,True);othrAdd2AddCity=GetCellValue(sht,rownum,othrAdd2AddCityCol,True)
        othrAdd2AddCountry=GetCellValue(sht,rownum,othrAdd2AddCountryCol,True);othrAdd2AddRgn=GetCellValue(sht,rownum,othrAdd2AddRgnCol,True)
        othrAdd2AddPostcode=GetCellValue(sht,rownum,othrAdd2AddPostcodeCol,True);othrAdd2AddMailRtngOnAdd=GetCellValue(sht,rownum,othrAdd2AddMailRtngOnAddCol,True)
        othrAdd2AddCounty=GetCellValue(sht,rownum,othrAdd2AddCountyCol,True)
        
        
        
        
        try:
            #############-----------------Check the logic for each field and append the message to errMsg list-----------------------
            '''
#           --------------------------------LOU---------------------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(11,0).value)==True:
              try: 
                if str(LOU).upper()!='IEI': 
                    if EntStatus!=None:
                        errMsg.append(['LOU',LOU,'Pre-LOU is not "IEI"'])
                        toolTipLst.append('Error Logic:'+'\n'+'Pre-LOU is not "IEI"')
              except:
                  pass 
            ''' 
            
#           ---------------------------------BusinessRegistryCountry-----------------------added 29th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(31,0).value)==True:
#              logic1               
                try:
                    if BRCountry=="":
                        errMsg.append(["BusinessRegistryCountry",BRCountry,"BusinessRegistryCountry should not be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"BusinessRegistryCountry should not be blank.")
                except:
                    pass
                
#                logic2
                try:
                    if BRCountry.strip().lower()!=cntry.strip().lower():
                        errMsg.append(["BusinessRegistryCountry",BRCountry,"BusinessRegistryCountry should match with text in 'CountryLegalForm'."])
                        toolTipLst.append('Error Logic:'+'\n'+"BusinessRegistryCountry should match with text in 'CountryLegalForm'.")
                except:
                    pass
                try:
                    if BRCountry.strip().lower()!=LFCountry1.strip().lower():
                        errMsg.append(["BusinessRegistryCountry",BRCountry,"BusinessRegistryCountry should match with text in 'LegalFormationAddressCountry'."])
                        toolTipLst.append('Error Logic:'+'\n'+"BusinessRegistryCountry should match with text in 'LegalFormationAddressCountry'.")
                except:
                    pass 
                
                
#           ------------------------------OtherAddress1Type--------------------------------new add 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(123,0).value)==True:
                try:
                    if othrAdd1Type!="":
                        errMsg.append(["OtherAddress1Type",othrAdd1Type,"OtherAddress1Type should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1Type should be blank.")
                except:
                    pass
                
#           ----------------------------------OtherAddress1AddressAddressNumber----------------------new add 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(124,0).value)==True:
                try:
                    if othrAdd1AddAddNo!="":
                        errMsg.append(["OtherAddress1AddressAddressNumber",othrAdd1AddAddNo,"OtherAddress1AddressAddressNumber should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressAddressNumber should be blank.")
                except:
                    pass
            
#            ----------------------------------OtherAddress1AddressAddressNumberWithinBuilding----------------------new add 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(125,0).value)==True:
                try:
                    if othrAdd1AddAddNoWthnBldg!="":
                        errMsg.append(["OtherAddress1AddressAddressNumber",othrAdd1AddAddNoWthnBldg,"OtherAddress1AddressAddressNumber should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressAddressNumber should be blank.")
                except:
                    pass
                
#           ----------------------------------OtherAddress1AddressLine1-------------------new add 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(126,0).value)==True:
                try:
                    if othrAdd1AddLine1!="":
                        errMsg.append(["OtherAddress1AddressLine1",othrAdd1AddLine1,"OtherAddress1AddressLine1 should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressLine1 should be blank.")
                except:
                    pass     
     
#           ------------------------------OtherAddress1AddressLine2-------------------new add 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(127,0).value)==True:
                try:
                    if othrAdd1AddLine2!="":
                        errMsg.append(["OtherAddress1AddressLine1",othrAdd1AddLine2,"OtherAddress1AddressLine1 should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressLine1 should be blank.")
                except:
                    pass 
     
#           ------------------------------OtherAddress1AddressLine3-------------------new add 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(128,0).value)==True:
                try:
                    if othrAdd1AddLine3!="":
                        errMsg.append(["OtherAddress1AddressLine3",othrAdd1AddLine3,"OtherAddress1AddressLine3 should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressLine3 should be blank.")
                except:
                    pass 
                
#           ------------------------------OtherAddress1AddressLine4-------------------new add 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(129,0).value)==True:
                try:
                    if othrAdd1AddLine4!="":
                        errMsg.append(["OtherAddress1AddressLine4",othrAdd1AddLine4,"OtherAddress1AddressLine4 should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressLine4 should be blank.")
                except:
                    pass 
                
#            ------------------------------OtherAddress1AddressCity-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(130,0).value)==True:
                try:
                    if othrAdd1AddCity!="":
                        errMsg.append(["OtherAddress1AddressCity",othrAdd1AddCity,"OtherAddress1AddressCity should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressCity should be blank.")
                except:
                    pass    
            
#            ------------------------------OtherAddress1AddressCounty-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(131,0).value)==True:
                try:
                    if othrAdd1AddCounty!="":
                        errMsg.append(["OtherAddress1AddressCounty",othrAdd1AddCounty,"OtherAddress1AddressCounty should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressCounty should be blank.")
                except:
                    pass 
                
#           ------------------------------OtherAddress1AddressCountry-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(132,0).value)==True:
                try:
                    if othrAdd1AddCntry!="":
                        errMsg.append(["OtherAddress1AddressCountry",othrAdd1AddCntry,"OtherAddress1AddressCountry should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressCountry should be blank.")
                except:
                    pass     
             
#             ------------------------------OtherAddress1AddressRegion-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(133,0).value)==True:
                try:
                    if othrAdd1AddRgn!="":
                        errMsg.append(["OtherAddress1AddressRegion",othrAdd1AddRgn,"OtherAddress1AddressRegion should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressRegion should be blank.")
                except:
                    pass  
                
#             ------------------------------OtherAddress1AddressPostcode-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(134,0).value)==True:
                try:
                    if othrAdd1AddPostcode!="":
                        errMsg.append(["OtherAddress1AddressPostcode",othrAdd1AddPostcode,"OtherAddress1AddressPostcode should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressPostcode should be blank.")
                except:
                    pass     
                
#             ------------------------------OtherAddress1AddressMailRoutingOnAddresses-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(135,0).value)==True:
                try:
                    if othrAdd1AddMailRtngOnAdd!="":
                        errMsg.append(["OtherAddress1AddressMailRoutingOnAddresses",othrAdd1AddMailRtngOnAdd,"OtherAddress1AddressMailRoutingOnAddresses should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress1AddressMailRoutingOnAddresses should be blank.")
                except:
                    pass      
           
#            ------------------------------OtherAddress2Type-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(136,0).value)==True:
                try:
                    if othrAdd2Type!="":
                        errMsg.append(["OtherAddress2Type",othrAdd2Type,"OtherAddress2Type should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2Type should be blank.")
                except:
                    pass    
           
#             ------------------------------OtherAddress2AddressAddressNumber-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(137,0).value)==True:
                try:
                    if othrAdd2AddAddNo!="":
                        errMsg.append(["OtherAddress2AddressAddressNumber",othrAdd2AddAddNo,"OtherAddress2AddressAddressNumber should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressAddressNumber should be blank.")
                except:
                    pass 
           
#           ------------------------------OtherAddress2AddressAddressNumberWithinBuilding-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(138,0).value)==True:
                try:
                    if othrAdd2AddAddNoWthnBldg!="":
                        errMsg.append(["OtherAddress2AddressAddressNumberWithinBuilding",othrAdd2AddAddNoWthnBldg,"OtherAddress2AddressAddressNumberWithinBuilding should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressAddressNumberWithinBuilding should be blank.")
                except:
                    pass 
                
#           ------------------------------OtherAddress2AddressLine1-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(139,0).value)==True:
                try:
                    if othrAdd2AddLine1!="":
                        errMsg.append(["OtherAddress2AddressLine1",othrAdd2AddLine1,"OtherAddress2AddressLine1 should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressLine1 should be blank.")
                except:
                    pass    
                
#            ------------------------------OtherAddress2AddressLine2-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(140,0).value)==True:
                try:
                    if othrAdd2AddLine2!="":
                        errMsg.append(["OtherAddress2AddressLine2",othrAdd2AddLine2,"OtherAddress2AddressLine2 should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressLine2 should be blank.")
                except:
                    pass     
           
           
#             ------------------------------OtherAddress2AddressLine3-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(141,0).value)==True:
                try:
                    if othrAdd2AddLine3!="":
                        errMsg.append(["OtherAddress2AddressLine3",othrAdd2AddLine3,"OtherAddress2AddressLine3 should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressLine3 should be blank.")
                except:
                    pass  
           
#            ------------------------------OtherAddress2AddressLine4-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(142,0).value)==True:
                try:
                    if othrAdd2AddLine4!="":
                        errMsg.append(["OtherAddress2AddressLine4",othrAdd2AddLine4,"OtherAddress2AddressLine4 should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressLine4 should be blank.")
                except:
                    pass 
           
#             ------------------------------OtherAddress2AddressCity-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(143,0).value)==True:
                try:
                    if othrAdd2AddLine4!="":
                        errMsg.append(["OtherAddress2AddressCity",othrAdd2AddCity,"OtherAddress2AddressCity should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressCity should be blank.")
                except:
                    pass 
           
#           ------------------------------OtherAddress2AddressCounty-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(144,0).value)==True:
                try:
                    if othrAdd2AddCounty!="":
                        errMsg.append(["OtherAddress2AddressCounty",othrAdd2AddCounty,"OtherAddress2AddressCounty should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressCounty should be blank.")
                except:
                    pass 
           
#            ------------------------------OtherAddress2AddressCountry-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(145,0).value)==True:
                try:
                    if othrAdd2AddCountry!="":
                        errMsg.append(["OtherAddress2AddressCountry",othrAdd2AddCountry,"OtherAddress2AddressCountry should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressCountry should be blank.")
                except:
                    pass
                
#           ------------------------------OtherAddress2AddressRegion-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(146,0).value)==True:
                try:
                    if othrAdd2AddRgn!="":
                        errMsg.append(["OtherAddress2AddressRegion",othrAdd2AddRgn,"OtherAddress2AddressRegion should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressCountry should be blank.")
                except:
                    pass   
                
             
#            ------------------------------OtherAddress2AddressPostcode-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(147,0).value)==True:
                try:
                    if othrAdd2AddPostcode!="":
                        errMsg.append(["OtherAddress2AddressPostcode",othrAdd2AddPostcode,"OtherAddress2AddressPostcode should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressPostcode should be blank.")
                except:
                    pass  
                
#             ------------------------------OtherAddress2AddressMailRoutingOnAddresses-------------------new add 28th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(148,0).value)==True:
                try:
                    if othrAdd2AddMailRtngOnAdd!="":
                        errMsg.append(["OtherAddress2AddressMailRoutingOnAddresses",othrAdd2AddMailRtngOnAdd,"OtherAddress2AddressMailRoutingOnAddresses should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"OtherAddress2AddressMailRoutingOnAddresses should be blank.")
                except:
                    pass
            
#           --------------------------------DirectValidationDocuments--------------removed 3/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(60,0).value)==True:
                try:
                    if "REGULATORY_FILING" in DirVldnDocs.upper():
                        if DirValdnSrcs.upper()=="ENTITY_SUPPLIED_ONLY":
                           errMsg.append(['Entity Legal Form',entLglForm,"If Validation Documents has 'REGULATORY_FILING', then 'DirectValidationSources' cannot be ENTITY_SUPPLIED_ONLY."])
                           toolTipLst.append('Error Logic:'+'\n'+"If Validation Documents has 'REGULATORY_FILING', then 'DirectValidationSources' cannot be ENTITY_SUPPLIED_ONLY.") 
                except:
                    pass
                
                try:
                    if "ACCOUNTS_FILING" in DirVldnDocs.upper():
                        if DirAccPrdStart=="":
                            errMsg.append(['Entity Legal Form',entLglForm,"If Validation Documents has 'ACCOUNTS_FILING ', then DirectAccountingPeriodStart is not blank."])
                            toolTipLst.append('Error Logic:'+'\n'+"If Validation Documents has 'ACCOUNTS_FILING ', then DirectAccountingPeriodStart is not blank.")
                except:
                    pass
            '''        
            '''                
#           -----------------------------Entity Legal Form--------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(118,0).value)==True:
              
              try: 
                if str(entLglForm).lstrip().rstrip()=='8888 - Other Please Specify': 
                    if LegalFormTxt=="":
                        errMsg.append(['Entity Legal Form',entLglForm,'If Entity Legal Form is "8888 - Other Please Specify", then LegalFormFreeText should not be blank.'])
                        toolTipLst.append('Error Logic:'+'\n'+'If Entity Legal Form is "8888 - Other Please Specify", then LegalFormFreeText should not be blank.')
              except:
                  pass 
            '''      
              
#           --------------------------------LEI Status----------updated 25-06-2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(6,0).value)==True:
              try: 
                if leiStatus!=None: 
                    '''
#                    logic1
                    if leiStatus.upper()!="ANNULLED":
                        if officialEnNm=="":
                           errMsg.append(['LEI Status',leiStatus,"If 'LEI Status' is not 'ANNULLED', then 'Offocial Entity Name' should not be blank."])
                           toolTipLst.append('Error Logic:'+'\n'+"If 'LEI Status' is not 'ANNULLED', then 'Offocial Entity Name' should not be blank.") 
                    '''       
#                     logic1
                    if leiStatus.upper().strip()=="ISSUED":
                        if EntStatus.upper()=="ACTIVE":
                            if ExpiryDate!="":
                               errMsg.append(['LEI Status',leiStatus,"If LEI Status is ISSUED and Entity Status is active, then expiry date must be blank."])
                               toolTipLst.append('Error Logic:'+'\n'+"If LEI Status is ISSUED and Entity Status is active, then expiry date must be blank.") 

#                    logic 2
                    if leiStatus.upper().strip()=="ISSUED" or leiStatus.upper().strip()=="LAPSED" or leiStatus.upper().strip()=="PENDING_TRANSFER" or leiStatus.upper().strip()=="PENDING_ARCHIVAL":
                        if EntStatus.upper().strip()!="ACTIVE":
                            errMsg.append(['LEI Status',leiStatus,"If LEI status is ISSUED OR LAPSED OR PENDING_TRANSFER OR PENDING_ARCHIVAL then Entity Status should be ACTIVE."])
                            toolTipLst.append('Error Logic:'+'\n'+"If LEI status = ISSUED OR LAPSED OR PENDING_TRANSFER OR PENDING_ARCHIVAL then Entity Status should be ACTIVE.") 
                     
#                    logic 3
                    if leiStatus.upper().strip()=="INACTIVE":
                         if not(EntEvnt.lower().strip()=="dissolved" and leiEvnt.lower().strip()=="dissolved" and EntStatus.lower().strip()=="dissolved"):
                             errMsg.append(['LEI Status',leiStatus,"If LEI Status is 'Inactive' then Entity Event, Lei Event and Entity Status should be 'Dissolved', 'Dissolved' and 'Inactive' respectively."])
                             toolTipLst.append('Error Logic:'+'\n'+"If LEI Status is 'Inactive' then Entity Event, Lei Event and Entity Status should be 'Dissolved', 'Dissolved' and 'Inactive' respectively.") 
                    
                     
#                     logic 4
                    if leiStatus.upper().strip()=="MERGED":
                       if not(EntEvnt.strip()=="Merger/Acquisition" and leiEvnt.strip()==" Merger/Acquisition" and EntStatus.strip()=="Inactive"):
                           errMsg.append(['LEI Status',leiStatus,"If LEI Status is 'Merged' then Entity Event, Lei Event and Entity Status should be 'Merger/Acquisition', 'Merger/Acquisition' and 'Inactive' respectively."])
                           toolTipLst.append('Error Logic:'+'\n'+"If LEI Status is 'Merged' then Entity Event, Lei Event and Entity Status should be 'Merger/Acquisition', 'Merger/Acquisition' and 'Inactive' respectively.")
                    
                    ''' 
#                    logic 3                     
                    if EntStatus.upper()=="ACTIVE":
                        if not(leiStatus.upper()=="ISSUED" or leiStatus.upper()=="LAPSED" or leiStatus.upper()=="PENDING_TRANSFER" or leiStatus.upper()=="PENDING_ARCHIVAL"):
                            errMsg.append(['LEI Status',leiStatus,"If Entity Status is 'ACTIVE', then LEI status should be ISSUED, LAPSED, PENDING_TRANSFER or PENDING_ARCHIVAL."])
                            toolTipLst.append('Error Logic:'+'\n'+"If Entity Status is 'ACTIVE', then LEI status should be ISSUED, LAPSED, PENDING_TRANSFER or PENDING_ARCHIVAL.") 
                              
#                    logic 4
                    if EntStatus!=None:
                        if str(leiStatus).upper()=='RETIRED' or str( leiStatus).upper()=='MERGED':  
                            if str(EntStatus).upper()!='INACTIVE':errMsg.append(['LEI Status',leiStatus,'LEI Status is '+leiStatus+', so Entity Status must be "Inactive"']);toolTipLst.append('Error Logic:'+'\n'+'If LEI Status is Retired or Merged then Entity Status must be Inactive')
                  
                    if AssociatedLEI!=None:
                        if str(leiStatus).upper()=='DUPLICATE':
                            if AssociatedLEI=='':errMsg.append(['LEI Status',leiStatus,'LEI Status is '+leiStatus+', so AssociatedLEI should be populated']);toolTipLst.append('Error Logic:'+'\n'+'if LEI Status is Duplicate then AssociatedLEI should be populated')
                    '''       
              except:
                  pass
     
#           -------------------------------LegalFormationAddressAddressNumberWithinBuilding---------------updated 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(115,0).value)==True:
                try:
                    if lglFormnAddAddNoBldg!="":
                        errMsg.append(['LegalFormationAddressAddressNumberWithinBuilding',lglFormnAddAddNoBldg,"LegalFormationAddressAddressNumberWithinBuilding should be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"LegalFormationAddressAddressNumberWithinBuilding should be blank.") 
                        
                except:
                    pass
#           -------------------------------Payment Status-----------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(10,0).value)==True:
                try: 
                    if payStatus.lower()=="complete" :
                        if "awaiting validation" in leiEvnt.lower() or "on hold" in leiEvnt.lower():
                              errMsg.append(['Payment Status',payStatus,"If Payment Status is COMPLETE, then LEI event should not be 'Awaiting validation' or 'On hold'."])
                              toolTipLst.append('Error Logic:'+'\n'+"If Payment Status is COMPLETE, then LEI event should not be 'Awaiting validation' or 'On hold'.")  
                except:
                     pass
                 
#           --------------------------------FundManager---------------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(18,0).value)==True:
                try:  
                    if '8888 - Other Please Specify' in entLglForm:
                       if "pension fund" not in LegalFormTxt.lower().strip(): 
                          if "fund" in LegalFormTxt.lower().strip() or "fondo" in LegalFormTxt.lower().strip() or "fonds" in LegalFormTxt.lower().strip() or "fondos" in LegalFormTxt.lower().strip() or "inversion" in LegalFormTxt.lower().strip() or "ucits" in LegalFormTxt.lower().strip() or "icvc" in LegalFormTxt.lower().strip() or "sicav" in LegalFormTxt.lower().strip():
                            if fundMg=="" or fundMg==None:
                                errMsg.append(['Fund Manager',fundMg,'If Entity LegalForm is "8888 - Other Please Specify" and LegalFormFreeText has "Fund" then Fund Manager should not be blank.'])
                                toolTipLst.append('Error Logic:'+'\n'+'If Entity Legal Form is "8888 - Other Please Specify" and LegalFormFreeText has "Fund" then Fund Manager should not be blank.')
                            if EntCat.lower()!="fund":
                                    errMsg.append(['EntityCategory',EntCat,'If Entity Legal Form = 8888 - Other Please Specify and Specify Other has the word Fund, Fondo, Fonds, Fondos, Inversion, UCITS, ICVC, SICAV except pension funds ,FundManager is not be blank, then Entity Category should be "Fund".'])
                                    toolTipLst.append('Error Logic:'+'\n'+'If Entity Legal Form = 8888 - Other Please Specify and Specify Other has the word Fund, Fondo, Fonds, Fondos, Inversion, UCITS, ICVC, SICAV except pension funds ,FundManager is not be blank, then Entity Category should be "Fund".')
                except:
                    pass 
            
            '''
#           ---------------------------------Invoice-----------------------------------                  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(5,0).value)==True:
              try:  
                requesterId=sht.cell(rownum,GetColumnNo(sht,"Requestor")).value
                
                invLst=[];l=[]
                if requesterId in InvoiceDict:   #First check if key is present in dic or not
                    invLst=InvoiceDict[requesterId]
                    
                    for n in invLst:l.append(sht.cell(n,icol).value)
                    for inv in invLst:
                        if inv==rownum:
                            errMsg.append(['Invoice',invoice,'Invoices : ' + ','.join(l) + '  should be same for the requester "'+requesterId+'"']);toolTipLst.append('Error Logic:'+'\n'+'If the first word in the Payment Method field is Invoice' +'\n'+ 'then for corresponding Requester, all the invoices should be same.')
              except:
                  pass
            ''' 
#           ----------------------------DirectRelationshipType-------------------------------updated 3/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(53,0).value)==True:
               if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A")  : 
                  try:
                      if DirRelType.strip()=="IS_ULTIMATELY_CONSOLIDATED_BY":
                         errMsg.append(['DirectRelationType',DirRelType,'DirectParent should not be "IS_ULTIMATELY_CONSOLIDATED_BY".'])
                         toolTipLst.append('Error Logic:'+'\n'+'DirectParent should not be "IS_ULTIMATELY_CONSOLIDATED_BY".') 
                  except:
                      pass
                  try:
                      if DirRelType.strip()=="IS_INTERNATIONAL_BRANCH_OF":
                          if EntCat.strip()!="Branch":
                             errMsg.append(['DirectRelationType',DirRelType,'If DirectRelationshipType is "IS_INTERNATIONAL_BRANCH_OF" then EntityCategory should be "Branch".'])
                             toolTipLst.append('Error Logic:'+'\n'+'If DirectRelationshipType is "IS_INTERNATIONAL_BRANCH_OF" then EntityCategory should be "Branch".')  
                  except:
                      pass
                  
#           ---------------------------O--------------------------------------------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(2,0).value)==True:
              try: 
                  if o=="" or o==None :
                     errMsg.append(['O',o,'O should not be blank.'])
                     toolTipLst.append('Error Logic:'+'\n'+'O should not be blank.') 
              except:
                  pass
              
              
#           -----------------------------UltimateRelationshipType---------------------------updated 4/7/2018  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(80,0).value)==True:
                if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A")  :
                  try:
                      if UlRlshpType.upper().strip()=="IS_DIRECTLY_CONSOLIDATED_BY":
                         errMsg.append(['UltimateRelationshipType',UlRlshpType,"UltimateRelationshipType should not be 'IS_DIRECTLY_CONSOLIDATED_BY'."])
                         toolTipLst.append('Error Logic:'+'\n'+"UltimateRelationshipType should not be 'IS_DIRECTLY_CONSOLIDATED_BY'.") 
                  except:
                      pass  
            
#            -------------------------------PNILegalFormationAddressLine1-----------------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(67,0).value)==True:
              try:
                 if DirPrnt!=None and PNILglFrmnAddL1!=None:
                     if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                          if PNILglFrmnAddL1!="":
                             errMsg.append(['PNILegalFormationAddressLine1',PNILglFrmnAddL1,'If DirectParent = N/A then PNILegalFormationAddressLine1 should be blank.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNILegalFormationAddressLine1 should be blank.')
                     if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                          if PNILglFrmnAddL1=="":
                             errMsg.append(['PNILegalFormationAddressLine1',PNILglFrmnAddL1,"If DirectParent has anything other than 'N/A' then PNILegalFormationAddressLine1 should not be blank."])
                             toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNILegalFormationAddressLine1 should not be blank.")
              except:
                  pass 
            '''  
           
#            -------------------------------PNI2LegalFormationAddressLine1--------------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(91,0).value)==True:
              try:
                 if UlPrnt!=None and PNI2LglFrmnAddL1!=None:  
                     if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                          if PNI2LglFrmnAddL1!="":
                             errMsg.append(['PNI2LegalFormationAddressLine1',PNI2LglFrmnAddL1,'If UltimateParent = N/A then PNI2LegalFormationAddressLine1 should be blank.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then PNI2LegalFormationAddressLine1 should be blank.')
                     if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                          if PNI2LglFrmnAddL1=="":
                             errMsg.append(['PNI2LegalFormationAddressLine1',PNI2LglFrmnAddL1,"If UltimateParent has anything other than 'N/A' then PNI2LegalFormationAddressLine1 should not be blank."])
                             toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2LegalFormationAddressLine1 should not be blank.")
              except:
                  pass      
            '''  
              
#           -------------------------------------HeadquartersAddressRegion------------new add 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(50,0).value)==True: 
                try:
                    if HQAddReg=="":
                        errMsg.append(['HeadquartersAddressRegion',HQAddReg,"HeadquartersAddressRegion should not be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+"HeadquartersAddressRegion should not be blank.")
                except:
                    pass
              
#           ------------------------------------PNILegalFormationAddressRegion-------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(69,0).value)==True: 
                try:
                 if  DirPrnt!=None and PNILglFormnAddRgn!=None:  
                     if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                          if PNILglFormnAddRgn!="":
                             errMsg.append(['PNILegalFormationAddressRegion',PNILglFormnAddRgn,'If DirectParent = N/A then PNILegalFormationAddressRegion should be blank.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNILegalFormationAddressRegion should be blank.')
                     if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                          if PNILglFormnAddRgn=="":
                             errMsg.append(['PNILegalFormationAddressRegion',PNILglFormnAddRgn,"If DirectParent has anything other than 'N/A' then PNILegalFormationAddressRegion should not be blank."])
                             toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNILegalFormationAddressRegion should not be blank.")
                except:
                  pass 
            '''
#           ------------------------------------PNI2LegalFormationAddressRegion-------------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(93,0).value)==True: 
                try:
                   if UlPrnt!=None and PNI2LglFrmnAddRgn!=None: 
                     if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                          if PNI2LglFrmnAddRgn!="":
                             errMsg.append(['PNI2LegalFormationAddressRegion',PNI2LglFrmnAddRgn,'If UltimateParent = N/A then PNI2LegalFormationAddressRegion should be blank.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then PNI2LegalFormationAddressRegion should be blank.')
                     if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                          if PNI2LglFrmnAddRgn=="":
                             errMsg.append(['PNI2LegalFormationAddressRegion',PNI2LglFrmnAddRgn,"If UltimateParent has anything other than 'N/A' then PNI2LegalFormationAddressRegion should not be blank."])
                             toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2LegalFormationAddressRegion should not be blank.")
                except:
                  pass 
            '''  
#           ----------------------------------PNILegalFormationAddressCountry----------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(70,0).value)==True: 
                try:
                  if DirPrnt!=None and PNILglFormnAddCntry!=None:
                     if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                          if PNILglFormnAddCntry!="":
                             errMsg.append(['PNILegalFormationAddressCountry',PNILglFormnAddCntry,'If DirectParent = N/A then PNILegalFormationAddressCountry should be blank.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNILegalFormationAddressCountry should be blank.')
                     if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                          if PNILglFormnAddCntry=="":
                             errMsg.append(['PNILegalFormationAddressCountry',PNILglFormnAddCntry,"If DirectParent has anything other than 'N/A' then PNILegalFormationAddressCountry should not be blank."])
                             toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNILegalFormationAddressCountry should not be blank.")
                except:
                  pass   
            '''
#           ----------------------------------PNI2LegalFormationAddressCountry-----------------updated 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(94,0).value)==True: 
                try:
                    if UlPrnt!=None and PNI2LglFrmnAddCntry!=None:
                         if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                              if PNI2LglFrmnAddCntry!="":
                                 errMsg.append(['PNI2LegalFormationAddressCountry',PNI2LglFrmnAddCntry,'If UltimateParent = N/A then PNI2LegalFormationAddressCountry should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then PNI2LegalFormationAddressCountry should be blank.')
                         if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNI2LglFrmnAddCntry=="":
                                 errMsg.append(['PNI2LegalFormationAddressCountry',PNI2LglFrmnAddCntry,"If UltimateParent has anything other than 'N/A' then PNI2LegalFormationAddressCountry should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2LegalFormationAddressCountry should not be blank.")
                except:
                  pass             
            '''
#            -------------------------------PNILegalFormationAddressPostCode--------------------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(71,0).value)==True: 
                try:
                    if DirPrnt!=None and PNILglFormnAddPstCode!=None:
                         if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                              if PNILglFormnAddPstCode!="":
                                 errMsg.append(['PNILegalFormationAddressPostCode',PNILglFormnAddPstCode,'If DirectParent = N/A then PNILegalFormationAddressPostCode should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNILegalFormationAddressPostCode should be blank.')
                         if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNILglFormnAddPstCode=="":
                                 errMsg.append(['PNILegalFormationAddressPostCode',PNILglFormnAddPstCode,"If DirectParent has anything other than 'N/A' then PNILegalFormationAddressPostCode should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNILegalFormationAddressPostCode should not be blank.")
                except:
                  pass 
            '''  
              
#            -------------------------------PNI2LegalFormationAddressPostCode---------------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(95,0).value)==True: 
                try:
                    if UlPrnt!=None and PNI2LglFrmnAddPstCode!=None:
                         if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                              if PNI2LglFrmnAddPstCode!="":
                                 errMsg.append(['PNI2LegalFormationAddressPostCode',PNI2LglFrmnAddPstCode,'If UltimateParent = N/A then PNI2LegalFormationAddressPostCode should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then PNI2LegalFormationAddressPostCode should be blank.')
                         if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNI2LglFrmnAddPstCode=="":
                                 errMsg.append(['PNI2LegalFormationAddressPostCode',PNI2LglFrmnAddPstCode,"If UltimateParent has anything other than 'N/A' then PNI2LegalFormationAddressPostCode should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2LegalFormationAddressPostCode should not be blank.")
                except:
                  pass               
            '''  
#           --------------------------------PNIHeadquartersAddressLine1------------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(72,0).value)==True: 
                try:
                    if DirPrnt!=None and PNIHQaddL1!=None: 
                         if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                              if PNIHQaddL1!="":
                                 errMsg.append(['PNIHeadquartersAddressLine1',PNIHQaddL1,'If DirectParent = N/A then PNIHeadquartersAddressLine1 should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNIHeadquartersAddressLine1 should be blank.')
                         if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNIHQaddL1=="":
                                 errMsg.append(['PNIHeadquartersAddressLine1',PNIHQaddL1,"If DirectParent has anything other than 'N/A' then PNIHeadquartersAddressLine1 should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNIHeadquartersAddressLine1 should not be blank.")
                except:
                  pass 
            '''  
#           --------------------------------PNI2HeadquartersAddressLine1-------------------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(96,0).value)==True: 
                try:
                    if UlPrnt!=None and PNI2HQAddL1!=None:
                         if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                              if PNI2HQAddL1!="":
                                 errMsg.append(['PNI2HeadquartersAddressLine1',PNI2HQAddL1,'If UltimateParent = N/A then PNI2HeadquartersAddressLine1 should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then PNI2HeadquartersAddressLine1 should be blank.')
                         if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNI2HQAddL1=="":
                                 errMsg.append(['PNI2HeadquartersAddressLine1',PNI2HQAddL1,"If UltimateParent has anything other than 'N/A' then PNI2HeadquartersAddressLine1 should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2HeadquartersAddressLine1 should not be blank.")
                except:
                  pass               
            '''  
#           -----------------------------PNIHeadquartersAddressRegion------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(74,0).value)==True: 
                try:
                    if DirPrnt!=None and PNIHQaddReg!=None:
                         if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                              if PNIHQaddReg!="":
                                 errMsg.append(['PNIHeadquartersAddressRegion',PNIHQaddReg,'If DirectParent = N/A then PNIHeadquartersAddressRegion should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNIHeadquartersAddressRegion should be blank.')
                         if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNIHQaddReg=="":
                                 errMsg.append(['PNIHeadquartersAddressRegion',PNIHQaddReg,"If DirectParent has anything other than 'N/A' then PNIHeadquartersAddressRegion should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNIHeadquartersAddressRegion should not be blank.")
                except:
                  pass  
            '''  
#           -----------------------------PNI2HeadquartersAddressRegion--------------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(98,0).value)==True: 
                try:
                    if UlPrnt!=None and PNI2HQAddRgn!=None: 
                         if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                              if PNI2HQAddRgn!="":
                                 errMsg.append(['PNI2HeadquartersAddressRegion',PNI2HQAddRgn,'If UltimateParent is N/A then PNI2HeadquartersAddressRegion should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is N/A then PNI2HeadquartersAddressRegion should be blank.')
                         if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNI2HQAddRgn=="":
                                 errMsg.append(['PNI2HeadquartersAddressRegion',PNI2HQAddRgn,"If UltimateParent has anything other than 'N/A' then PNI2HeadquartersAddressRegion should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2HeadquartersAddressRegion should not be blank.")
                except:
                  pass               
            '''  
              
#           -------------------------------PNIHeadquartersAddressPostCode-----------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(75,0).value)==True: 
                try:
                    if DirPrnt!=None and PNIHQaddPstCode!=None:
                         if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                              if PNIHQaddPstCode!="":
                                 errMsg.append(['PNIHeadquartersAddressPostCode',PNIHQaddPstCode,'If DirectParent = N/A then PNIHeadquartersAddressPostCode should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNIHeadquartersAddressPostCode should be blank.')
                         if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNIHQaddPstCode=="":
                                 errMsg.append(['PNIHeadquartersAddressPostCode',PNIHQaddPstCode,"If DirectParent has anything other than 'N/A' then PNIHeadquartersAddressPostCode should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNIHeadquartersAddressPostCode should not be blank.")
                except:
                  pass
            '''  
              
#           -------------------------------PNI2HeadquartersAddressPostCode----------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(99,0).value)==True: 
                try:
                    if UlPrnt!=None and PNI2HQAddPstCd!=None:
                         if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                              if PNI2HQAddPstCd!="":
                                 errMsg.append(['PNI2HeadquartersAddressPostCode',PNI2HQAddPstCd,'If UltimateParent = N/A then PNI2HeadquartersAddressPostCode should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then PNI2HeadquartersAddressPostCode should be blank.')
                         if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNI2HQAddPstCd=="":
                                 errMsg.append(['PNI2HeadquartersAddressPostCode',PNI2HQAddPstCd,"If UltimateParent has anything other than 'N/A' then PNI2HeadquartersAddressPostCode should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2HeadquartersAddressPostCode should not be blank.")
                except:
                  pass               
            '''  
#           -----------------------------PNIHeadquartersAddressCountry------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(76,0).value)==True: 
                try:
                    if DirPrnt!=None and PNIHQaddCntry!=None:
                         if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                              if PNIHQaddCntry!="":
                                 errMsg.append(['PNIHeadquartersAddressCountry',PNIHQaddCntry,'If DirectParent = N/A then PNIHeadquartersAddressCountry should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNIHeadquartersAddressCountry should be blank.')
                         if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNIHQaddCntry=="":
                                 errMsg.append(['PNIHeadquartersAddressCountry',PNIHQaddCntry,"If DirectParent has anything other than 'N/A' then PNIHeadquartersAddressCountry should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNIHeadquartersAddressCountry should not be blank.")
                except:
                  pass 
            '''  
              
#           -----------------------------PNI2HeadquartersAddressCountry-----------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(100,0).value)==True: 
                try:
                    if UlPrnt!=None and PNI2HQAddCntry!=None:
                         if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                              if PNI2HQAddCntry!="":
                                 errMsg.append(['PNI2HeadquartersAddressCountry',PNI2HQAddCntry,'If UltimateParent is N/A then PNI2HeadquartersAddressCountry should be blank.'])
                                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then PNI2HeadquartersAddressCountry should be blank.')
                         if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                              if PNI2HQAddCntry=="":
                                 errMsg.append(['PNI2HeadquartersAddressCountry',PNI2HQAddCntry,"If UltimateParent has anything other than 'N/A' then PNI2HeadquartersAddressCountry should not be blank."])
                                 toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2HeadquartersAddressCountry should not be blank.")
                except:
                  pass 
            '''  
#            
##           ------------------------------PNIRegistrationAuthorityID------------------------------------------
#            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(65,0).value)==True: 
#                try:
#                     '''   
#                     if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
#                          if PNIRegAuthID!="":
#                             errMsg.append(['PNIRegistrationAuthorityID',PNIRegAuthID,'If DirectParent = N/A then PNIRegistrationAuthorityID should be blank.'])
#                             toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNIHeadquartersAddressCountry should be blank.')
#                     ''' 
#                     if DirPrnt!=None and PNIRegAuthID!=None:
#                         if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
#                              if PNIRegAuthID=="":
#                                 errMsg.append(['PNIRegistrationAuthorityID',PNIRegAuthID,"If DirectParent has anything other than 'N/A' then PNIRegistrationAuthorityID should not be blank."])
#                                 toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNIRegistrationAuthorityID should not be blank.")
#                except:
#                  pass
              
                      
##           ------------------------------PNI2RegistrationAuthorityID------------------------------------------
#            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(101,0).value)==True: 
#                try:
#                    if UlPrnt!=None and  PNI2RegAuthID!=None:
#                         '''   
#                         if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
#                              if PNI2RegAuthID!="":
#                                 errMsg.append(['PNI2RegistrationAuthorityID',PNI2RegAuthID,'If UltimateParent is N/A then PNI2RegistrationAuthorityID should be blank.'])
#                                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is N/A then PNI2HeadquartersAddressCountry should be blank.')
#                         '''
#                         if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
#                              if PNI2RegAuthID=="":
#                                 errMsg.append(['PNI2RegistrationAuthorityID',PNI2RegAuthID,"If UltimateParent has anything other than 'N/A' then PNI2RegistrationAuthorityID should not be blank."])
#                                 toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2RegistrationAuthorityID should not be blank.")
#                except:
#                  pass              
            
##           ------------------------------PNIBusinessRegisterEntityID------------------------------------------
#            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(56,0).value)==True: 
#                try:
#                    if DirPrnt!=None and PNIbnsRegEnID!=None:
#                     if DirPrnt!="" and str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
#                          if PNIbnsRegEnID!="":
#                             errMsg.append(['PNIBusinessRegisterEntityID',PNIbnsRegEnID,'If DirectParent = N/A then PNIBusinessRegisterEntityID should be blank.'])
#                             toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNIBusinessRegisterEntityID should be blank.')
#                     if DirPrnt!="" and not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
#                          if PNIbnsRegEnID=="":
#                             errMsg.append(['PNIBusinessRegisterEntityID',PNIbnsRegEnID,"If DirectParent has anything other than 'N/A' then PNIBusinessRegisterEntityID should not be blank."])
#                             toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then PNIBusinessRegisterEntityID should not be blank.")
#                except:
#                  pass
              
              
##           ------------------------------PNI2BusinessRegisterEntityID------------------------------------------
#            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(102,0).value)==True: 
#                try:
#                    if UlPrnt!=None and PNI2bsnsRegEnID!=None:
#                         if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
#                              if PNI2bsnsRegEnID!="":
#                                 errMsg.append(['PNI2BusinessRegisterEntityID',PNI2bsnsRegEnID,'If UltimateParent is N/A then PNI2BusinessRegisterEntityID should be blank.'])
#                                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is N/A then PNI2BusinessRegisterEntityID should be blank.')
#                         if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
#                              if PNI2bsnsRegEnID=="":
#                                 errMsg.append(['PNI2BusinessRegisterEntityID',PNI2bsnsRegEnID,"If UltimateParent has anything other than 'N/A' then PNI2BusinessRegisterEntityID should not be blank."])
#                                 toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then PNI2BusinessRegisterEntityID should not be blank.")
#                except:
#                  pass              
              
              
#           ---------------------------------------LEI Event -----------updated-26th Jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(7,0).value)==True:
              try: 
                  
#                   logic 1
                    if  leiEvnt.lower().strip()=="dissolved":
                        if not(EntEvnt.lower().strip()=="dissolved" and leiStatus.lower().strip()=="inactive" and EntStatus.lower().strip()=="inactive"):
                            errMsg.append(['LEI Event',leiEvnt,'LEI Event has Dissolved then'+'\n'+'1.Entity Event should be Dissolved'+'\n'+'2. LEI Status should be Inactive'+'\n'+'3.Entity Status should be Inactive.'])
                            toolTipLst.append('Error Logic:'+'\n'+'LEI Event has Dissolved then'+'\n'+'1.Entity Event should be Dissolved'+'\n'+'2. LEI Status should be Inactive'+'\n'+'3.Entity Status should be Inactive.')
                  
#                   logic 2
                    if leiEvnt.strip()=="Merger/Acquisition":
                        if not(EntEvnt.strip()==" Merger/Acquisition" and leiStatus.lower().strip()=="merged" and EntStatus.lower().strip()=="inactive"):
                           errMsg.append(['LEI Event',leiEvnt,'LEI Event has Merger/Acquisition then'+'\n'+'1.Entity Event should be Merger/Acquisition'+'\n'+'2. LEI Status should be Merged'+'\n'+'3.Entity Status should be Inactive.'])
                           toolTipLst.append('Error Logic:'+'\n'+'LEI Event has Dissolved then'+'\n'+'1.Entity Event should be Merger/Acquisition'+'\n'+'2. LEI Status should be Merged'+'\n'+'3.Entity Status should be Inactive.') 
                  
                    '''                  
                    if str(leiEvnt).upper() in evtLst:
                       if not((str(EntEvnt).upper() in evtLst) and (str(EntStatus).upper()=='INACTIVE') and (ExpiryDate!='') and (leiStatus.upper() in leiStatusLst)):
                           errMsg.append(['LEI Event',leiEvnt,'LEI Event has Dissolved or Merger/acquisition then'+'\n'+'1.Entity Event should be Dissolved or Merger/acquisition'+'\n'+'2. EntityStatus should be Inactive & ExpiryDate should not be blank'+'\n'+'3.LEI Status should be Merged/Retired'])
                           toolTipLst.append('Error Logic:'+'\n'+'if LEI event has dissolved or merger/acquisition then'+'\n'+'1.Entity Event should be Dissolved or Merger/acquisition'+'\n'+'2. EntityStatus should be Inactive & ExpiryDate should not be blank'+'\n'+'3.LEI Status should be Merged/Retired')
                    '''        
              except:
                  pass
                
#           -----------------------------Entity Status-----------Added on 19th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(8,0).value)==True:
              try: 
                    if str(EntStatus).upper() =='INACTIVE':
                        if ExpiryDate=="":
                           errMsg.append(['Entity Status',EntStatus,'If Entity Status is Inactive, then Expiry Date should not be blank.'])
                           toolTipLst.append('Error Logic:'+'\n'+'If Entity Status is Inactive, then Expiry Date should not be blank.') 
                        ''' 
                        if not((str(leiEvnt).upper() in evtLst) and (str(EntEvnt).upper() in evtLst) and (ExpiryDate!='') and (leiStatus.upper() in leiStatusLst)):
                           errMsg.append(['Entity Status',EntStatus,'If Entity Status is Inactive then'+'\n'+'1.LEI Event field should be Dissolved or Merger/acquisition'+'\n'+'2.Entity Event filed should be Dissolved or Merger/acquisition'+'\n'+'3.ExpiryDate should not be blank'+'\n'+'4.LEI Status should be Merged/Retired']);toolTipLst.append('Error Logic:'+'\n'+'if Entity Status is Inactive then'+'\n'+'1.LEI Event field should be Dissolved or Merger/acquisition'+'\n'+'2.Entity Event filed should be Dissolved or Merger/acquisition'+'\n'+'3.ExpiryDate should not be blank'+'\n'+'4.LEI Status should be Merged/Retired')
                        '''           
              except:
                  pass
                
#           -------------------------------------Entity Event -----------updated 26th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(9,0).value)==True:
              try:
                  
                    if EntEvnt.lower().strip()=="dissolved":
                        if not(leiEvnt.lower().strip()=="dissolved" and leiStatus.lower().strip()=="inactive" and EntStatus.lower().strip()=="inactive"):
                            errMsg.append(['Entity Event',EntEvnt,'If Entity Event is Dissolved then'+'\n'+'1.LEI Event field should be Dissolved'+'\n'+'2. LEI Status should be Inactive'+'\n'+'3. Entity Status should be "Inactive"'])
                            toolTipLst.append('Error Logic:'+'\n'+'If Entity Event is Dissolved then'+'\n'+'If Entity Event is Dissolved then'+'\n'+'1.LEI Event field should be Dissolved'+'\n'+'2. LEI Status should be Inactive'+'\n'+'3. Entity Status should be "Inactive"')
                   
                    if EntEvnt.strip()=="Merger/Acquisition":
                        if not(leiEvnt.strip()=="Merger/Acquisition" and leiStatus.lower().strip()=="merged" and EntStatus.lower().strip()=="inactive"):
                            errMsg.append(['Entity Event',EntEvnt,'If Entity Event is Merger/Acquisition then'+'\n'+'1.LEI Event field should be Merger/Acquisition'+'\n'+'2. LEI Status should be Merged'+'\n'+'3. Entity Status should be "Inactive"'])
                            toolTipLst.append('Error Logic:'+'\n'+'If Entity Event is Merger/Acquisition then'+'\n'+'1.LEI Event field should be Merger/Acquisition'+'\n'+'2. LEI Status should be Merged'+'\n'+'3. Entity Status should be "Inactive"')
                    
                    '''
                    if str(EntEvnt).upper() in evtLst:
                       if not((str(leiEvnt).upper() in evtLst) and (str(EntStatus).upper()=='INACTIVE') and (ExpiryDate!='') and (leiStatus in leiStatusLst)):
                           errMsg.append(['Entity Event',EntEvnt,'Entity Event has Dissolved or Merger/acquisition']);toolTipLst.append('Error Logic:'+'\n'+'If Entity event has dissolved or merger/acquisition then'+'\n'+'1.LEI Event should be Dissolved or Merger/acquisition'+'\n'+'2. EntityStatus should be Inactive & ExpiryDate should not be blank'+'\n'+'3. LEI Status should be Merged/Retired')
                    '''       
              except:
                  pass  

#           ------------------------------------LEI-----------Upated:25-06-2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(12,0).value)==True:
               try:    
                   if lei=='':
                        if not(str(leiEvnt).upper().strip()=='ON HOLD' or str(leiEvnt).upper().strip()=='AWAITING VALIDATION' or str(leiEvnt).upper().strip()=='DUPLICATE' ):
                            errMsg.append(['LEI',lei,'LEI Event="'+ leiEvnt +'" should be On Hold or Awaiting Validation or Duplicate if LEI is blank.']);toolTipLst.append('Error Logic:'+'\n'+'If LEI is blank then LEI event should be On Hold/Awaiting Validation/Duplicate.')
                   else:
                       if (str(leiEvnt).upper().strip()=='ON HOLD' or str(leiEvnt).upper().strip()=='AWAITING VALIDATION' or str(leiEvnt).upper().strip()=='DUPLICATE' ):
                            errMsg.append(['LEI',lei,'LEI Event should not be On Hold or Awaiting Validation or Duplicate if LEI is not blank.']);toolTipLst.append('Error Logic:'+'\n'+'LEI Event should not be On Hold/Awaiting Validation/Duplicate if LEI is not blank.')
               except:
                   pass
           
            
#           --------------------------------Anglicised Entity Name----------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(14,0).value)==True:
              try: 
                  if angEnName!=None :
                    if isEnglish(angEnName)==False:
                        errMsg.append(['Anglicised Entity Name',angEnName,'AnglicisedName has non English character']);toolTipLst.append('Error Logic:'+'\n'+'If AnglicisedName has any Non English Character.')
                    if angEnName!='' and isEnglish(officialEnNm)==True:
                        errMsg.append(['Anglicised Entity Name',angEnName,'For any value of AnglicisedName, OfficialEntityName "'+officialEnNm +'" should have atleast one non-english character.']);toolTipLst.append('Error Logic:'+'\n'+'If AnglicisedEntityName has any value then+'+'\n'+'OfficialEntityName should have atleast one non english or special character.')     
              except:
                  pass
          
#           ---------------------------------Entity Legal Form--------------updated 26th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(118,0).value)==True:
              try:
                  if entLglForm.strip()=="8888 - Other Please Specify":
                      if LegalFormTxt=="":
                          errMsg.append(['Entity Legal Form',entLglForm,"If '8888 - Other Please Specify' then the field 'Specify Other' should not be blank."])
                          toolTipLst.append('Error Logic:'+'\n'+"If '8888 - Other Please Specify' then the field 'Specify Other' should not be blank.")
                  '''
                  if entLglForm!=None:
                    if entLglForm=='':errMsg.append(['Entity Legal Form',entLglForm,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'Entity Legal Form should not be blank.')
                    else:
                        abbr,abbrFF,langCheck=look4_LFlogic(officialEnNm,cntry, logicSht)
    #                   L1            
                        if (str(entLglForm).upper()=='PLEASE SPECIFY OTHER' or str(entLglForm).upper()=='OTHER - PLEASE SPECIFY') and LegalFormTxt=='' : errMsg.append(['Entity Legal Form',entLglForm,'LegalFormFreeText should not be blank if Entity Legal Form is: OTHER - please specify']);toolTipLst.append('Error Logic:'+'\n'+' If Entity Legal Form is "Please specify other" then'+'\n'+'LegalFormFreeText should not be blank')     
    
    #                   L2                    
                        if str(entLglForm).upper() in LglLst and "pension fund" not in str(entLglForm).lower() and fundMg=='' :                        
                            errMsg.append(['Entity Legal Form',entLglForm,'If Entity Legal Form is "' + entLglForm \
                            +'" then fund manager should not be blank.'])
                            toolTipLst.append('Error Logic:'+'\n'+'If Entity Legal Form has specified values then fund \
                            manager should not be blank.')
                        if str(entLglForm).upper() in LglLst and "pension fund" not in str(entLglForm).lower() and str(EntCat).lower()!="fund":
                            errMsg.append(['Entity Legal Form',entLglForm,'If Entity Legal Form is "' + entLglForm \
                            +'" then Entity Category should be "Fund"'])
                            toolTipLst.append('Error Logic:'+'\n'+'If Entity Legal Form has specified values then  \
                            Entity Category should be "Fund".')
                  '''        
              except:
                  pass
                        
                        
#           -----------------------------LegalFormFreeText-----------------------updated 6/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(17,0).value)==True:
               
              try:
#                  print elfSht.nrows
                  for i in range (0,elfSht.nrows) :
#                      print cntry.upper().strip()
#                      print (elfSht.cell(i,0).value).upper().strip()
#                      wait=input("ww")
                      if cntry.upper().strip() in (elfSht.cell(i,0).value).upper().strip(): 
#                          print "en1"
#                          wait=input("11")
                          if LegalFormTxt.strip()== elfSht.cell(i,1).value.strip():
#                             print "1" 
#                             print LegalFormTxt.strip()
#                             print "2"
#                             print elfSht.cell(i,1).value
#                             wait=input("dd")
                             errMsg.append(['Specify Other',LegalFormTxt,'Specify Other contains already existing Entity legal Form for the CountryLegalForm.'])
                             toolTipLst.append('Error Logic:'+'\n'+'Specify Other contains already existing Entity legal Form for the CountryLegalForm.')
                             break
              except:
                  pass
              '''  
              try:  
                if LegalFormTxt!=None:  
                    if LegalFormTxt!='' :
                        if look4_LegalForms(LegalFormTxt,LegalFrmSht)!=None:
                            errMsg.append(['LegalFormTxt',LegalFormTxt,'LegalFormText should not have a Legal Form, already available as a drop down option in database'])
                            toolTipLst.append('Error Logic:'+'\n'+'LegalFormFreeText should not have data from '+'\n'+' "Existing Legal Forms-UnaVista" sheet')
                        if 'FUND' in uLegalFormTxt and 'SUB' in uLegalFormTxt:
                            specCnt=len(re.findall(r' ',uLegalFormTxt))
                            if specCnt>0:errMsg.append(['LegalFormTxt',LegalFormTxt,'Sub-Fund is not in correct format']);toolTipLst.append('Error Logic:'+'\n'+'Incorrect format, correct format being Sub-Fund (without spaces)')
              except:
                  pass
              '''    
#           -----------------------------HeadquartersAddressLine1---------------UPDATED 16/7/2018
            
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(19,0).value)==True:
                try:
                    if HQAddressLine1=="":
                       errMsg.append(['HeadquartersAddressLine1',HQAddressLine1,'HeadquartersAddressLine1 should not be blank.'])
                       toolTipLst.append('Error Logic:'+'\n'+'HeadquartersAddressLine1 should not be blank.') 
                except:
                    pass
            '''    
                lglFrmLst=["fund", "fondo", "fonds", "fondos", "inversion", "ucits", "icvc", "sicav"]
                lglFrmTxtLst=["fund","unit trust"]
                try:
                 if any(word in str(entLglForm).lower() for word in lglFrmLst) \
                    or any(wrd in str(LegalFormTxt).lower() for wrd in lglFrmTxtLst):
                    if str(HQAddressLine1)[:3].lower()!='c/o': errMsg.append(['HeadquartersAddressLine1',HQAddressLine1,'HeadquartersAddressLine1 must begin with "C/O"']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressCountry should not be blank.HeadquartersAddressLine1 must begin with "C/O".')
                except:
                    pass
            '''     
#           ---------------------HeadQuartersAddressTown/City---------------------------
            cntExList=['LUXEMBOURG','PANAMA','GIBRALTAR']
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(21,0).value)==True:

              try:  
                if  HQCity!=None: 
                    if not str(HQCity).upper() in cntExList:
                        for row_num in range(cntrySht.nrows):
                            row_value = cntrySht.row_values(row_num)                     
                            if (row_value[0]).lower() == HQCity.lower():
                                errMsg.append(['HeadQuartersAddressTown/City',HQCity,'HeadQuartersAddressTown/City should not have a country name in it.'])
                                toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressTown/City should not have a country name in it.') 
                                break
#                    if HQCity=='':errMsg.append(['HeadQuartersAddressTown/City',HQCity,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressTown/City should not be blank.')
#                    else:
                    if re.findall(r'\d+',HQCity)!=[]:errMsg.append(['HeadQuartersAddressTown/City',HQCity,'HQ Town/city has numeric value in it']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressTown/City should not '+'\n'+'have numeric value in it')
#                    if (not str(HQCity).upper() in cntExList) and (cntry in str(HQCity).upper()):errMsg.append(['HeadQuartersAddressTown/City',HQCity,'HQ Town/city has a country name "' + cntry + '" in it']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressTown/City should not '+'\n'+'have a country name in it except from Luxembourg, Panama, Gibraltar')
              except:
                  pass
            
#           -------------------------------HeadQuartersAddressCounty/State--------------updated 27th jun 2018

            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(22,0).value)==True:
              try: 

                '''  
                if HQCounty!=None:  
                    if HQCounty!='':
                        if re.findall(r'\d+',HQCounty)!=[]:errMsg.append(['HeadQuartersAddressCounty/State',HQCounty,'HQ County/State has numeric value in it']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressCounty/State should not '+'\n'+' have numeric value in it')
                        if cntry in str(HQCounty).upper():errMsg.append(['HeadQuartersAddressCounty/State',HQCounty,'HQ County/State has a country name "' + cntry + '" in it']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressCounty/State should not '+'\n'+'have a country name in it')
#                '''
#                logic1        
                if not(HQCounty=="" and HQCity==""):
                    if HQCounty.lower()==HQCity.lower():
                         errMsg.append(['HeadQuartersAddressCounty/State',HQCounty,'HeadquartersAddressCounty/State should not match HeadquartersAddressTown/City.'])
                         toolTipLst.append('Error Logic:'+'\n'+'HeadquartersAddressCounty/State should not match HeadquartersAddressTown/City.') 
                 
#                logic2 
                if not(HQCounty=="" and HQAddReg==""):
                    if HQCounty.lower()==HQAddReg.lower() :
                         errMsg.append(['HeadQuartersAddressCounty/State',HQCounty,'HeadquartersAddressCounty/State should not match HeadquartersAddressRegion.'])
                         toolTipLst.append('Error Logic:'+'\n'+'HeadquartersAddressCounty/State should not match HeadquartersAddressRegion.') 
             
#                logic3
                if HQCounty=="":
                    errMsg.append(['HeadQuartersAddressCounty/State',HQCounty,'HeadquartersAddressCounty/State should not be blank.'])
                    toolTipLst.append('Error Logic:'+'\n'+'HeadquartersAddressCounty/State should not be blank.') 
              except:
                  pass
            
#           -----------------------------HeadQuartersAddressCountry-----------------Added on 8th Nov 2016
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(23,0).value)==True:
              try:  
                if HQCntry!=None:  
                    if HQCntry=='':errMsg.append(['HeadQuartersAddressCountry',HQCntry,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressCountry should not be blank.')
              except:
                  pass
        
#           ------------------------HeadQuartersAddressPostCode--------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(24,0).value)==True:
              try: 
                if HQPostCode!=None: 
                    if (str(HQPostCode).lower()==str(HQAddAddNo).lower()) or \
                       (str(HQPostCode).lower()==str(HQAddAddNoBldg).lower()) or \
                       (str(HQPostCode).lower()==str(HQCity).lower()) or \
                       (str(HQPostCode).lower()==str(HQCntry).lower()):
                            errMsg.append(['HeadquartersAddressPostCode',HQPostCode,'Value in HeadQuartersAddressPostCode should not be repeated in HeadquartersAddressAddressNumber,HeadquartersAddressAddressNumberWithinBuilding,HeadquartersAddressTown,City,HeadquartersAddressCounty/State.'])
                            toolTipLst.append('Error Logic:'+'\n'+'Value in HeadQuartersAddressPostCode should not be repeated in HeadquartersAddressAddressNumber,HeadquartersAddressAddressNumberWithinBuilding,HeadquartersAddressTown,City,HeadquartersAddressCounty/State.')                            
                      
                    
#                    if HQPostCode=='':errMsg.append(['HeadQuartersAddressPostCode',HQPostCode,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressPostCode should not be blank.')
#                    else:
                    if HQPostCode.upper()=='NA' or HQPostCode.upper()=='N/A':errMsg.append(['HeadQuartersAddressPostCode',HQPostCode,'PostCode is N/A']);toolTipLst.append('Error Logic:'+'\n'+'HeadQuartersAddressPostCode Should not have'+'\n'+' NA, N/A in this field')
                        
              except:
                  pass
            
#           -----------------------------LegalFormationAddressLine1---------------removed 27th jun 2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(25,0).value)==True:
              try: 
                if LFaddress1!=None:  
                    if LFaddress1=='':errMsg.append(['LegalFormationAddressLine1',LFaddress1,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressLine1 should not be blank.')
              except:
                  pass
            '''
#           ---------------------------LegalFormationAddressTown/City--------------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(27,0).value)==True:
              try: 
                if (not str(LFCity).upper() in cntExList):
                    for row_num in range(cntrySht.nrows):
                        row_value = cntrySht.row_values(row_num)                     
                        if (row_value[0]).lower() == LFCity.lower():
                            errMsg.append(['LegalFormationAddressTown/City',LFCity,'LegalFormationAddressTown/City should not have a country name in it.'])
                            toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressTown/City should not have a country name in it.') 
                            break
#                    if LFCity=='':errMsg.append(['LegalFormationAddressTown/City',LFCity,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressTown/City should not be blank.')
#                    else:
                if re.findall(r'\d+',LFCity)!=[]:
                    errMsg.append(['LegalFormationAddressTown/City',LFCity,'LegalForm Town/City has numeric value in it' ])
                    toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressTown/City should not '+'\n'+'have numeric value in it')
#                        if (not str(LFCity).upper() in cntExList) and cntry in str(LFCity).upper():errMsg.append(['LegalFormationAddressTown/City',LFCity,'LegalForm Town/City has a country name "' + cntry + '" in it']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressTown/City should not '+'\n'+'have a country name in it except from Luxembourg, Singapore, Hong Kong, Gibraltar')
              except:
                  pass
           
#           --------------------------------LegalFormationAddressCounty/State--------------updated 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(28,0).value)==True:
              try: 
                '''  
                if LFCountry!=None: 
                    
                    if LFCountry!='':
                        if re.findall(r'\d+',LFCountry)!=[]:errMsg.append(['LegalFormationAddressCounty/State',LFCountry,'LegalForm County/State has numeric value in it']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCounty/State should not'+'\n'+'have numeric value in it')
                        if cntry in LFCountry:errMsg.append(['LegalFormationAddressCounty/State',LFCountry,'LegalForm County/State has a country name "' + cntry + '" in it']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCounty/State should not'+'\n'+'have a country name in it')
                if not(LFCounty=="" and LFCity==""):
                '''        

                if LFCounty.lower()==LFCity.lower():
                         errMsg.append(['LegalFormationAddressCounty/State',LFCounty,'LegalFormationAddressCounty/State should not match LegalFormationAddressTown/City.'])
                         toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCounty/State should not match LegalFormationAddressTown/City.') 
                         
                if not(LFCounty=="" and lglFrmnAddRgn==""):
                    if LFCounty.lower()==lglFrmnAddRgn.lower() :
                         errMsg.append(['LegalFormationAddressCounty/State',LFCounty,'LegalFormationAddressCounty/State should not match LegalFormationAddressRegion.'])
                         toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCounty/State should not match LegalFormationAddressRegion.') 
                         
                if LFCounty=="":
                    errMsg.append(['LegalFormationAddressCounty/State',LFCounty,'LegalFormationAddressCounty/State should not be blank.'])
                    toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCounty/State should not be blank.') 
                         
              except:
                  pass
            
#           ---------------------------------LegalFormationAddressCountry-----------------------updated 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(29,0).value)==True:
              try: 
                if LFCountry1=="":
                    errMsg.append(['LegalFormationAddressCountry',LFCountry1,'LegalFormationAddressCountry should not be blank.'])
                    toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCountry should not be blank.') 
                  
                '''  
                if LFCountry1!=None:  
                    if LFCountry1=='':errMsg.append(['LegalFormationAddressCountry',LFCountry1,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCountry should not be blank.')
                    else:
                        if str(LFCountry1).upper()!=str(BRCountry).upper() :errMsg.append(['LegalFormationAddressCountry',LFCountry1,'LegalFormationAddressCountry should match the BusinessRegistryCountry "' + BRCountry + '".']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCountry should match '+'\n'+'with text in BusinessRegistryCountry field')
                        if str(LFCountry1).upper()!=cntry:errMsg.append(['LegalFormationAddressCountry', LFCountry1,'LegalFormationAddressCountry should match the CountryLegalForm "' + cntry + '".']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressCountry should match'+'\n'+'with the text in CountryLegalForm field')
                '''             
              except:
                   pass
               
#           ----------------------------------LegalFormationAddressRegion---------------added 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(116,0).value)==True:
                try: 
                    if lglFrmnAddRgn=="":
                        errMsg.append(['LegalFormationAddressRegion',lglFrmnAddRgn,'LegalFormationAddressRegion should not be blank.'])
                        toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressRegion should not be blank.') 
                except:
                    pass
    
#           -----------------------------------LegalFormationAddressPostCode-----------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(30,0).value)==True:
              try:
                if (str(LFPostCode).lower().strip()==str(LglFrmnAddAddNo).lower().strip()) or \
                       (str(LFPostCode).lower().strip()==str(lglFormnAddAddNoBldg).lower().strip()) or \
                       (str(LFPostCode).lower().strip()==str(LFCity).lower().strip()) or \
                       (str(LFPostCode).lower().strip()==str(LFCounty).lower().strip()):
                            errMsg.append(['LegalFormationAddressPostCode',LFPostCode,'Value in LegalFormationAddressPostCode should not be repeated in LegalFormationAddressAddressNumber,LegalFormationAddressAddressNumberWithinBuilding,LegalFormationAddressTown/City,LegalFormationAddressCounty/State.'])
                            toolTipLst.append('Error Logic:'+'\n'+'Value in LegalFormationAddressPostCode should not be repeated in LegalFormationAddressAddressNumber,LegalFormationAddressAddressNumberWithinBuilding,LegalFormationAddressTown/City,LegalFormationAddressCounty/State.')  
              except:
                  pass
              try:
                if LFPostCode!=None:  
                    if LFPostCode=='':errMsg.append(['LegalFormationAddressPostCode',LFPostCode,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressPostCode should not be blank.')
                    else:
                        if LFPostCode=='NA' or LFPostCode=='N/A':errMsg.append(['LegalFormationAddressPostCode',LFPostCode,'PostCode should not have N/A']);toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressPostCode Should not have'+'\n'+'NA, N/A in this field')
                             
              except:
                  pass
        
#           ----------------------------------OfficialBusinessRegistryFreeText-------------------updated 2/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(33,0).value)==True:
              '''  
#               L1 
              try:
                if OBReg!=None:  
                    if OBReg.upper().lstrip().rstrip()=="RA888888 - OTHER - PLEASE SPECIFY  - THEN -":
                      if OBRegTxt=="":
                       errMsg.append(['OfficialBusinessRegistryFreeText',OBRegTxt,'If RegistrationAuthorityID is "RA888888 - OTHER - PLEASE SPECIFY" then OfficialBusinessRegistryFreeText can not be blank.'])
                       toolTipLst.append('Error Logic:'+'\n'+'If RegistrationAuthorityID is "RA888888 - OTHER - PLEASE SPECIFY" then OfficialBusinessRegistryFreeText can not be blank.') 
              except:
                  pass
#              old logics- L1 
              try: 
                if OBRegTxt!=None:  
                    if OBRegTxt!='' and look4_BRegistry(OBRegTxt,businessSht)!=None:                                                            
                        errMsg.append(['OfficialBusinessRegistryFreeText',OBRegTxt,'OBRfreeText should not have business Registry, already available as a drop down option in database'])
                        toolTipLst.append('Error Logic:'+'\n'+'OBRfreeText should not have Business Registry from already available '+'\n'+'as a drop down option under Business Registry in database')   
              except:
                  pass
              
#              old logics -L2
              try: 
                if OBRegTxt!=None:  
                    if OBRegTxt!='' and offBusRegRef=="":                                                            
                        errMsg.append(['OfficialBusinessRegistryFreeText',OBRegTxt," If there is anything in the 'OfficialBusinessRegistryFreeText' field, then 'OfficialBusinessRegistryReference' should not be blank"])
                        toolTipLst.append('Error Logic:'+'\n'+" If there is anything in the 'OfficialBusinessRegistryFreeText' field, then 'OfficialBusinessRegistryReference' should not be blank")   
              except:
                  pass
              '''
#              logic1
              try:
                if "document" in str(OBRegTxt).lower(): 
                    if offBusRegRef.lower()!="document":
                        errMsg.append(['OfficialBusinessRegistryFreeText',OBRegTxt,'OfficialBusinessRegistryFreeText should not have the word DOCUMENT.'])
                        toolTipLst.append('Error Logic:'+'\n'+'If OfficialBusinessRegistryFreeText should not have the word DOCUMENT.')
              except:
                  pass
              
#              logic2
              try:
                  
                  if OBRegTxt.lower().strip()==offBusRegRef.lower().strip():
                      if not(OBRegTxt=="" and offBusRegRef==""):
                          errMsg.append(['OfficialBusinessRegistryFreeText',OBRegTxt,'OfficialBusinessRegistryFreeText should not be equal to OfficialBusinessRegistryReference.'])
                          toolTipLst.append('Error Logic:'+'\n'+'OfficialBusinessRegistryFreeText should not be equal to OfficialBusinessRegistryReference.') 
              except:
                  pass
          
#           -----------------------OfficialBusinessRegistryReference--------------------updated 2/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(119,0).value)==True:
              '''      
#              old logics-L1
              try:
                  if offBusRegRef!="":
                      if OBReg=="":
                           errMsg.append(['OfficialBusinessRegistryReference',str(offBusRegRef),"If there is anything in the OfficialBusinessRegistryReference' field then 'RegistrationAuthorityID' should not be blank "])
                           toolTipLst.append('Error Logic:'+'\n'+"If there is anything in the OfficialBusinessRegistryReference' field then 'RegistrationAuthorityID' should not be blank ") 
              except:
                  pass
              
#              old logics-L2
              try:
                  if offBusRegRef!="" and OBReg=="OTHER - PLEASE SPECIFY":
                      if OBRegTxt=="":
                           errMsg.append(['OfficialBusinessRegistryReference',str(offBusRegRef)," If there is anything in the OfficialBusinessRegistryReference' and OfficialBusinessRegistry has 'OTHER - PLEASE SPECIFY', then 'OfficialBusinessRegistryFreeText' should not be blank. "])
                           toolTipLst.append('Error Logic:'+'\n'+" If there is anything in the OfficialBusinessRegistryReference' and OfficialBusinessRegistry has 'OTHER - PLEASE SPECIFY', then 'OfficialBusinessRegistryFreeText' should not be blank.") 
              except:
                  pass
              '''
              try:
                 if offBusRegRef!=None:
    #              new logics L1
                   if offBusRegRef=='' or str(offBusRegRef).lower().lstrip().rstrip().replace(" ","")=='n/a' or str(offBusRegRef).lower().lstrip().rstrip().replace(" ","")=='na':
                       errMsg.append(['OfficialBusinessRegistryReference',str(offBusRegRef),'OfficialBusinessRegistryReference can not be blank or Whitespaces, empty strings and variations of N/A'])
                       toolTipLst.append('Error Logic:'+'\n'+'OfficialBusinessRegistryReference can not be blank or Whitespaces, empty strings and variations of N/A.') 
              except:
                  pass
              
              '''
              try:
                  if OBReg!=None:
#                      new logics-L2               
                      if OBReg.replace(" ","")=="RA999999-RegistryNA(Trust,SubFund)-UploadSupportingDocumentation" and str(offBusRegRef).lower()!="document":
                          errMsg.append(['OfficialBusinessRegistryReference',str(offBusRegRef),'If RegistrationAuthorityID is "RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation" then OfficialBusinessRegistryReference must be Document.'])
                          toolTipLst.append('Error Logic:'+'\n'+'If RegistrationAuthorityID is "RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation" then OfficialBusinessRegistryReference must be Document.') 
              except:
                  pass
              
#              old logics-L3
              try:    
               OBRegRefLst=['SC', 'OC', 'SL', 'LP', 'SO', 'NI', 'ZC', 'SP' ]    
               if str(cntry).lower()=="united kingdom" and str(OBReg).upper().replace(" ","")=="RA000585-COMPANIESHOUSE":
                 if not(offBusRegRef.lstrip().rstrip()[:1]=="1" or offBusRegRef.lstrip().rstrip()[:1]=="0" or \
                          any(wrd in str(offBusRegRef)[:2].upper() for wrd in OBRegRefLst)):                     
                          errMsg.append(['OfficialBusinessRegistryReference',str(offBusRegRef),'If country is United Kingdom and OfficialBusinessRegistry is Companies House then - OfficialBusinessRegistryReference should start with 1, SC, OC, 0 (zero), SL, LP, SO, NI, ZC, SP.'])
                          toolTipLst.append('Error Logic:'+'\n'+'If country is United Kingdom and OfficialBusinessRegistry is Companies House then - OfficialBusinessRegistryReference should start with 1, SC, OC, 0 (zero), SL, LP, SO, NI, ZC, SP.') 
               if str(cntry).upper()=='AUSTRALIA':
                 if re.findall(r'^\d{3} \d{3} \d+$',str(offBusRegRef))==[]:errMsg.append(['OfficialBusinessRegistryRef',offBusRegRef,'For Australia, Registry number should be in the format like 123 456 789']);toolTipLst.append('Error Logic:'+'\n'+'If country is Australia/France then business registry number should have'+'\n'+'9 digits only in the format:123 456 789 (123 space 456 space 789)')
               if str(cntry).upper()=='SWITZERLAND':   
                  if str(offBusRegRef).strip()[:3]!="CHE": 
                    errMsg.append(['OfficialBusinessRegistryRef',str(offBusRegRef),'For Switzerland, OfficialBusinessRegistryReference should start with CHE']);toolTipLst.append('Error Logic:'+'\n'+'If country is Switzerland then OfficialBusinessRegistryReference should start with CHE') 
               if str(cntry).upper()=='CYPRUS':
                   if re.findall(r'^HE \d+$',str(offBusRegRef))==[]:errMsg.append(['OfficialBusinessRegistryReference',offBusRegRef,'For Cyprus, OfficialBusinessRegistryReference should be in the format like HE 169425']);toolTipLst.append('Error Logic:'+'\n'+'If country is Cyprus then OfficialBusinessRegistryReference should start with'+'\n'+'HE and should follow the format: HE 169425')
               if str(cntry).upper()=='GERMANY':
                   if re.findall(r'^HRA \d+$|^HRB \d+$',str(offBusRegRef.upper()))==[]:errMsg.append(['OfficialBusinessRegistryReference',offBusRegRef,'For Germany, OfficialBusinessRegistryReference should be in the format like HRA 45617 or HRB 77473']);toolTipLst.append('Error Logic:'+'\n'+'If country is Germany then OfficialBusinessRegistryReference should'+'\n'+'start with HRA/HRB & follow the format: HRA 45617 or HRB 77473')
               if  str(cntry).upper()=='FRANCE':
                   if re.findall(r'^\d{3} \d{3} \d+$',str(offBusRegRef))==[]:errMsg.append(['OfficialBusinessRegistryReference',offBusRegRef,'For France, OfficialBusinessRegistryReference should be in the format like 123 456 789']);toolTipLst.append('Error Logic:'+'\n'+'If country is Australia/France then OfficialBusinessRegistryReference should have'+'\n'+'9 digits only in the format:123 456 789 (123 space 456 space 789')
               if  str(cntry).upper()=='ROMANIA':
                   if re.findall(r'^J',str(offBusRegRef))==[]:errMsg.append(['OfficialBusinessRegistryReference',offBusRegRef,'For Romania,OfficialBusinessRegistryReference should begin with "J"']);toolTipLst.append('Error Logic:'+'\n'+'If country is Australia/France then OfficialBusinessRegistryReference should begin with "J"')
              except:
                  pass
              '''
#           ----------------------------------BIC AND FRN------------------removed 2/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(35,0).value)==True:  
              try:  
                if bic!='':
                    if str(bic).split(' ',1)[0].upper()!=str(officialEnNm).split(' ',1)[0].upper():errMsg.append(['BIC',bic,'First word "'+bic.split(' ',1)[0]+'" should match with the first word of OfficialEntityName "' +officialEnNm.split(' ',1)[0] + '"']);toolTipLst.append('Error Logic:'+'\n'+'BIC first word should match with'+'\n'+'the first word of OfficialEntityName')
              except:
                  pass
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(36,0).value)==True:  
             try:  
                if FRN!='':
                    if str(FRN).split(' ',1)[0].upper()!=str(officialEnNm).split(' ',1)[0].upper():errMsg.append(['FRN',FRN,'First word "'+FRN.split(' ',1)[0]+'" should match with the first word of OfficialEntityName "' +officialEnNm.split(' ',1)[0] + '"']);toolTipLst.append('Error Logic:'+'\n'+'FRN first word should match with'+'\n'+'the first word of OfficialEntityName')
             except:
                 pass
            '''
            
#           ---------------------------------AssociatedLEI--------------------removed 2/7/2018
            '''
           if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(41,0).value)==True:
              try:  
                if EnableCrawl==1:
                    if AssociatedLEI=='' and fundMg!='':
                        LeiVal=ValidateLEI(fundMg)
                        if LeiVal!='' and LeiVal!=None:errMsg.append(['AssociatedLEI',AssociatedLEI,'LEI "' +LeiVal+ '" has been found for fund manager "'+ fundMg +'"']);toolTipLst.append('Error Logic:'+'\n'+'If LEI is blank and fund manager has value,'+'\n'+' then validate LEI for this fund manager from website: '+'\n'+'https://www.gleif.org/en/lei/search')
              except:
                  pass
           '''
#           ----------------------------------ISIN-----------------------------removed 2/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(37,0).value)==True:
              try:  
                if EnableCrawl==1:
                    if ISIN=='':
                        webIsin=ValidateISIN(officialEnNm)
                        if webIsin!='' and webIsin!=None:errMsg.append(['ISIN',ISIN,'ISIN "' +webIsin+ '" has been found for official entity "'+ officialEnNm +'"']);toolTipLst.append('Error Logic:'+'\n'+'If ISIN is found check web for ISIN (http://www.isincodes.net/)'+'\n'+'If found on web then highlight')
              except:
                  pass
            '''
#           --------------------------------Linked SMF Issuer--------------------removed 2/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(38,0).value)==True:  
              try:
                  if LinkedIssuer!=None:
                    if LinkedIssuer!='':
                        if str(LinkedIssuer).split(' ',1)[0].upper()!=str(officialEnNm).split(' ',1)[0].upper():errMsg.append(['Linked SMF Issuer',LinkedIssuer,'First word "'+LinkedIssuer.split(' ',1)[0]+'" should match the first word of OfficialEntityName "' +officialEnNm.split(' ',1)[0] + '"']);toolTipLst.append('Error Logic:'+'\n'+'Linked SMF Issuer first word should match with'+'\n'+'the first word of OfficialEntityName')
              except:
                  pass
            '''
#           ----------------------------Expiry Date--------------------updated 2/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(43,0).value)==True: 
              try: 
#                  logic 1
                  if ExpiryDate!="":
                      if not(EntEvnt.upper().strip()=="DISSOLVED" or EntEvnt.upper().strip()=="DUPLICATE" or EntEvnt.upper().strip()=="MERGER/ACQUISITION"):
                          errMsg.append(['Expiry Date',ExpiryDate,"If expiry date is not blank, then Entity Event should be DISSOLVED or DUPLICATE or MERGER/ACQUISITION."])
                          toolTipLst.append('Error Logic:'+'\n'+"If expiry date is not blank, then Entity Event should be DISSOLVED or DUPLICATE or MERGER/ACQUISITION.") 
                  '''
                  if ExpiryDate >= infoLUDt:
                     errMsg.append(['Expiry Date',ExpiryDate,"InformationLastUpdated must be ahead of Expiry Date."])
                     toolTipLst.append('Error Logic:'+'\n'+"InformationLastUpdated must be ahead of Expiry Date.") 
                  if leiStatus.lower()=="issued":
                      if ExpiryDate !="":
                          errMsg.append(['Expiry Date',ExpiryDate,"'LEI Status' is 'ISSUED' then Expiry Date field should be blank."])
                          toolTipLst.append('Error Logic:'+'\n'+"'LEI Status' is 'ISSUED' then Expiry Date field should be blank.")
                 
                  
                  if ExpiryDate !='':
#                      isDate=isinstance(ExpiryDate, datetime.date) or isinstance(ExpiryDate, datetime.datetime)
#                      print isDate
#                      print ExpiryDate
                      dateObj=None
                        
                      if type(ExpiryDate)==unicode or type(ExpiryDate)==str:
                          dateObj=pandas.to_datetime(ExpiryDate)
                      else:
                          dateObj= xlrd.xldate.xldate_as_datetime(ExpiryDate, wb.datemode)
                      exp_dt=dateObj.date()
                      
                      isDate=isinstance(exp_dt, datetime.date)
                      if isDate:
                          print leiEvnt
                          print str(leiEvnt).upper() in evtLst
                          print str(EntEvnt).upper() in evtLst
                          print str(EntStatus).upper()=='INACTIVE'
                          print leiStatus in leiStatusLst
                          print not((str(leiEvnt).upper() in evtLst) and (str(EntEvnt).upper() in evtLst) and (str(EntStatus).upper()=='INACTIVE') and (leiStatus in leiStatusLst))
                         
                          if not((str(leiEvnt).upper() in evtLst) and (str(EntEvnt).upper() in evtLst) and (str(EntStatus).upper()=='INACTIVE') and (leiStatus in leiStatusLst)):
                              errMsg.append(['Expiry Date',ExpiryDate,'If Expiry Date has any date in it then'+'\n'+'1.LEI Event filed should be Dissolved or Merger/acquisition'+'\n'+'2.Entity Event filed should be Dissolved or Merger/acquisition'+'\n'+'3.Entity Status should be Inactive'+'\n'+'4.LEI Status should be Merged/Retired'])
                              toolTipLst.append('Error Logic:'+'\n'+'if Expiry Date has any date in it then'+'\n'+'1.LEI Event filed should be Dissolved or Merger/acquisition'+'\n'+'2.Entity Event filed should be Dissolved or Merger/acquisition'+'\n'+'3.Entity Status should be Inactive'+'\n'+'4.LEI Status should be Merged/Retired')
                  '''           
              except:
                  pass
              try:
                  if EntStatus.lower().strip()=="inactive":
                      if ExpiryDate=="":
                         errMsg.append(['Expiry Date',ExpiryDate,"Expiry Date should not be blank if Entity Status is Inactive."])
                         toolTipLst.append('Error Logic:'+'\n'+"Expiry Date should not be blank if Entity Status is Inactive.") 
              except:
                  pass
           
#           --------------------------------Annual Renewal Date---------------------updated 5/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(44,0).value)==True:
              '''  
              if leiStatus!=None and leiEvnt!=None and str(leiStatus).upper()=='ISSUED' and str(leiEvnt).upper()=="VALIDATED": 
#              new logic 
             
              try:  
                    if ARDate > firstAss:                        
                        arDate=int(ARDate)
                        arDate=datetime.datetime.fromordinal(datetime.datetime(1900,1,1).toordinal()+arDate-2)
                        errMsg.append(['Annual Renewal Date',(str(arDate).split(" "))[0],"Annual renewal date can not be greater than the date existing in 'First Assigned'."])
                        toolTipLst.append('Error Logic:'+'\n'+"Annual renewal date can not be greater than the date existing in 'First Assigned'.")
              except:
                  pass
              '''   
              
#              logic 1
              try:
                  if leiStatus.upper().strip()=="ISSUED" and leiEvnt.upper().strip()=="VALIDATED":
                      dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=ARDate-2)
                      annRenewDate =dtRead.date()
                      dtRead=datetime.datetime(1900, 1, 1) + datetime.timedelta(days=PaySet_Dt-2)
                      dtNewPaymentSetUp =datetime.date(dtRead.year+1,dtRead.month, dtRead.day) + datetime.timedelta(days=60)
                      paySet=dtRead.date()
                      if not(paySet<=annRenewDate<=dtNewPaymentSetUp):
                           errMsg.append(['Annual Renewal Date',str(annRenewDate),"If 'LEI Status' is 'ISSUED' and 'LEI Event' is 'Validated', Annual renewal Date >=  Payment Setup Date & Annual renewal date <= 1 Year + 60 days from Payment Setup Date."])
                           toolTipLst.append('Error Logic:'+'\n'+"If 'LEI Status' is 'ISSUED' and 'LEI Event' is 'Validated', Annual renewal Date >=  Payment Setup Date & Annual renewal date <= 1 Year + 60 days from Payment Setup Date.")
#                      else:
#                          print "yes"
#                      wait=input("11")    
##                     logic 2      
#                      if annRenewDate> dtNewPaymentSetUp:
#                          errMsg.append(['Annual Renewal Date',str(annRenewDate),"If 'LEI Status' is 'ISSUED' and 'LEI Event' is 'Validated',Annual renewal date <= 1 Year + 60 days from Payment Setup Date."])
#                          toolTipLst.append('Error Logic:'+'\n'+"If 'LEI Status' is 'ISSUED' and 'LEI Event' is 'Validated', Annual renewal date <= 1 Year + 60 days from Payment Setup Date.")
              except:
                  pass
              '''
#              L2
              try:
                     print  leiStatus.upper()
                     print leiEvnt.upper()
                     
#                    if ARDate!='':
                     if leiStatus.upper()=="ISSUED" and leiEvnt.upper()=="VALIDATED":
                        
                        dateObj=None
                        dateObjLstUp=None
                        if type(ARDate)==unicode or type(ARDate)==str:
                          dateObj=pandas.to_datetime(ARDate)
                          dateObjLstUp=pandas.to_datetime(infoLUDt)
                          
                        else:
                          dateObj= xlrd.xldate.xldate_as_datetime(ARDate, wb.datemode)
                          dateObjLstUp=xlrd.xldate.xldate_as_datetime(infoLUDt, wb.datemode)
                      
                        ar_dt=dateObj.date()
#                        ar_day=dateObj.weekday()
                        lstUpdt=dateObjLstUp.date()
                        
                        ftrDate2=lstUpdt+timedelta(days=1)
#                        ar_dt=datetime.date(dateObj)                          
    #                    dateTday=datetime.now().date()
#                        lstUpdt=datetime.date(dateObjLstUp)
    #                    ftrDate=datetime.date(dateTday.year + 1, dateTday.month, dateTday.day)
                        ftrDate=datetime.date(lstUpdt.year + 1, lstUpdt.month, lstUpdt.day)
                        ftrDate1=ftrDate+timedelta(days=60)
#                        print ar_dt
#                        print ftrDate1
#                        print ftrDate2
#                        print ar_dt<=ftrDate1
#                        print ftrDate2<ar_dt
#                        wait=input("hh")
                        if not(ftrDate2<ar_dt<=ftrDate1):
                            errMsg.append(['Annual Renewal Date',str(ar_dt),"If 'LEI Status' is 'ISSUED' and 'LEI Event' is 'Validated', Annual renewal Date should be at least 1 Business day ahead & Annual renewal date <= InformationLastUpdated date + 1 Year + 60 days."])
                            toolTipLst.append('Error Logic:'+'\n'+"If 'LEI Status' is 'ISSUED' and 'LEI Event' is 'Validated', Annual renewal Date should be at least 1 Business day ahead & Annual renewal date <= InformationLastUpdated date + 1 Year + 60 days.")
                 
#    #              L1
#                       
#                        if ar_dt>ftrDate:
#                            errMsg.append(['Annual Renewal Date',str(ar_dt),"Renewal date should be prior to Today's date+1 Year +60 days."])
#                            toolTipLst.append('Error Logic:'+'\n'+"Renewal date should be prior to Today's date+1 Year +60 days.")
                               
              except:
                  print traceback.print_exc()
#                  wait=input("error hu mai")
                  pass
             '''

#           -----------------------------Validation Sources------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(45,0).value)==True:                               
#              LegalFrmlst= ['pension fund', 'pension scheme', 'trust', 'charity', 'partnership']   #used in validation sources

              try:
#                 new logic
                 if Sources.upper().strip()=="ENTITY_SUPPLIED_ONLY":
                     if OBReg.strip()!="RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation":
                         errMsg.append(['Validation Sources',Sources,"If Validation Sources is ENTITY_SUPPLIED_ONLY', then RegistrationAuthorityID' is RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation."])
                         toolTipLst.append('Error Logic:'+'\n'+"If Validation Sources is ENTITY_SUPPLIED_ONLY', then RegistrationAuthorityID' is RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation.")
              except:
                  pass
              try:
#                     logic 2
                  if Sources.upper().strip()=="ENTITY_SUPPLIED_ONLY":
                     if OBRegTxt.strip()!="":
                        errMsg.append(['Validation Sources',Sources,"If Validation Sources is ENTITY_SUPPLIED_ONLY', then OfficialBusinessRegistryFreeText should be blank."]) 
                        toolTipLst.append('Error Logic:'+'\n'+"If Validation Sources is ENTITY_SUPPLIED_ONLY', then OfficialBusinessRegistryFreeText should be blank.")
              except:
                  pass
              '''
#              L1
              try:
#                if any(wrd in LegalForm.lower() for wrd in LegalFrmlst): 
                  if str(entLglForm).lower().lstrip().rstrip().replace(" ","")=='pensionfund' or\
                     str(entLglForm).lower().lstrip().rstrip().replace(" ","")=='pensionscheme' or\
                     str(entLglForm).lower().lstrip().rstrip().replace(" ","")=='trust' or\
                     str(entLglForm).lower().lstrip().rstrip().replace(" ","")=='partnership':
                         if Sources.upper()=="FULLY_CORROBORATED":
                            errMsg.append(['Validation Sources',Sources," If 'LegalForm' is Pension Fund, Pension Scheme, Trust, Charity, Partnership then 'Validation Sources' should not have 'FULLY_CORROBORATED'"])
                            toolTipLst.append('Error Logic:'+'\n'+" If 'LegalForm' is Pension Fund, Pension Scheme, Trust, Charity, Partnership then 'Validation Sources' should not have 'FULLY_CORROBORATED'")
                  
#                      if  not(str(leiEvnt).upper()=="AWAITING VALIDATION" or str(leiEvnt).upper()=="HOLD") :
#                            if Sources=="": 
#                                errMsg.append(['Validation Sources',Sources,"If LEI Event is not Awaiting Validation or Hold then this filed should not be blank"])
#                                toolTipLst.append('Error Logic:'+'\n'+"If LEI Event is not Awaiting Validation or Hold then this filed should not be blank")                                                                              
              except:
                  pass
#              L2
              try:
                  if Sources.upper()=="FULLY_CORROBORATED":
                     if OBRegTxt=="" or offBusRegRef=="":
                        errMsg.append(['Validation Sources',Sources," If Validation Sources is FULLY_CORROBORATED then OfficialBusinessRegistryFreeText and OfficialBusinessRegistryReference should not be blank"])
                        toolTipLst.append('Error Logic:'+'\n'+"If Validation Sources is FULLY_CORROBORATED then OfficialBusinessRegistryFreeText and OfficialBusinessRegistryReference should not be blank") 
              except:
                  pass
#               L3
              try:    
                leiEvntLst=["awaiting validation","hold"]
                if not any(wrd in leiEvnt.lower() for wrd in leiEvntLst):
                    if Sources=='':
                        errMsg.append(['Validation Sources',Sources,"If LEI Event is not 'Awaiting Validation' or Hold so field should not be blank"]);toolTipLst.append('Error Logic:'+'\n'+'If LEI Event is not Awaiting Validation or Hold then this filed should not be blank')
              except:
                  pass
              '''
              
#           ----------------------DirectParent----------- 
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(51,0).value)==True:

    #              l1
                  try:
                    if DirExcptnRsn!="" and DirExcptnRsn!=None:      
                      if str(DirPrnt).lstrip().rstrip().lower()!="n/a":                   
                          errMsg.append(['Direct Parent',DirPrnt,'Should be N/A if DirectExceptionReason is not blank.'])
                          toolTipLst.append('Error Logic:'+'\n'+'Should be N/A if DirectExceptionReason is not blank.')
                  except:
                      pass
                  
    #              l2
                  try:
                      if officialEnNm.lower()==DirPrnt.lower():
                          errMsg.append(['Direct Parent',DirPrnt,'Should not match with data in OfficialEntityName.'])
                          toolTipLst.append('Error Logic:'+'\n'+'Should not match with data in OfficialEntityName.') 
                  except:
                      pass
                  
    #              l3
                  try:
                      if DirPrnt.strip().upper()=="N/A":
                          if UlPrnt.strip().upper()!="N/A":
                              errMsg.append(['Direct Parent',DirPrnt,'If DirectParent is "N/A" then UltimateParent should also be "N/A".'])
                              toolTipLst.append('Error Logic:'+'\n'+'If DirectParent is "N/A" then UltimateParent should also be "N/A".')  
                  except:
                      pass
                  
    #              L4
                  try:
                      if DirPrnt=="":
                          errMsg.append(['Direct Parent',DirPrnt,'DirectParent should not be blank.'])
                          toolTipLst.append('Error Logic:'+'\n'+'DirectParent should not be blank.')  
                  except:
                      pass
#           -----------------------------PNILegalFormationAddressCity-----------------------removed 4/7/2018
            '''  
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(68,0).value)==True:
             try:
                 if DirPrnt!=None and PNILglFormnAddCity!=None: 
                     if DirPrnt!="" and str(DirPrnt).lower()=="n/a" or str(DirPrnt).lower()=="na":
                         if PNILglFormnAddCity!="":
                            errMsg.append(['PNILegalFormationAddressCity',PNILglFormnAddCity,'If DirectParent = N/A then PNILegalFormationAddressCity should be blank.'])
                            toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNILegalFormationAddressCity should be blank.') 
             except:
                 pass
                
#             L1
             try:
                 if DirPrnt!=None: 
                     if PNILglFormnAddCity!=None:
                         if DirPrnt!="" and str(DirPrnt).lower()!="n/a" :
                           if PNILglFormnAddCity=='':
                              errMsg.append(['PNILegalFormationAddressCity',PNILglFormnAddCity,'If DirectParent is anything except N/A then PNILegalFormationAddressCity should not be blank.'])
                              toolTipLst.append('Error Logic:'+'\n'+'If DirectParent is anything except N/A then PNILegalFormationAddressCity should not be blank.')
                         
                         if DirPrnt!="" and str(DirPrnt).lower()!="n/a" and PNILglFormnAddCity!="" and PNILglFormnAddCntry!="":
                             
                              if (str(PNILglFormnAddCity).upper() not in cntExList):
                                  if str(PNILglFormnAddCntry).lower()in str(PNILglFormnAddCity).lower():
                                      errMsg.append(['PNILegalFormationAddressCity',PNILglFormnAddCity,'If DirectParent is anything except N/A then PNILegalFormationAddressCity field should not be PNILegalFormationAddressCountry except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" '])                               
                                      toolTipLst.append('Error Logic:'+'\n'+'If DirectParent is anything except N/A then PNILegalFormationAddressCity field should not have a country name in it except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" ') 
             except:
                   pass
             try: 
                 if DirPrnt!=None and PNILglFormnAddCity!=None:
                      if DirPrnt!="" and str(DirPrnt).lower()!="n/a":
                          if PNILglFormnAddCity!='':                   
        #              L2               
                            if re.findall(r'\d+',PNILglFormnAddCity)!=[]:
                             errMsg.append(['PNILegalFormationAddressCity',PNILglFormnAddCity,'If DirectParent is anything except N/A then PNILegalFormationAddressCity should not have numeric value in it.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If DirectParent is anything except N/A then PNILegalFormationAddressCity should not have numeric value in it.')     
             except:
                   pass
               
            '''   
#           ------------------------------------PNI2LegalFormationAddressCity----------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(92,0).value)==True:
             try:
                 if UlPrnt!=None and PNI2LglFrmnAddCity!=None:
                     if UlPrnt!="" and str(UlPrnt).lower()=="n/a" or str(UlPrnt).lower()=="na":
                         if PNI2LglFrmnAddCity!="":
                            errMsg.append(['PNI2LegalFormationAddressCity',PNI2LglFrmnAddCity,'If UltimateParent = N/A then PNI2LegalFormationAddressCity should be blank.'])
                            toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then PNI2LegalFormationAddressCity should be blank.') 
             except:
                 pass
                
#             L1
             try:
                 if UlPrnt!=None and PNI2LglFrmnAddCity!=None:
                     if UlPrnt!="" and str(UlPrnt).lower()!="n/a" :
                        if PNI2LglFrmnAddCity=='':
                          errMsg.append(['PNI2LegalFormationAddressCity',PNI2LglFrmnAddCity,'If UltimateParent is anything except N/A then PNI2LegalFormationAddressCity should not be blank.'])
                          toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2LegalFormationAddressCity should not be blank.')
                 if UlPrnt!=None and PNI2LglFrmnAddCity!=None and PNI2LglFrmnAddCntry!=None:
                     if UlPrnt!="" and str(UlPrnt).lower()!="n/a" and PNI2LglFrmnAddCity!="" and PNI2LglFrmnAddCntry!="":
                         
                          if (str(PNI2LglFrmnAddCity).upper() not in cntExList):
                              if str(PNI2LglFrmnAddCntry).lower()in str(PNI2LglFrmnAddCity).lower():
                                  errMsg.append(['PNI2LegalFormationAddressCity',PNI2LglFrmnAddCity,'If UltimateParent is anything except N/A then PNI2LegalFormationAddressCity field should not be PNI2LegalFormationAddressCountry except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" '])                               
                                  toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2LegalFormationAddressCity field should not have a country name in it except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" ') 
             except:
                   pass
             try:   
                 if UlPrnt!=None and PNI2LglFrmnAddCity!=None:
                      if UlPrnt!="" and str(UlPrnt).lower()!="n/a" and PNI2LglFrmnAddCity!='':                   
        #              L2                       
                        if re.findall(r'\d+',PNI2LglFrmnAddCity)!=[]:
                             errMsg.append(['PNI2LegalFormationAddressCity',PNI2LglFrmnAddCity,'If UltimateParent is anything except N/A then PNI2LegalFormationAddressCity should not have numeric value in it.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2LegalFormationAddressCity should not have numeric value in it.')     
             except:
                   pass   
            '''
#           -----------------------PNIHeadquartersAddressCity----------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(73,0).value)==True:
#              logic 1
             try: 
                  if DirPrnt!=None and PNIHQaddCity!=None:
                      if DirPrnt!="" and str(DirPrnt).lower()=="n/a" or str(DirPrnt).lower()=="na":
                          
                          if PNIHQaddCity!="":
                            errMsg.append(['PNIHeadquartersAddressCity',PNIHQaddCity,'If DirectParent = N/A then PNIHeadquartersAddressCity should be blank.'])
                            toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then PNIHeadquartersAddressCity should be blank.')
             except:
                 pass
             
             try:
                 if DirPrnt!=None and PNIHQaddCity!=None:
                      if DirPrnt!="" and not(str(DirPrnt).lower()=="n/a" or str(DirPrnt).lower()=="na") :
                         if PNIHQaddCity=='':                 
                          errMsg.append(['PNIHeadquartersAddressCity',PNIHQaddCity,'If DirectParent is anything except N/A then PNIHeadquartersAddressCity should not be blank.'])
                          toolTipLst.append('Error Logic:'+'\n'+'If DirectParent is anything except N/A then PNIHeadquartersAddressCity should not be blank.')
             except:
                 pass
             
             try:
                 if DirPrnt!=None and PNIHQaddCity!=None and PNIHQaddCntry!=None:
                  if DirPrnt!="" and str(DirPrnt).lower()!="n/a" and PNIHQaddCity!="" and PNIHQaddCntry!="":
                      if (str(PNIHQaddCity).upper() not in cntExList):
                          if str(PNIHQaddCntry).lower()in str(PNIHQaddCity).lower():
                              errMsg.append(['PNIHeadquartersAddressCity',PNIHQaddCity,'If DirectParent is anything except N/A then PNIHeadquartersAddressCity field should not be PNILegalFormationAddressCountry except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" '])                               
                              toolTipLst.append('Error Logic:'+'\n'+'If DirectParent is anything except N/A then PNIHeadquartersAddressCity field should not be PNIHeadquartersAddressCountry except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" ') 
             except:
                 pass 
             
             try: 
                 if DirPrnt!=None and PNIHQaddCity!=None:
                      if DirPrnt!="" and str(DirPrnt).lower()!="n/a" and PNIHQaddCity!='':                   
            #              L2
                           if re.findall(r'\d+',PNIHQaddCity)!=[]:
                             errMsg.append(['PNIHeadquartersAddressCity',PNIHQaddCity,'If DirectParent is anything except N/A then PNIHeadquartersAddressCity should not have numeric value in it.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If DirectParent is anything except N/A then PNIHeadquartersAddressCity should not have numeric value in it.')  
             except:
                 pass
             
            ''' 
#           -----------------------PNI2HeadquartersAddressCity------------------removed 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(97,0).value)==True:
#              logic 1
             try: 
                if UlPrnt!=None and PNI2HQAddCity!=None:  
                  if UlPrnt!="" and str(UlPrnt).lower()=="n/a" or str(UlPrnt).lower()=="na":
                      
                      if PNI2HQAddCity!="":
                        errMsg.append(['PNI2HeadquartersAddressCity',PNI2HQAddCity,'If UltimateParent is N/A then PNI2HeadquartersAddressCity should be blank.'])
                        toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is N/A then PNI2HeadquartersAddressCity should be blank.')
             except:
                 pass
             
             try:
                 if UlPrnt!=None and PNI2HQAddCity!=None:
                      if UlPrnt!="" and not(str(UlPrnt).lower()=="n/a" or str(UlPrnt).lower()=="na") :
                         if PNI2HQAddCity=='':                 
                          errMsg.append(['PNI2HeadquartersAddressCity',PNI2HQAddCity,'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not be blank.'])
                          toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not be blank.')
             except:
                 pass
             
             try:
                 if UlPrnt!=None and PNI2HQAddCity!=None and PNI2HQAddCntry!=None:
                  if UlPrnt!="" and str(UlPrnt).lower()!="n/a" :
                      if PNI2HQAddCity!="" :
                          if PNI2HQAddCntry!="":
                              if (str(PNI2HQAddCity).upper() not in cntExList):
                                  if str(PNI2HQAddCntry).lower()in str(PNI2HQAddCity).lower():
                                      errMsg.append(['PNI2HeadquartersAddressCity',PNI2HQAddCity,'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity field should not be PNI2LegalFormationAddressCountry except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" '])                               
                                      toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity field should not be PNI2HeadquartersAddressCountry except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" ') 
             except:
                 pass 
             
             try: 
                 if UlPrnt!=None and PNI2HQAddCity!=None:
                   if UlPrnt!="" and str(UlPrnt).lower()!="n/a" and PNI2HQAddCity!='':                   
        #              L2
                       if re.findall(r'\d+',PNI2HQAddCity)!=[]:
                         errMsg.append(['PNI2HeadquartersAddressCity',PNI2HQAddCity,'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not have numeric value in it.'])
                         toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not have numeric value in it.')  
             except:
                 pass             
            '''  
#           ------------------------------DirectRelationshipStatus-----------------updated 3/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(58,0).value)==True:
                if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A"):
#              logic1                
                     try:
                          if DirPrnt.upper().strip()!="N/A" or DirPrnt.upper().strip()!="NA":
                              if DirRelshpStatus=="":
                                  errMsg.append(['DirectRelationshipStatus',DirRelshpStatus,"DirectRelationshipStatus can not be blank if DirectParent is not N/A."])
                                  toolTipLst.append('Error Logic:'+'\n'+"DirectRelationshipStatus can not be blank if DirectParent is not N/A.")
                     except:
                          pass
                     try:
                           if DirPrnt.upper().strip()=="N/A" or DirPrnt.upper().strip()=="NA":
                             if DirRelshpStatus!="":
                               errMsg.append(['DirectRelationshipStatus',DirRelshpStatus,"DirectRelationshipStatus should be blank if DirectParent is N/A."])
                               toolTipLst.append('Error Logic:'+'\n'+"DirectRelationshipStatus should be blank if DirectParent is N/A.") 
                     except:
                          pass
                     
        #             logic2
                     try:
                          if leiStatus.strip().lower()!="merged" and leiStatus.strip().lower()!="inactive":
                              if DirRelshpStatus.strip().lower()=="inactive":
                                  errMsg.append(['DirectRelationshipStatus',DirRelshpStatus,"If LEI Status is not Merged or Inactive then DirectRelationshipStatus can not be Inactive."])
                                  toolTipLst.append('Error Logic:'+'\n'+"If LEI Status is not Merged or Inactive then DirectRelationshipStatus can not be Inactive.")
                     except:
                          pass
                  

#           ----------------------------DirectAccountingPeriodStart--------------updated 3/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(54,0).value)==True:
              if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A"):  
                try:
                    '''  
                    if DirVldnDocs.strip()=="ACCOUNTS_FILING":                        
                        if DirAccPrdStart=="":
                             errMsg.append(['DirectAccountingPeriodStart',str(DirAccPrdStart),'If DirectValidationDocuments = ACCOUNTS_FILING, then DirectAccountingPeriodStart should not be blank.'])                  
                             toolTipLst.append('Error Logic:'+'\n'+'If DirectValidationDocuments = ACCOUNTS_FILING, then DirectAccountingPeriodStart should not be blank.')
                    '''
                    if DirPrnt!="" and DirPrnt.upper()!="N/A":
                        if DirAccPrdStart=="":
                            errMsg.append(['DirectAccountingPeriodStart',str(DirAccPrdStart),'If DirectParent has anything other than N/A then DirectAccountingPeriodStart should not be blank.'])                  
                            toolTipLst.append('Error Logic:'+'\n'+'If DirectParent has anything other than N/A then DirectAccountingPeriodStart should not be blank.')
                               
                except:
                    pass
                      
#           ----------------------------DirectRelationshipPeriodStart--------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(61,0).value)==True:
              if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A") :
                    try: 
                        if DirPrnt.strip().upper()!="N/A":
                            if DirRlnshpPrdStrt=="":
                                errMsg.append(['DirectRelationshipPeriodStart',str(DirRlnshpPrdStrt).split(" ")[0],"If DirectParent does not euqal to N/A then DirectRelationshipPeriodStart should not be blank."])                 
                                toolTipLst.append('Error Logic:'+'\n'+"If DirectParent does not euqal to N/A then DirectRelationshipPeriodStart should not be blank.")
                    except:
                          pass
                  
                  
#           ------------------------------DirectFilingPeriodStart--------------updated 4/7/2018          
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(55,0).value)==True:
                if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A"):
                   try:
                       if "REGULATORY_FILING" in DirVldnDocs.upper().strip():
                           if DirFilngPrdStrt=="":
                              errMsg.append(['DirectFilingPeriodStart',DirFilngPrdStrt,"If DirectValidationDocuments has 'REGULATORY_FILING' then DirectFilingPeriodStart can not be blank."])                 
                              toolTipLst.append('Error Logic:'+'\n'+"If DirectValidationDocuments has 'REGULATORY_FILING' then DirectFilingPeriodStart can not be blank.") 
                       elif "REGULATORY_FILING" not in DirVldnDocs.upper().strip() or DirVldnDocs=="" : 
                           if DirFilngPrdStrt!="":
                              errMsg.append(['DirectFilingPeriodStart',DirFilngPrdStrt,"If DirectValidationDocuments is not 'REGULATORY_FILING' then DirectFilingPeriodStart should be blank."])                 
                              toolTipLst.append('Error Logic:'+'\n'+"If DirectValidationDocuments is not 'REGULATORY_FILING' then DirectFilingPeriodStart should be blank.") 
                   except:
                       pass
        
#           --------------------------------UltimateParent-------------------------updated 4/7/2018    
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(78,0).value)==True:
#                logic1
                try:
                   if UlExcptRsn!="":
                       if UlPrnt.strip().upper()!="N/A":
                           errMsg.append(['UltimateParent',UlPrnt,'Should be N/A if UltimateExceptionReason is not blank.'])                 
                           toolTipLst.append('Error Logic:'+'\n'+'Should be N/A if UltimateExceptionReason is not blank.') 
                except:
                    pass
                
#                logic2
                try:
                    if UlPrnt.lower().strip()==officialEnNm.lower().strip():
                        errMsg.append(['UltimateParent',UlPrnt,'Should not match with data in OfficialEntityName.'])                 
                        toolTipLst.append('Error Logic:'+'\n'+'Should not match with data in OfficialEntityName.') 
                except:
                    pass
             
#                logic3
                try:
                    if UlPrnt.upper().strip()=="N/A":
                        if DirPrnt.upper().strip()!="N/A":
                            errMsg.append(['UltimateParent',UlPrnt,'If UltimateParent is N/A then DirectParent should be N/A.'])                 
                            toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is N/A then DirectParent should be N/A.') 
                except:
                    pass
                
#                logic4
                try:
                    if UlPrnt=="":
                        errMsg.append(['UltimateParent',UlPrnt,'UltimateParent should not be blank.'])                 
                        toolTipLst.append('Error Logic:'+'\n'+'UltimateParent should not be blank.') 
                except:
                    pass
                '''       
    #            L1
                 try:
                     if UlPrnt!=None:
                         if UlPrnt=="":
                           errMsg.append(['UltimateParent',UlPrnt,'UltimateParent should not be blank.'])                 
                           toolTipLst.append('Error Logic:'+'\n'+'UltimateParent should not be blank.')    
                 except:
                     pass
                 
    #              l2
                 try:
                      mtch=[]
                      if UlPrnt!=None:
                          if str(UlPrnt).lower().lstrip().rstrip()==str(officialEnNm).lower().lstrip().rstrip() :
                              mtch.append("OfficialEntityName")
                          if str(UlPrnt).lower().lstrip().rstrip()==str(altEnName).lower().lstrip().rstrip() :
                              mtch.append("AlternateEntityName")
                          if str(UlPrnt).lower().lstrip().rstrip()==str(prevEnNam).lower().lstrip().rstrip():
                              mtch.append("Previous Entity Name")
                          if str(UlPrnt).lower().lstrip().rstrip()==str(angEnName).lower().lstrip().rstrip():
                              mtch.append("Anglicised Entity Name")
                          if len(mtch)!=0:
                              errorField=""
                              for i in range(len(mtch)):
                                  
                                  errorField=errorField+mtch[i]
                                  errorField=errorField+"\n"
                              errMsg.append(['UltimateParent',UlPrnt,'DirectParent should not match '+errorField])
                              toolTipLst.append('Error Logic:'+'\n'+'DirectParent should not match '+errorField)
                          
                 except:
                      pass
    #            L3
                 try:  
                  if UlPrnt!="" and UlPrnt!=None:  
                      if UlExcptRsn!=None:
                        if str(UlPrnt).upper()=="N/A" and UlExcptRsn=="":
                           errMsg.append(['UltimateParent',UlPrnt,'If UltimateParent is "N/A" if UltimateExceptionReason must not be blank.'])                 
                           toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is "N/A" if UltimateExceptionReason must not be blank.')  
                        elif str(UlPrnt).upper()!="N/A" and UlExcptRsn!="":
                           errMsg.append(['UltimateParent',UlPrnt,'If UltimateParent is not "N/A" if UltimateExceptionReason must be blank.'])                 
                           toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is not "N/A" if UltimateExceptionReason must beblank.') 
                 except:
                     pass
                '''
                ''' 
    #            L2
                 try:
                  if UlPrnt.lower()=="n/a":
                    notBlnkFlds=[]
                    if UlLEI!=None and UlLEI.lower()!="":notBlnkFlds.append("UltimateLEI")
                    if UlRlshpType!=None and UlRlshpType!="": notBlnkFlds.append("UltimateRelationshipType")
                    if UlRlshpStatus!=None and UlRlshpStatus!="":notBlnkFlds.append("UltimateRelationshipStatus")
                    if UlVldnSrcs!=None and UlVldnSrcs!="":notBlnkFlds.append("UltimateValidationSources")
                    if UlVldnDocs!=None and UlVldnDocs!="":notBlnkFlds.append("UltimateValidationDocuments")
                    if UlRlshpQlfrCat!=None and UlRlshpQlfrCat!="":notBlnkFlds.append("UltimateRelationshipQualifierCategory")
                    if UlAccPrdStrt!=None and UlAccPrdStrt!="":notBlnkFlds.append("UltimateAccountingPeriodStart")
                    if UlAccPrdEnd!=None and UlAccPrdEnd!="":notBlnkFlds.append("UltimateAccountingPeriodEnd")
                    if UlRlshpPrdStrt!=None and UlRlshpPrdStrt!="":notBlnkFlds.append("UltimateRelationshipPeriodStart")
                    if UlRlshpPrdEnd!=None and UlRlshpPrdEnd!="":notBlnkFlds.append("UltimateRelationshipPeriodEnd")
                    if UlFlngPrdStrt!=None and UlFlngPrdStrt!="":notBlnkFlds.append("UltimateFilingPeriodStart")
                    if UlFlngPrdEnd!=None and UlFlngPrdEnd!="":notBlnkFlds.append("UltimateFilingPeriodEnd")
                    if PNI2LglFrmnAddL1!=None and PNI2LglFrmnAddL1!="":notBlnkFlds.append("PNI2LegalFormationAddressLine1")
                    if PNI2LglFrmnAddCity!=None and PNI2LglFrmnAddCity!="":notBlnkFlds.append("PNI2LegalFormationAddressCity")
                    if PNI2LglFrmnAddRgn!=None and PNI2LglFrmnAddRgn!="":notBlnkFlds.append("PNI2LegalFormationAddressRegion")
                    if PNI2LglFrmnAddCntry!=None and PNI2LglFrmnAddCntry!="":notBlnkFlds.append("PNI2LegalFormationAddressCountry")
                    if PNI2LglFrmnAddPstCode!=None and PNI2LglFrmnAddPstCode!="":notBlnkFlds.append("PNI2LegalFormationAddressPostCode")
                    if PNI2HQAddL1!=None and PNI2HQAddL1!="":notBlnkFlds.append("PNI2HeadquartersAddressLine1")
                    if PNI2HQAddCity!=None and PNI2HQAddCity!="":notBlnkFlds.append("PNI2HeadquartersAddressCity")
                    if PNI2HQAddRgn!=None and PNI2HQAddRgn!="":notBlnkFlds.append("PNI2HeadquartersAddressRegion")
                    if PNI2HQAddPstCd!=None and PNI2HQAddPstCd!="":notBlnkFlds.append("PNI2HeadquartersAddressPostCode")
                    if PNI2HQAddCntry!=None and PNI2HQAddCntry!="":notBlnkFlds.append("PNI2HeadquartersAddressCountry")
                    if PNI2RegAuthID!=None and PNI2RegAuthID!="":notBlnkFlds.append("PNI2RegistrationAuthorityID")
                    if PNI2bsnsRegEnID!=None and PNI2bsnsRegEnID!="":notBlnkFlds.append("PNI2BusinessRegisterEntityID")
                    if notBlnkFlds!=[]:    
                        errLst=""
                        for i in range(len(notBlnkFlds)):
                            errLst=errLst+notBlnkFlds[i]
                            errLst=errLst+"\n"
                        errMsg.append(['UltimateParent',UlPrnt,'If UltimateParent is "N/A" then ' +str(errLst)+ "must be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is "N/A" then ' +str(errLst)+ "must be blank.")    
                 except:
                     pass
    #            L3
                 try:    
                  if UlPrnt.lower()!="n/a":    
                    blnkFlds=[]
                    if UlRlshpType=="": blnkFlds.append("UltimateRelationshipType")
                    if UlRlshpStatus=="":blnkFlds.append("UltimateRelationshipStatus")
                    if UlVldnSrcs=="":blnkFlds.append("UltimateValidationSources")
                    if UlVldnDocs=="":blnkFlds.append("UltimateValidationDocuments")
                    if UlRlshpQlfrCat=="":blnkFlds.append("UltimateRelationshipQualifierCategory")
                    if UlAccPrdStrt=="":blnkFlds.append("UltimateAccountingPeriodStart")
                    if UlRlshpPrdStrt=="":blnkFlds.append("UltimateRelationshipPeriodStart")
                    if UlFlngPrdStrt=="":blnkFlds.append("UltimateFilingPeriodStart")
                    if PNI2LglFrmnAddL1=="":blnkFlds.append("PNI2LegalFormationAddressLine1")
                    if PNI2LglFrmnAddCity=="":blnkFlds.append("PNI2LegalFormationAddressCity")
                    if PNI2LglFrmnAddRgn=="":blnkFlds.append("PNI2LegalFormationAddressRegion")
                    if PNI2LglFrmnAddCntry=="":blnkFlds.append("PNI2LegalFormationAddressCountry")
                    if PNI2LglFrmnAddPstCode=="":blnkFlds.append("PNI2LegalFormationAddressPostCode")
                    if PNI2HQAddL1=="":blnkFlds.append("PNI2HeadquartersAddressLine1")
                    if PNI2HQAddCity=="":blnkFlds.append("PNI2HeadquartersAddressCity")
                    if PNI2HQAddRgn=="":blnkFlds.append("PNI2HeadquartersAddressRegion")
                    if PNI2HQAddPstCd=="":blnkFlds.append("PNI2HeadquartersAddressPostCode")
                    if PNI2HQAddCntry=="":blnkFlds.append("PNI2HeadquartersAddressCountry")
                    if PNI2RegAuthID=="":blnkFlds.append("PNI2RegistrationAuthorityID")
                    if PNI2bsnsRegEnID=="":blnkFlds.append("PNI2BusinessRegisterEntityID")
                    if blnkFlds!=[]:    
                        errLst=""
                        for i in range(len(blnkFlds)):
                            errLst=errLst+blnkFlds[i]
                            errLst=errLst+"\n"
                        errMsg.append(['UltimateParent',UlPrnt,'If UltimateParent is not "N/A" then ' +str(errLst)+ "must not be blank."])
                        toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is not "N/A" then ' +str(errLst)+ "must not be blank.")    
                 except:
                     pass
    #            L4
                 try:    
                  if UlPrnt.lower()!="n/a":    
                    if UlExcptRsn!="":
                        errMsg.append(['UltimateParent',UlPrnt,'If UltimateParent has anything other than "N/A" then UltimateExceptionReason field should be blank.'])
                        toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then UltimateExceptionReason field should be blank") 
                 except:
                     pass
                ''' 
             
#           ----------------------UltimateValidationDocuments--------------------removed 4/7/2018
            ''' 
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(83,0).value)==True: 
                try:
                    if "REGULATORY_FILING" in UlVldnDocs.upper():
                        if UlVldnSrcs.upper()=="ENTITY_SUPPLIED_ONLY":
                           errMsg.append(['UltimateValidationDocuments',UlVldnDocs,'If UltimateValidationDocuments has "REGULATORY_FILING", then "UltimateValidationSources" cannot be ENTITY_SUPPLIED_ONLY.'])
                           toolTipLst.append('Error Logic:'+'\n'+"If UltimateValidationDocuments has 'REGULATORY_FILING', then 'UltimateValidationSources' cannot be ENTITY_SUPPLIED_ONLY.")       
                except:
                    pass
                try:
                    if "ACCOUNTS_FILING" in UlVldnDocs.upper():
                        if UlAccPrdStrt=="":
                           errMsg.append(['UltimateValidationDocuments',UlVldnDocs,'If UltimateValidationDocuments has "ACCOUNTS_FILING", then UltimateAccountingPeriodStart can not be blank.'])
                           toolTipLst.append('Error Logic:'+'\n'+"If UltimateValidationDocuments has 'ACCOUNTS_FILING ', then UltimateAccountingPeriodStart can not be blank.") 
                except:
                    pass
             '''       
#           ----------------------UltimateRelationshipStatus----------------updated 4/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(81,0).value)==True:
#              logic1  
               if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A") :
                  try:
                      if UlRlshpStatus=="" or UlRlshpStatus.strip().lower()=="inactive":
                          errMsg.append(['UltimateRelationshipStatus',UlRlshpStatus,"If UltimateRelationshipStatus should not be blank OR 'Inactive'."])
                          toolTipLst.append('Error Logic:'+'\n'+"If UltimateRelationshipStatus should not be blank OR 'Inactive'.") 
                  except:
                      pass
                  try:
                      if UlPrnt.upper().strip()!="N/A" or UlPrnt.upper().strip()!="NA":
                          if UlRlshpStatus=="":
                              errMsg.append(['UltimateRelationshipStatus',UlRlshpStatus,"UltimateRelationshipStatus can not be blank if UltimateParent is not N/A."])
                              toolTipLst.append('Error Logic:'+'\n'+"UltimateRelationshipStatus can not be blank if UltimateParent is not N/A.")
                  except:
                      pass
                  try:
                       if UlPrnt.upper().strip()=="N/A" or UlPrnt.upper().strip()=="NA":
                         if UlRlshpStatus!="":
                           errMsg.append(['UltimateRelationshipStatus',UlRlshpStatus,"UltimateRelationshipStatus should be blank if UltimateParent is N/A."])
                           toolTipLst.append('Error Logic:'+'\n'+"UltimateRelationshipStatus should be blank if UltimateParent is N/A.") 
                  except:
                      pass
#             
              
            
#           ------------------------UltimateAccountingPeriodStart------------updared 4/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(85,0).value)==True:
                if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A"):
                    try:
                       if UlVldnDocs.strip()=="ACCOUNTS_FILING":
                          if UlAccPrdStrt=="":
                              errMsg.append(['UltimateAccountingPeriodStart',UlAccPrdStrt,'If UltimateValidationDocuments is "ACCOUNTS_FILING", then UltimateAccountingPeriodStart should not be blank.'])
                              toolTipLst.append('Error Logic:'+'\n'+'If UltimateValidationDocuments is "ACCOUNTS_FILING", then UltimateAccountingPeriodStart should not be blank.')     
                    except:
                        pass
            
#           ------------------------------UltimateRelationshipPeriodStart-----------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(87,0).value)==True:
                if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A" ) :
                  try:
                     if UlRlshpType.strip()=="IS_ULTIMATELY_CONSOLIDATED_BY":
                         if UlRlshpPrdStrt=="":
                             errMsg.append(['UltimateRelationshipPeriodStart',UlRlshpPrdStrt,'If UltimateRelationshipType is "IS_ULTIMATELY_CONSOLIDATED_BY", then UltimateRelationshipPeriodStart should not be blank.'])
                             toolTipLst.append('Error Logic:'+'\n'+'If UltimateRelationshipType is "IS_ULTIMATELY_CONSOLIDATED_BY", then UltimateRelationshipPeriodStart should not be blank.')
                  except:
                      pass
#           ----------------------------- UltimateFilingPeriodStart---------------updated 4/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(89,0).value)==True:
               if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A" ) :
                   try:
                       if "REGULATORY_FILING" in UlVldnDocs.upper().strip():
                           if UlFlngPrdStrt=="":
                               errMsg.append(['UltimateFilingPeriodStart',str(UlFlngPrdStrt),'If UltimateValidationDocuments has "REGULATORY_FILING" then UltimateFilingPeriodStart can not be blank.'])                 
                               toolTipLst.append('Error Logic:'+'\n'+'If UltimateValidationDocuments has "REGULATORY_FILING" then UltimateFilingPeriodStart can not be blank.')  
                   except:
                       pass

##           -----------------------------PNI2HeadquartersAddressCity-----------------------------
#            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(97,0).value)==True:
#             try:   
##            L1
#              if UlPrnt!="" and UlPrnt.lower()!="n/a" and PNI2HQAddCity=='':
#                  errMsg.append(['PNI2HeadquartersAddressCity',PNI2HQAddCity,'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not be blank.'])
#                  toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not be blank.')
#             except:
#                 pass
#
#             if UlPrnt!="" and UlPrnt.lower()!="n/a" and PNI2HQAddCity!="" and PNI2HQAddCntry!="":
#              try:
#               if (PNI2HQAddCity.upper() not in cntExList):
#                  if PNI2HQAddCntry.lower()in PNI2HQAddCity.lower():
#                     errMsg.append(['PNI2HeadquartersAddressCity',PNI2HQAddCity,'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity field should not have a country name in it except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" '])                               
#                     toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not have a country name except:"Luxembourg", "Singapore","Hong Kong" and "Gibraltar" ') 
#              except:
#                  pass
#              
#             if UlPrnt!="" and UlPrnt.lower()!="n/a" and PNI2HQAddCity!='':   
#              try:   
##              L2
#               if re.findall(r'\d+',PNI2HQAddCity)!=[]:
#                 errMsg.append(['PNI2HeadquartersAddressCity',PNI2HQAddCity,'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not have numeric value in it.'])
#                 toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent is anything except N/A then PNI2HeadquartersAddressCity should not have numeric value in it.')                 
#              except:
#                  pass
            '''
#           -----------------------------Requestor-----------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(3,0).value)==True:  
             try: 
                 if req!=None:
                      if req!="" and ('gmail' in str(req).lower() or "yahoo" in str(req).lower() or "aol" in str(req).lower() or "hotmail" in str(req).lower() or "outlook" in str(req).lower()):
                        errMsg.append(['Requestor',req,'Please check for authorisation as requestor does not have a company Email domain.'])                 
                        toolTipLst.append('Error Logic:'+'\n'+'Please check for authorisation as requestor does not have a company Email domain.') 
             except:
                 pass
            '''  
            #--------------------Dup Chk Res-----------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(117,0).value)==True:
                try:
                    if dupChkRes.lower().strip()!= "complete":
                        errMsg.append(['Dup Chk Res',dupChkRes,'Dup Chk Res is not "Complete".'])                 
                        toolTipLst.append('Error Logic:'+'\n'+'Dup Chk Res is not "Complete".')    
                except:
                    pass
                 
            # ---------------------------------OfficialEntityName------------updated 26th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(13,0).value)==True:
             try:  
                 if officialEnNm!=None:
                      '''
                      if officialEnNm=='':errMsg.append(['OfficialEntityName',officialEnNm,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'officialEntity Name should not be blank.')
                      '''
                      offEnName=officialEnNm  
                      
        #             make a list of all punctuations and digits in the off entity name 
                      alphPuncLst=[i for i in offEnName if all(j.isdigit() or j in string.punctuation for j in i)]
                      if len(offEnName)==len(alphPuncLst):
                            errMsg.append(['OfficialEntityName',offEnName,'Entity name can not be just figures or punctuation marks.'])                 
                            toolTipLst.append('Error Logic:'+'\n'+'Entity name can not be just figures or punctuation marks.')  
             except:
                pass
           
#          --------------------------------AlternateEntityName--------------------updated 26th june 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(104,0).value)==True:   
             errorFld=""
             try: 
                 if altEnName!=None and officialEnNm!=None:
                      if altEnName!="":
                          if str(altEnName).lower().strip()==str(officialEnNm).lower().strip():                     
                              errorFld="OfficialEntityName"
             except:
                 pass
             try:
                 if prevEnNam!=None and altEnName!=None:
                     if prevEnNam!="" and altEnName!="":
                         if str(altEnName).lower().strip()==str(prevEnNam).lower().strip():
                             errorFld=errorFld+", "+"Previous Entity Name"
             except:
                 pass
             try:
                 if angEnName!=None and altEnName!=None:
                     if angEnName!="" and altEnName!="":
                         if str(altEnName).lower().strip()==str(angEnName).lower().strip():
                             errorFld=errorFld+", "+"Anglicised Entity Name" 
             except:
                 pass
             try:    
              if errorFld!="":    
                errMsg.append(['AlternateEntityName',altEnName,'AlternateEntityName should not be same as'+' '+errorFld+'.'])                 
                toolTipLst.append('Error Logic:'+'\n'+'AlternateEntityName should not be same as'+' '+errorFld+'.') 
             except:
                 pass
          
#           -----------------------------Previous Entity Name-------------updated 26th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(103,0).value)==True: 
             errorFld=""
             try:
                 if prevEnNam!=None and officialEnNm!=None:
                  if prevEnNam!="":
                      if str(prevEnNam).lower().strip()==str(officialEnNm).lower().strip():                     
                          errorFld="OfficialEntityName" 
             except:
                 pass
             try:
                 if prevEnNam!=None and altEnName!=None:
                     if prevEnNam!="" and altEnName!="":
                          if str(altEnName).lower().strip()==str(prevEnNam).lower().strip():
                              errorFld=errorFld+", "+"AlternateEntityName"
             except:
                 pass
             try:
                 if angEnName!=None and prevEnNam!=None:
                      if angEnName!="" and prevEnNam!="":
                          if str(prevEnNam).lower()==str(angEnName).lower():
                              errorFld=errorFld+", "+"Anglicised Entity Name" 
             except:
                 pass
             try:
              if errorFld!="":
                  
                if str(prevEnNam).lower()==str(officialEnNm).lower() or str(prevEnNam).lower()==str(altEnName).lower() or str(prevEnNam).lower()==str(angEnName).lower():
                   errMsg.append(['Previous Entity Name',prevEnNam,'Previous Entity Name should not be same as'+' '+errorFld+'.'])                 
                   toolTipLst.append('Error Logic:'+'\n'+'Previous Entity Name should not be same as'+' '+errorFld+'.') 
             except:
                 pass
                   
                   
#           -------------------EntityCategory-----------------------updated 26th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(105,0).value)==True:               

              lglFrmLst=['Fund', 'Fondo', 'Fonds', 'Fondos', 'Inversion', 'UCITS', 'ICVC', 'SICAV']
              notLglFrmLst=['Fund', 'Sub-Fund', 'Fondo', 'Fonds', 'Fondos', 'Inversion']
              
#             LOGIC 1 
              if EntCat.lower().strip()=='fund':
                  if not(entLglForm.strip() in lglFrmLst):
                      errMsg.append(['EntityCategory',EntCat,'If "Entity Category" is "Fund" then "Entity Legal Form" should contain - Fund, Fondo, Fonds, Fondos, Inversion, UCITS, ICVC, SICAV.'])                 
                      toolTipLst.append('Error Logic:'+'\n'+'If "Entity Category" is "Fund" then "Entity Legal Form" should contain - Fund, Fondo, Fonds, Fondos, Inversion, UCITS, ICVC, SICAV.')
                      
                  if  "OTHER - please specify" in entLglForm.strip():
                      if not("fund" in LegalFormTxt.strip().lower()):
                          errMsg.append(['EntityCategory',EntCat,'If "Entity Category" is "Fund" and "Entity Legal Form" has "OTHER - please specify" then LegalFormFreeText must contain the word "Fund".'])                 
                          toolTipLst.append('Error Logic:'+'\n'+'If "Entity Category" is "Fund" and "Entity Legal Form" has "OTHER - please specify" then LegalFormFreeText must contain the word "Fund".') 
              
#              LOGIC 2                
              elif EntCat.lower().strip()=="n/a" or EntCat.lower().strip()=="branch" or EntCat.lower().strip()=="sole proprietor" :
                  if entLglForm.strip() in notLglFrmLst:
                      errMsg.append(['EntityCategory',EntCat,'If Entity Category is N/A (OR) Branch (OR) Sole Proprietor - then Entity Legal Form should not have - Fund, Sub-Fund, Fondo, Fonds, Fondos, Inversion.'])                 
                      toolTipLst.append('Error Logic:'+'\n'+'If Entity Category is N/A (OR) Branch (OR) Sole Proprietor - then Entity Legal Form should not have - Fund, Sub-Fund, Fondo, Fonds, Fondos, Inversion.') 
            
            '''    
#           ---------------------------------Direct LEI---------------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(57,0).value)==True: 
               try:
                   if str(DirPrnt).upper().lstrip().rstrip()=="N/A" or str(DirPrnt).upper().lstrip().rstrip()=="NA" :
                     if DirLEI!="" and DirLEI!=None:  
                       if not(str(DirLEI).lower().lstrip().rstrip()=="N/A" or str(DirLEI).upper().lstrip().rstrip()=="NA"): 
                           errMsg.append(['DirectLEI',DirLEI,'If DirectParent is N/A then DirectLEI should also be NA'])                 
                           toolTipLst.append('Error Logic:'+'\n'+'If DirectParent is N/A then DirectLEI should also be NA')
               except:
                 pass 
            ''' 
#           -----------------------------------------UltimateLEI------------------------removed 4/7/2018
            ''' 
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(79,0).value)==True: 
               try:
                   if UlPrnt!="" and UlPrnt!=None:
                    if str(UlPrnt).lower()=="n/a" or str(UlPrnt).lower()=="na":
                      if UlLEI!="" and UlLEI!=None:  
                       if not(str(UlLEI).lower()=="n/a" or str(UlLEI).lower()=="na"):
                          errMsg.append(['UltimateLEI',UlLEI,'If UltimateParent = N/A then UltimateLEI should be N/A' +"not"+str(UlLEI)])                 
                          toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then UltimateLEI should be N/A'+"not"+str(UlLEI))                      
               except:
                   pass
            '''   
#          ------------------------------------------DirectExceptionReason--------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(52,0).value)==True:  
                try:
                    if DirPrnt!=None:
                       if str(DirPrnt).upper().lstrip().rstrip()=="N/A" or str(DirPrnt).upper().lstrip().rstrip()=="NA" :
                           if DirExcptnRsn=="":
                               errMsg.append(['DirectExceptionReason',DirExcptnRsn,'DirectExceptionReason should not be blank if DirectParent has N/A'])                 
                               toolTipLst.append('Error Logic:'+'\n'+'DirectExceptionReason should not be blank if DirectParent has N/A') 
                except:
                    pass
#          ----------------------------------UltimateExceptionReason--------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(77,0).value)==True:  
                try:
                   if UlPrnt.upper()=="N/A":
                       if UlExcptRsn=="":
                           errMsg.append(['UltimateExceptionReason',UlExcptRsn,"If UltimateParent has 'N/A', then UltimateExceptionReason should not be blank."])                 
                           toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has 'N/A', then UltimateExceptionReason should not be blank.")  
                except:
                    pass  
                
                
                
#           ---------------------------------CountryLegalForm----------------updated 26th jun 2018   
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(15,0).value)==True: 
               try:
                 if not(cntry.strip().lower()== BRCountry.lower().strip()==LFCountry1.lower().strip()==VldnAuthIDcntry.lower().strip()):
                      errMsg.append(['CountryLegalForm',cntry,'CountryLegalForm,BusinessRegistryCountry,LegalFormationAddressCountry and ValidationAuthorityIDCountry must be equal.'])                 
                      toolTipLst.append('Error Logic:'+'\n'+'CountryLegalForm,BusinessRegistryCountry,LegalFormationAddressCountry and ValidationAuthorityIDCountry must be equal.')
                   
                                   
                   
                 '''  
                 for row_num in range(elfSht.nrows):
                     row_value = elfSht.row_values(row_num)                     
                     if (row_value[0]).lower() == cntry.lower():
                         if unicode(row_value[1])==unicode(LegalFormTxt):
                             errMsg.append(['CountryLegalForm',cntry,'Legal form already part of ELF list and should not be added to Specify Others.'])                 
                             toolTipLst.append('Error Logic:'+'\n'+'Legal form already part of ELF list and should not be added to Specify Others.') 
                 
                 if cntry=='':errMsg.append(['CountryLegalForm',cntry,'Field value is blank']);toolTipLst.append('Error Logic:'+'\n'+'CountryLegalForm should not be blank.')
                 else:                      
                  cntryLglFormLst=["fund","sub-fund","fondo","fonds","fondos","inversion","ucits","icvc"]
                  if str(LegalForm).lower() not in cntryLglFormLst or "other - please specify" not in str(LegalForm).lower().lstrip().rstrip():  
                     if str(cntry).lower()!=str(BRCountry).lower() and str(cntry).lower()!=LFCounty.lower():
                        errMsg.append(['CountryLegalForm',cntry,'Should match with text in "BusinessRegistryCountry" and "LegalFormationAddressCountry".'])                 
                        toolTipLst.append('Error Logic:'+'\n'+'Should match with text in "BusinessRegistryCountry" and "LegalFormationAddressCountry".')                     
                  elif "other - please specify" in str(LegalForm).lower().lstrip().rstrip():             
                      if not "fund" in str(LegalFormTxt).lower():
                         errMsg.append(['CountryLegalForm',cntry,'If "LegalForm" is Other then "LegalFormFreeText" should contain - Fund.'])                 
                         toolTipLst.append('Error Logic:'+'\n'+'If "LegalForm" is Other then "LegalFormFreeText" should contain - Fund.') 
                 '''        
               except:
                   pass
            
            
#           --------------------------HeadquartersAddressAddressNumber---------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(106,0).value)==True: 
              try: 
                  '''
                  if HQAddAddNo!=None:
                      if HQAddAddNo!="" and not HQAddAddNo.isdigit(): 
                          errMsg.append(['HeadquartersAddressAddressNumber',HQAddAddNo,' Address number should not have anything except numeric data.']); toolTipLst.append('Error Logic:'+'\n'+'Address number should not have anything except numeric data.') 
                  '''
                  if HQAddAddNo!="" :
                      errMsg.append(['HeadquartersAddressAddressNumber',HQAddAddNo,' HeadquartersAddressAddressNumber number should be blank.'])
                      toolTipLst.append('Error Logic:'+'\n'+'HeadquartersAddressAddressNumber number should be blank.') 
              except:
                  pass
            
#           ------------------------------ HeadquartersAddressAddressNumberWithinBuilding---------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(107,0).value)==True:
              try:  
                  if HQAddAddNoBldg!="":
                      errMsg.append(['HeadquartersAddressAddressNumberWithinBuilding',HQAddAddNoBldg,'"HeadquartersAddressAddressNumberWithinBuilding" field should be blank.']); toolTipLst.append('Error Logic:'+'\n'+'"HeadquartersAddressAddressNumberWithinBuilding" field should be blank.') 
              except:
                  pass
                
            
#           ---------------------------------LegalFormationAddressAddressNumber------------------updated 27th jun 2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(108,0).value)==True:  
              try:
                  '''
                  if LglFrmnAddAddNo!=None:
                    if LglFrmnAddAddNo!="" and not LglFrmnAddAddNo.isdigit():
                       errMsg.append(['LegalFormationAddressAddressNumber',LglFrmnAddAddNo,' LegalFormationAddressAddressNumber should not have anything except numeric data.'])                 
                       toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressAddressNumber should not have anything except numeric data.') 
                  ''' 
                  if LglFrmnAddAddNo!="":
                      errMsg.append(['LegalFormationAddressAddressNumber',LglFrmnAddAddNo,' LegalFormationAddressAddressNumber should be blank.'])                 
                      toolTipLst.append('Error Logic:'+'\n'+'LegalFormationAddressAddressNumber should be blank.')  
              except:
                  pass

            
#          ----------------------------------------- ValidationAuthorityIDCountry------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(109,0).value)==True: 
             try: 
                 if VldnAuthIDcntry!=None and BRCountry!=None and VldnAuthIDcntry!=None:
                  if VldnAuthIDcntry!="" and BRCountry!="" and str(VldnAuthIDcntry).lower()!=str(BRCountry).lower():
                    errMsg.append(['ValidationAuthorityIDCountry',VldnAuthIDcntry,' ValidationAuthorityIDCountry should match BusinessRegistryCountry.'])                 
                    toolTipLst.append('Error Logic:'+'\n'+'ValidationAuthorityIDCountry should match BusinessRegistryCountry.')
             except:
                 pass

            
#           --------------------------------------ValidationAuthorityID--------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(110,0).value)==True: 
             try:        
                if str(VlnAuthId).lower()!=str(OBReg).lower():
                     errMsg.append(['ValidationAuthorityID',VlnAuthId,'ValidationAuthorityID should match RegistrationAuthorityID.'])                 
                     toolTipLst.append('Error Logic:'+'\n'+'ValidationAuthorityID should match RegistrationAuthorityID.')
             except:
                 pass
             
             try:
                 if str(VlnAuthId).strip() in str(OthrValAuthId).strip():
                     errMsg.append(['ValidationAuthorityID',VlnAuthId,'"ValidationAuthorityID" should not contain the text in "OtherValidationAuthorityID".'])                 
                     toolTipLst.append('Error Logic:'+'\n'+'"ValidationAuthorityID" should not contain the text in "OtherValidationAuthorityID".')
             except:
                 pass
            
            
#           -----------------------------------OtherValidationAuthorityID---------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(111,0).value)==True:
              try: 
                  if OthrValAuthId!=None and OBRegTxt!=None:
                    if str(OthrValAuthId).lower()==str(VlnAuthId).lower() or str(OthrValAuthId).lower()==str(VlnAuthEnId).lower():
                      errMsg.append(['OtherValidationAuthorityID',OthrValAuthId,'OtherValidationAuthorityID should not match ValidationAuthorityID or ValidationAuthorityEntityID.'])                 
                      toolTipLst.append('Error Logic:'+'\n'+'OtherValidationAuthorityID should not match ValidationAuthorityID or ValidationAuthorityEntityID.')
              except:
                  pass
              try:
                  if "N/A" in OthrValAuthId.upper() or "NA" in OthrValAuthId.upper() or  OthrValAuthId.upper()==".":
                      errMsg.append(['OtherValidationAuthorityID',OthrValAuthId,'OtherValidationAuthorityID should not contain NA, N/A or "."'])                 
                      toolTipLst.append('Error Logic:'+'\n'+'OtherValidationAuthorityID should not contain NA, N/A or "."') 
              except:
                  pass
          
          
#           -------------------------------ValidationAuthorityEntityID-------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(112,0).value)==True: 
             try:   
                 if VlnAuthEnId!=None and offBusRegRef!=None:
                   if VlnAuthEnId!="" and offBusRegRef!="" and str(VlnAuthEnId).lower()!=str(offBusRegRef).lower():
                      errMsg.append(['ValidationAuthorityEntityID',VlnAuthEnId,'ValidationAuthorityEntityID should match OfficialBusinessRegistryReference.'])                 
                      toolTipLst.append('Error Logic:'+'\n'+'ValidationAuthorityEntityID should match OfficialBusinessRegistryReference.')
             except:
                 pass
             '''
             try:
               if VlnAuthEnId=="" or "n/a" in str(VlnAuthEnId).lower().replace(" ",""):errMsg.append(['ValidationAuthorityEntityID',VlnAuthEnId, 'ValidationAuthorityEntityID can not be blank or N/A.']);toolTipLst.append('Error Logic:'+'\n'+'ValidationAuthorityEntityID can not be blank or N/A.')
             except:
                 pass
             '''
#           ------------------------------DirectRelationshipQualifierCategory-----------------removed 3/7/2018
            ''' 
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(62,0).value)==True: 
                try:
                  if DirPrnt!=None and DirPrnt!="":  
                    if str(DirPrnt).lower()=="na" or str(DirPrnt).lower()=="n/a" :
                      if DirRlnQlfr!=None:  
                       if DirRlnQlfr!="":
                         errMsg.append(['DirectRelationshipQualifierCategory',DirRlnQlfr,'If DirectParent = N/A then DirectRelationshipQualifierCategory should be blank.'+str(DirRlnQlfr)])
                         toolTipLst.append('Error Logic:'+'\n'+'If DirectParent = N/A then DirectRelationshipQualifierCategory should be blank.'+str(DirRlnQlfr))
                     
                    if not(str(DirPrnt).lower().lstrip().rstrip()=="na" or str(DirPrnt).lower().lstrip().rstrip()=="n/a"):                  
                       if DirRlnQlfr!=None:    
                          if DirRlnQlfr=="":
                             errMsg.append(['DirectRelationshipQualifierCategory',DirRlnQlfr,"If DirectParent has anything other than 'N/A' then DirectRelationshipQualifierCategory should not be blank."+str(DirRlnQlfr)])
                             toolTipLst.append('Error Logic:'+'\n'+"If DirectParent has anything other than 'N/A' then DirectRelationshipQualifierCategory should not be blank."+str(DirRlnQlfr)) 
                except:
                    pass
            '''    
#           ------------------------------UltimateRelationshipQualifierCategory-----------------deleted 4/7/2018
            '''
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(84,0).value)==True: 
                try:
                  if UlPrnt!=None and UlRlshpQlfrCat!=None:  
                    if UlPrnt!="" and str(UlPrnt).lower()=="na" or str(UlPrnt).lower()=="n/a" :
                       
                      if UlRlshpQlfrCat!="":
                         errMsg.append(['UltimateRelationshipQualifierCategory',UlRlshpQlfrCat,'If UltimateParent = N/A then UltimateRelationshipQualifierCategory should be blank.'])
                         toolTipLst.append('Error Logic:'+'\n'+'If UltimateParent = N/A then UltimateRelationshipQualifierCategory should be blank.')
                     
                    if UlPrnt!="" and not(str(UlPrnt).lower().lstrip().rstrip()=="na" or str(UlPrnt).lower().lstrip().rstrip()=="n/a"):                  
                          if UlRlshpQlfrCat=="":
                             errMsg.append(['UltimateRelationshipQualifierCategory',UlRlshpQlfrCat,"If UltimateParent has anything other than 'N/A' then UltimateRelationshipQualifierCategory should not be blank."])
                             toolTipLst.append('Error Logic:'+'\n'+"If UltimateParent has anything other than 'N/A' then UltimateRelationshipQualifierCategory should not be blank.") 
                except:
                    pass 
            '''    
#           ------------------------------DirectValidationSources--------------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(59,0).value)==True:
                if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A") :  
                    if DirValdnSrcs.strip()=="ENTITY_SUPPLIED_ONLY":                
                        if dirVldnRef.strip()!="URL not publicly available, validated from document provided by client"  :
                            errMsg.append(['DirectValidationSources',DirValdnSrcs,'If "Directvalidationsource" is ENTITY_SUPPLIED_ONLY then "DirectValidationReference must be "URL not publicly available, validated from document provided by client".'])
                            toolTipLst.append('Error Logic:'+'\n'+'If "Directvalidationsource" is ENTITY_SUPPLIED_ONLY then "DirectValidationReference must be "URL not publicly available, validated from document provided by client".')
               
               
                 
#           ------------------------------UltimateValidationSources--------------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(82,0).value)==True: 
                if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A")  :
                     try:
                        if UlVldnSrcs.strip()=="ENTITY_SUPPLIED_ONLY": 
                            if "URL not publicly available, validated from document provided by client" not in ulVldnRef.strip():
                                errMsg.append(['UltimateValidationSources',UlVldnSrcs,'If "UltimateValidationSources" is "ENTITY_SUPPLIED_ONLY" then "UltimateValidationReference" should be "URL not publicly available, validated from document provided by client".'])
                                toolTipLst.append('Error Logic:'+'\n'+'If "UltimateValidationSources" is "ENTITY_SUPPLIED_ONLY" then "UltimateValidationReference" should be "URL not publicly available, validated from document provided by client".')
                     except:
                        pass
                
                 
#           ------------------------------DirectValidationReference--------------------updated 5/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(121,0).value)==True:
                if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A")  :
                    try: 
                        if DirVldnDocs.strip()=="REGULATORY_FILING" or DirVldnDocs.strip()=="ACCOUNTS_FILING":
                            if DirValdnSrcs.strip().upper()=="ENTITY_SUPPLIED_ONLY":
                                if dirVldnRef.strip()!="URL not publicly available, validated from document provided by client":
                                    errMsg.append(['DirectValidationReference',dirVldnRef,"If 'DirectValidationDocuments' is 'REGULATORY_FILING' or 'ACCOUNTS_FILING' and 'DirectValidationSources' is 'ENTITY_SUPPLIED_ONLY', then 'DirectValidationReference' should be 'URL not publicly available, validated from document provided by client'."])
                                    toolTipLst.append('Error Logic:'+'\n'+"If 'DirectValidationDocuments' is 'REGULATORY_FILING' or 'ACCOUNTS_FILING' and 'DirectValidationSources' is 'ENTITY_SUPPLIED_ONLY', then 'DirectValidationReference' should be 'URL not publicly available, validated from document provided by client'.") 
                            else:        
                                if dirVldnRef=="":
                                    errMsg.append(['DirectValidationReference',dirVldnRef,"If 'DirectValidationDocuments' is REGULATORY_FILING or ACCOUNTS_FILING, then DirectValidationReference should not be blank."])
                                    toolTipLst.append('Error Logic:'+'\n'+"If 'DirectValidationDocuments' is REGULATORY_FILING or ACCOUNTS_FILING, then DirectValidationReference should not be blank.")        
                    except:
                        pass
                    try:
                       if len(dirVldnRef)>499:
                            errMsg.append(['DirectValidationReference',dirVldnRef,"If 'DirectValidationDocuments' must be less than or equal to 499 characters."])
                            toolTipLst.append('Error Logic:'+'\n'+"If 'DirectValidationDocuments' must be less than or equal to 499 characters.")
                    except:
                        pass
                    try:
                        if not(DirVldnDocs.strip()=="REGULATORY_FILING" or DirVldnDocs.strip()=="ACCOUNTS_FILING"): 
                            if dirVldnRef!="":
                                errMsg.append(['DirectValidationReference',dirVldnRef,"If 'DirectValidationDocuments' contains anything other than REGULATORY_FILING or ACCOUNTS_FILING, then DirectValidationReference should be blank."])
                                toolTipLst.append('Error Logic:'+'\n'+"If 'DirectValidationDocuments' contains anything other than REGULATORY_FILING or ACCOUNTS_FILING, then DirectValidationReference should be blank.")        
                    except:
                        pass
            
#            ------------------------------UltimateValidationReference--------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(122,0).value)==True: 
                if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A" ):  
                    try: 
                        if UlVldnDocs.strip()=="REGULATORY_FILING" or UlVldnDocs.strip()=="ACCOUNTS_FILING":
                            if UlVldnSrcs.strip().upper()=="ENTITY_SUPPLIED_ONLY":
                                if ulVldnRef.strip()!="URL not publicly available, validated from document provided by client":
                                    errMsg.append(['UltimateValidationReference',ulVldnRef,"If 'UltimateValidationDocuments' is 'REGULATORY_FILING' or 'ACCOUNTS_FILING' and 'UltimateValidationSources' is 'ENTITY_SUPPLIED_ONLY', then 'UltimateValidationReference' should be 'URL not publicly available, validated from document provided by client'."])
                                    toolTipLst.append('Error Logic:'+'\n'+"If 'UltimateValidationDocuments' is 'REGULATORY_FILING' or 'ACCOUNTS_FILING' and 'UltimateValidationSources' is 'ENTITY_SUPPLIED_ONLY', then 'UltimateValidationReference' should be 'URL not publicly available, validated from document provided by client'.") 
                            else:        
                                if ulVldnRef=="":
                                    errMsg.append(['UltimateValidationReference',ulVldnRef,"If 'UltimateValidationDocuments' is REGULATORY_FILING or ACCOUNTS_FILING, then UltimateValidationReference should not be blank."])
                                    toolTipLst.append('Error Logic:'+'\n'+"If 'UltimateValidationDocuments' is REGULATORY_FILING or ACCOUNTS_FILING, then UltimateValidationReference should not be blank.")        
                    except:
                        pass
                    
                    try:
                        if len(ulVldnRef)>499:
                            errMsg.append(['UltimateValidationReference',ulVldnRef,"If 'UltimateValidationDocuments' must be less than or equal to 499 characters."])
                            toolTipLst.append('Error Logic:'+'\n'+"If 'UltimateValidationDocuments' must be less than or equal to 499 characters.")       
                    except:
                        pass
                    try:
                        if not(UlVldnDocs.strip()=="REGULATORY_FILING" or UlVldnDocs.strip()=="ACCOUNTS_FILING"): 
                            if ulVldnRef!="":
                                errMsg.append(['DirectValidationReference',dirVldnRef,"If 'DirectValidationDocuments' contains anything other than REGULATORY_FILING or ACCOUNTS_FILING, then DirectValidationReference should be blank."])
                                toolTipLst.append('Error Logic:'+'\n'+"If 'DirectValidationDocuments' contains anything other than REGULATORY_FILING or ACCOUNTS_FILING, then DirectValidationReference should be blank.")        
                    except:
                        pass
#           ------------------------------------Comment-----------------commented 2/7/2018
            '''     
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(113,0).value)==True:  
                try:   
                  if "hold" in str(leiEvnt).lower() and Cmnt=="":
                      errMsg.append(['Comment',Cmnt,'If LEI Event is "On Hold" then Comment should not be blank.'])                 
                      toolTipLst.append('Error Logic:'+'\n'+'If LEI Event is "On Hold" then Comment should not be blank.') 
                except:
                     pass
        
            '''        
#           --------------------------------Create Date----------------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(114,0).value)==True: 
             try:   
                 if CrtDate!='':
                   
                    date1 = (CrtDate - 25569) * 86400.0
                    crtDateFrmla = datetime.utcfromtimestamp(date1).strftime("%Y-%m-%d")
                 
                    if type(CrtDate)==unicode or type(CrtDate)==str:
                            dateObj=pandas.to_datetime(CrtDate)                                   
                    else:  
                        dateObj= xlrd.xldate.xldate_as_datetime(CrtDate, wb.datemode)
                    crt_dt=datetime.date(dateObj)    
                 if ARDate!='':                 
                   try:
                    if type(ARDate)==unicode or type(ARDate)==str:
                            dateObj=pandas.to_datetime(ARDate)                                   
                    else:    
                        dateObj= xlrd.xldate.xldate_as_datetime(ARDate, wb.datemode)
                    ar_dt=datetime.date(dateObj)  
                   except:
                   
                    traceback.print_exc()
    #             if type(CrtDate)==float:
    #                crtDateFrmla="="+str(CrtDate)+"*1"
                 if CrtDate!='' and ARDate!='':
                     if crt_dt>ar_dt:
                       errMsg.append(['Create Date',crtDateFrmla,'Annual Renewal Date should be post CreateDate.'])                 
                       toolTipLst.append('Error Logic:'+'\n'+'Annual Renewal Date should be post CreateDate.') 
             except:
                   pass


#           ----------------------------------RegistrationAuthorityID-----------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(32,0).value)==True:  
#             new logics-L1 &L2
             try:
                 if "RA888888 - OTHER - PLEASE SPECIFY" in OBReg.upper().strip():
                     if OBRegTxt=="" or offBusRegRef=="":
                         errMsg.append(['RegistrationAuthorityID',OBReg,"If the 'RegistrationAuthorityID' is 'OTHER - PLEASE SPECIFY' then 'OfficialBusinessRegistryFreeText', 'OfficialBusinessRegistryReference' should not be blank."])                 
                         toolTipLst.append('Error Logic:'+'\n'+"If the 'RegistrationAuthorityID' is 'OTHER - PLEASE SPECIFY' then 'OfficialBusinessRegistryFreeText', 'OfficialBusinessRegistryReference' should not be blank.") 
                 elif "RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation" in OBReg.strip():
                     if offBusRegRef.strip()!="":
                         errMsg.append(['RegistrationAuthorityID',OBReg,"If the 'RegistrationAuthorityID' is 'RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation -' then 'OfficialBusinessRegistryReference'should be blank."])                 
                         toolTipLst.append('Error Logic:'+'\n'+"If the 'RegistrationAuthorityID' is 'RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation -' then 'OfficialBusinessRegistryReference'should be blank.") 
             except:
                 pass
             
            
#             try:  
#                 if OBReg!=None:
#                    if OBReg!='' and OBReg.lstrip().rstrip()=='RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation' and str(OBRegRef).upper()!="DOCUMENT":
#                       errMsg.append(['RegistrationAuthorityID',OBReg,'If RegistrationAuthorityID is "RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation" then  OfficialBusinessRegistryReference SHOULD BE "DOCUMENT".'])                 
#                       toolTipLst.append('Error Logic:'+'\n'+'If RegistrationAuthorityID is "RA999999 - Registry NA (Trust, Sub Fund) - Upload Supporting Documentation" then  OfficialBusinessRegistryReference SHOULD BE "DOCUMENT".') 
#             except:
#                pass
#            
#           --------------------------------DirectAccountingPeriodEnd------------------updated 3/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(63,0).value)==True: 
                if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A") :  
                    try: 
                        if DirAccPrdEnd!="":    
                          try:  
                            dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=DirAccPrdEnd-2)
                            dirAccPrdEndDate =dtRead.date()
                          except:
                              errMsg.append(['DirectAccountingPeriodEnd',DirAccPrdEnd,'DirectAccountingPeriodEnd should be be a date or blank.'])                 
                              toolTipLst.append('Error Logic:'+'\n'+'DirectAccountingPeriodEnd should be a date or blank.')
                          if DirAccPrdStart!="":
                              dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=DirAccPrdStart-2)
                              dirAccPrdStrtDate =dtRead.date()
                          if  dirAccPrdStrtDate> dirAccPrdEndDate:
                              errMsg.append(['DirectAccountingPeriodEnd',DirAccPrdEnd,'DirectAccountingPeriodEnd date must be ahead of DirectAccountingPeriodStart date.'])                 
                              toolTipLst.append('Error Logic:'+'\n'+'DirectAccountingPeriodEnd date must be ahead of DirectAccountingPeriodStart date.')
                    except:
                        pass
            
#           ----------------------------DirectRelationshipPeriodEnd-----------------updated 4/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(64,0).value)==True: 
                if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A"):  
                    try:
                        if DirRlnshpPrdEnd!="": 
                          errMsg.append(['DirectRelationshipPeriodEnd',DirRlnshpPrdEnd,'DirectRelationshipPeriodEnd date must be blank.'])                 
                          toolTipLst.append('Error Logic:'+'\n'+'DirectRelationshipPeriodEnd date must be blank.')  
                          try:  
                            dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=DirRlnshpPrdEnd-2)
                            dirRlnshpPrdEndDate =dtRead.date()
                          except:
                              errMsg.append(['DirectRelationshipPeriodEnd','DirectRelationshipPeriodEnd should be be a date or blank.'])                 
                              toolTipLst.append('Error Logic:'+'\n'+'DirectRelationshipPeriodEnd should be a date or blank.')
                          if DirRlnshpPrdStrt!="":
                              dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=DirRlnshpPrdStrt-2)
                              dirRlnshpPrdStrtDate =dtRead.date()
                          if  dirRlnshpPrdStrtDate> dirRlnshpPrdEndDate:
                              errMsg.append(['DirectRelationshipPeriodEnd',DirRlnshpPrdEnd,'DirectRelationshipPeriodEnd date must be ahead of DirectRelationshipPeriodStart date.'])                 
                              toolTipLst.append('Error Logic:'+'\n'+'DirectRelationshipPeriodEnd date must be ahead of DirectRelationshipPeriodStart date.')
                    except:
                        pass
                
#              try:        
#                if DirRlnshpPrdStrt!="" and DirRlnshpPrdEnd!="":
#                    dateObj=None;dateObj1=None                   
#                    if type(DirRlnshpPrdEnd)==unicode or type(DirRlnshpPrdEnd)==str:
#                        dateObj=pandas.to_datetime(DirRlnshpPrdEnd)                                   
#                    else: dateObj=xlrd.xldate.xldate_as_datetime(sht.cell(rownum,DirRlnshpPrdEndCol).value, wb.datemode)                                                       
#                    dateVal=datetime.date(dateObj)
#                    if type(DirRlnshpPrdStrt)==unicode or type(DirRlnshpPrdStrt)==str:
#                        dateObj1=pandas.to_datetime(DirRlnshpPrdStrt)                                   
#                    else: dateObj1=xlrd.xldate.xldate_as_datetime(sht.cell(rownum,DirRlnshpPrdStrtCol).value, wb.datemode)                                                      
#                    dateVal1=datetime.date(dateObj1)                  
#                    if dateVal1>dateVal:
#                      errMsg.append(['DirectRelationshipPeriodEnd',DirRlnshpPrdEnd,'DirectRelationshipPeriodEnd date must be ahead of DirectRelationshipPeriodStart date.'])                 
#                      toolTipLst.append('Error Logic:'+'\n'+'DirectRelationshipPeriodEnd date must be ahead of DirectRelationshipPeriodStart date.')
#              except:
#                  pass
           
#           -------------------------------DirectFilingPeriodEnd--------------updated 4/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(66,0).value)==True: 
                if not(DirPrnt.upper().strip()=="NA" or DirPrnt.upper().strip()=="N/A"): 
                    try: 
                        if DirFlngPrdEnd!="":    
                              try:  
                                dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=DirFlngPrdEnd-2)
                                dirFlngPrdEndDate =dtRead.date()
                              except:
                                  errMsg.append(['DirectFilingPeriodEnd',DirFlngPrdEnd,'DirectFilingPeriodEnd should be be a date or blank.'])                 
                                  toolTipLst.append('Error Logic:'+'\n'+'DirectFilingPeriodEnd should be a date or blank.')
                              if DirFilngPrdStrt!="":
                                  dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=DirFilngPrdStrt-2)
                                  dirFlngPrdStrtDate =dtRead.date()
                              if  dirFlngPrdStrtDate> dirFlngPrdEndDate:
                                  errMsg.append(['DirectFilingPeriodEnd',DirFlngPrdEnd,'DirectFilingPeriodEnd date must be ahead of DirectFilingPeriodStart date.'])                 
                                  toolTipLst.append('Error Logic:'+'\n'+'DirectFilingPeriodEnd date must be ahead of DirectFilingPeriodStart date.')
                    except:
                       pass
            
#           ----------------------------UltimateAccountingPeriodEnd-------------------------
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(86,0).value)==True:
                if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A"):
                      try:
                          if UlAccPrdEnd!="":    
                              try:  
                                dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=UlAccPrdEnd-2)
                                ulAccPrdEndDate =dtRead.date()
                              except:
                                  errMsg.append(['UltimateAccountingPeriodEnd',UlAccPrdEnd,'UltimateAccountingPeriodEnd should be be a date or blank.'])                 
                                  toolTipLst.append('Error Logic:'+'\n'+'UltimateAccountingPeriodEnd should be a date or blank.')
                              if UlAccPrdStrt!="":
                                  dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=UlAccPrdStrt-2)
                                  ulAccPrdStrtDate =dtRead.date()
                              if  ulAccPrdStrtDate> ulAccPrdEndDate:
                                  errMsg.append(['UltimateAccountingPeriodEnd',UlAccPrdEnd,'UltimateAccountingPeriodEnd date must be ahead of UltimateAccountingPeriodStart date.'])                 
                      except:
                           pass
                      try:
                          if UlAccPrdEnd!="":
                              errMsg.append(['UltimateAccountingPeriodEnd',UlAccPrdEnd,'UltimateAccountingPeriodEnd should be blank.'])                 
                              toolTipLst.append('Error Logic:'+'\n'+'UltimateAccountingPeriodEnd should be blank.')
                      except:
                          pass
                
#           ----------------------------UltimateRelationshipPeriodEnd-------------------updated 4/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(88,0).value)==True:
               if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A" ) :
                  try: 
                     if UlRlshpPrdEnd!="":    
                      try:  
                        dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=UlRlshpPrdEnd-2)
                        ulRlshPrdEndDate =dtRead.date()
                      except:
                          errMsg.append(['UltimateRelationshipPeriodEnd',UlRlshpPrdEnd,'UltimateRelationshipPeriodEnd should be be a date or blank.'])                 
                          toolTipLst.append('Error Logic:'+'\n'+'UltimateRelationshipPeriodEnd should be a date or blank.')
                      if UlRlshpPrdStrt!="":
                          dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=UlRlshpPrdStrt-2)
                          ulRlshPrdStrtDate =dtRead.date()
                      if  ulRlshPrdStrtDate> ulRlshPrdEndDate:
                          errMsg.append(['UltimateRelationshipPeriodEnd',UlRlshpPrdEnd,'UltimateRelationshipPeriodEnd date must be ahead of UltimateRelationshipPeriodStart date.'])                 
                  except:
                     pass
                  try:
                      if UlRlshpPrdEnd!="":
                          errMsg.append(['UltimateRelationshipPeriodEnd',UlRlshpPrdEnd,'UltimateRelationshipPeriodEnd should be blank.'])                 
                          toolTipLst.append('Error Logic:'+'\n'+'UltimateRelationshipPeriodEnd should be blank.')
                  except:
                      pass
                  
#           ---------------------------UltimateFilingPeriodStart-------------------------------------      
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(89,0).value)==True:
                if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A") :
                   try:
                       if "REGULATORY_FILING" in UlVldnDocs.upper().strip():
                           if UlFlngPrdStrt=="":
                              errMsg.append(['UltimateFilingPeriodStart',UlFlngPrdStrt,"If UltimateValidationDocuments has 'REGULATORY_FILING' then UltimateFilingPeriodStart can not be blank."])                 
                              toolTipLst.append('Error Logic:'+'\n'+"If UltimateValidationDocuments has 'REGULATORY_FILING' then UltimateFilingPeriodStart can not be blank.") 
                       elif "REGULATORY_FILING" not in UlVldnDocs.upper().strip() or UlVldnDocs=="" : 
                           if UlFlngPrdStrt!="":
                              errMsg.append(['UltimateFilingPeriodStart',UlFlngPrdStrt,"If UltimateValidationDocuments is not 'REGULATORY_FILING' then UltimateFilingPeriodStart should be blank."])                 
                              toolTipLst.append('Error Logic:'+'\n'+"If UltimateValidationDocuments is not 'REGULATORY_FILING' then UltimateFilingPeriodStart should be blank.") 
                   except:
                       pass     
#           ----------------------------------UltimateFilingPeriodEnd---------------updated 4/7/2018
            if fieldList==[] or dataSht.nrows-1==len(fieldList) or IsfieldExist(fieldList,dataSht.cell(90,0).value)==True:  
                if not(UlPrnt.upper().strip()=="NA" or UlPrnt.upper().strip()=="N/A" ) :
                      try: 
                          if UlFlngPrdEnd!="":    
                              try:  
                                dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=UlFlngPrdEnd-2)
                                ulFlngPrdEndDate =dtRead.date()
                              except:
                                  errMsg.append(['UltimateFilingPeriodEnd',UlFlngPrdEnd,'UltimateFilingPeriodEnd should be be a date or blank.'])                 
                                  toolTipLst.append('Error Logic:'+'\n'+'UltimateFilingPeriodEnd should be a date or blank.')
                              if UlFlngPrdStrt!="":
                                  dtRead = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=UlFlngPrdStrt-2)
                                  ulFlngPrdStrtDate =dtRead.date()
                              if  ulFlngPrdStrtDate> ulFlngPrdEndDate:
                                  errMsg.append(['UltimateFilingPeriodEnd',UlFlngPrdEnd,'UltimateFilingPeriodEnd date must be ahead of UltimateFilingPeriodStart date.'])                 
                                  toolTipLst.append('Error Logic:'+'\n'+'UltimateFilingPeriodEnd date must be ahead of UltimateFilingPeriodStart date.')
                      except:
                         pass
                      try:
                          if UlFlngPrdEnd!="":
                              errMsg.append(['UltimateFilingPeriodEnd',UlFlngPrdEnd,'UltimateFilingPeriodEnd should be blank.'])                 
                              toolTipLst.append('Error Logic:'+'\n'+'UltimateFilingPeriodEnd should be blank.')
                      except:
                          pass
                 
        except:
            traceback.print_exc()
            
        if len(errMsg)>0:
            msgList.append(errMsg);toolTip.append(toolTipLst) 
            l= [keyId,o,officialEnNm,len(errMsg)];EntityLst.append(l) 
        if val>sht.nrows : break 
        else: bar.startLoop(val) ;bar.label2.setText('Processing QC for the entity -: ' + str(officialEnNm) +'................')     
        val=val+1
            
  except:
        traceback.print_exc()                 
    #except:pass        
#****************************GUI interface begin here***************************//
def main(app,maxRows,start_time):
    if getattr(sys,'frozen',False):
        exePath=os.path.dirname(os.path.realpath(sys.executable))
    elif __file__:
        exePath=os.path.dirname(os.path.abspath(__file__))
    end_time = datetime.datetime.now()
    dur=end_time - start_time
    if  EntityLst==[]: ctypes.windll.user32.MessageBoxA(0,"No errors found for the selected fields, Please select other fields and try again." ,"Quality Check-No data found",1);sys.exit(app.exec_()) 
    else:
        w= QcWindow(EntityLst,msgList,toolTip,exePath,maxRows,dur)       
        w.show()
        sys.exit(app.exec_())     #sys.exit() method ensures a clean exit from mail loop; exec is a Python keyword so exec_() was used instead.   
    #except Exception as e:ctypes.windll.user32.MessageBoxA(0,'Please make sure that you have selected input file'+'\n'+'For further assistance, Please connect with Transformation Team');sys.exit(1)
if __name__ == "__main__":
    try:
        SuperMainFn()
    except Exception as e:
        print e
        ctypes.windll.user32.MessageBoxA(0,"Error has occured :"  + str(e),'Evalueserve-LEI Data Quality Assessment',1)  
        


