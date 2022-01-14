import os
import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.Qt import Qt
import serial
import serial.tools.list_ports
import threading
import Terminal_ui3 as ui
import configure as config
import datetime
import configparser
import icon

class Main(QWidget, ui.Ui_MainWindow):
    print("class Main")    
    def __init__(self):
        print("Main__init__")   
        super().__init__()
        self.setupUi(self)
        self.Connect.setAutoFillBackground(True)

        self.Configure.clicked.connect(self.Configure_click)

         
        self.Connect.clicked.connect(self.Connect_click)
        self.Clear.clicked.connect(self.clear_click)
        self.Timestamp.clicked.connect(self.Timestamp_click)
        self.log.clicked.connect(self.openflie)
        self.command.returnPressed.connect(self.command_function) #Press Enter callback 
        self.Connect.setStyleSheet("background-color: rgb(255, 0, 0);")  
        self.Timestamp.setStyleSheet("background-color: rgb(0, 255, 0);")  
        self.OutputText.setFont(QFont('Consolas', 9))
        self.setWindowIcon(QIcon(':img/Icon.ico')) #若ICON在同一層目錄
        
        self.Configure.setFocusPolicy(Qt.ClickFocus)#NoFocus   ClickFocus
        self.Clear.setFocusPolicy(Qt.ClickFocus)
        self.Timestamp.setFocusPolicy(Qt.ClickFocus)
        self.log.setFocusPolicy(Qt.ClickFocus)
        self.Connect.setFocusPolicy(Qt.ClickFocus)
        self.OutputText.setFocusPolicy(Qt.ClickFocus)
        self.command.setFocusPolicy(Qt.ClickFocus)        
        self.frame.setFocusPolicy(Qt.ClickFocus)         
        self.setFocusPolicy(Qt.ClickFocus)       
        
        self.OutputText.installEventFilter(self)
      
        # 定时器接收数据
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.data_receive)      
         
        #self.ReadUARTThread = threading.Thread(target=self.ReadUART)
        #self.ReadUARTThread.start()
          
        #self.le  = MyLineEdit()
        
        #self.le.signalTabPressed[str].connect(self.update)
        
        print("baudrate = ",ser.baudrate)   
        self.com_open()
        self.OutputText.setReadOnly(False)
        
        self.OutputText.setStyleSheet("background-color: rgb(0, 0, 0);""color: rgb(255, 255, 255);")  
        
        #self.OutputText.setTextColor(QColor(255, 0, 0))
        #self.OutputText.setStyleSheet("background-color: rgb(0, 0, 0);")          
        #self.OutputText.setTextColor(QColor(255, 255, 255))
        #self.OutputText.setStyleSheet("background-color: rgb(0, 0, 0);") 
        
  
        
    def update(self, keyPressed):
        print("!!!!!!!!!!!!!!update")
        #newtext = str(self.le2.text()) + str(keyPressed)  #"tab pressed "
        #self.le2.setText(newtext)
         
         
    def Connect_click(self):
        print("Connect_click")
        if (ser.isOpen()==False):
            self.com_open()
        else:
            self.com_close()      
         
    def com_open(self):
        #ser.baudrate = 115200
        print("com_open")
        print("ser.port =",ser.port )   
        print("ser.baudrate =",ser.baudrate )      
        print("ser.parity =",ser.parity )   
        print("ser.bytesize =",ser.bytesize )       
        print("ser.stopbits =",ser.stopbits )   
       
        
        if (ser.port==None)|(ser.port==""):
            self.Configure_click()
        else:
            try:            
                ser.open()
                print("ser.open()")
                global port_open
                port_open = True
                print ("port_open = ",port_open)
                self.Connect.setStyleSheet("background-color: rgb(0, 255, 0);")
                self.Connect.setText("Connect")
                # 打开串口接收定时器，周期为2ms
                self.timer.start(2)
                #self.ReadUARTThread.start()
            except:
                 print("Can not open")
            
            

    def com_close(self):
        print("com_close")
        global port_open
        port_open = False
        self.timer.stop()
        ser.close()
        self.Connect.setStyleSheet("background-color: rgb(255, 0, 0);")  
        self.Connect.setText("Disconnect")
    
    def command_function(self):
        print ("port_open = ",port_open)
        if port_open == False:
            return

        # 获取到text光标
        textCursor = self.OutputText.textCursor()   
                     
        # 滚动到底部
        textCursor.movePosition(textCursor.End)
                    
        # 设置光标到text中去
        self.OutputText.setTextCursor(textCursor)
        
        input_s = self.command.text()

        self.command.clear()
        input_s = (input_s + '\n').encode('utf-8')
        print ("input_s =", input_s)        
        ser.write(input_s)

        #wig.close()
        
    def Configure_click(self): 
        print("Configure_click")      
        wig.comport_clear()
        wig.comport_scan()
        
        if (ser.isOpen()==True):
            ser.close()
          
        wig.setWindowModality(Qt.ApplicationModal) # lock sub window
        wig.show()
        
    def clear_click(self):    
        self.OutputText.clear()             
  
    def Timestamp_click(self):        
        ct = datetime.datetime.now()
        ct=ct.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        #date_time = ct.strftime("%Y-%m-%d, %H:%M:%S")
        #print("date and time:",date_time)
        
        formated_str = "[%s]"%(ct);
        print("current time: ", formated_str)   
        
        self.writeflie(formated_str)  
        self.writeflie(" ============== timestamp ================ ")            
        self.writeflie("\n") 
        '''
        global Timestamp_flag        
        if Timestamp_flag==True:
            Timestamp_flag = False
            self.Timestamp.setStyleSheet("background-color: rgb(255, 0, 0);")  
        else:
            Timestamp_flag = True
            self.Timestamp.setStyleSheet("background-color: rgb(0, 255, 0);")
        '''
        # 获取到text光标
        textCursor = self.OutputText.textCursor()   
                     
        # 滚动到底部
        textCursor.movePosition(textCursor.End)
                    
        # 设置光标到text中去
        self.OutputText.setTextCursor(textCursor)

        
                
    '''    
    def ReadUART(self):
        print("Threading...")
        while (True):#port_open ==
            try:
                #ch = ser.read().decode(encoding='ascii')
                data = ser.read()
                #self.s2__receive_text.insertPlainText(ch.decode('iso-8859-1'))
                
                #ch = ser.read()
                #print(data,end='')
                self.OutputText.insertPlainText(data.decode('iso-8859-1'))
                #self.OutputText.insertPlainText(data)
            except:
                print("Something wrong in receiving.")
    '''

    # 接收数据
    def data_receive(self):
        global first_line_character
        
        try:
            num = ser.inWaiting()
        except:
            self.com_close()
            return None
        if num > 0:
            data = ser.read(num)
            num = len(data)
            
            print(data,end='')   
            
            timestamp_now = datetime.datetime.now()
            timestamp_now=timestamp_now.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            timestamp_now_formated = "[%s] "%(timestamp_now);
                                 
            #ch = data.decode(encoding='ascii')
            ch = data.decode(encoding='ascii',errors='backslashreplace') # errors for arduino reset
            
        
            i=0;
            n=0;
            Outstr=""
            
            global Timestamp_flag
            
            for i in range(num):
                if first_line_character:
                    if Timestamp_flag==True:
                        Outstr = Outstr + timestamp_now_formated
                    first_line_character = False
                    
                if ch[n] == '\n':
                    Outstr=Outstr+ch[n]
                    first_line_character = True
                elif ch[n] == '\r':
                    pass
                else:
                    Outstr=Outstr+ch[n]

                n=n+1     
            
            
            #print(Outstr,end='')       
            self.writeflie(Outstr)            
            #print(ch,end='')             
            #print("type=",type(ch))
            #print(data,end='')        

            data2 = data.decode('iso-8859-1')
            data2 = data2.replace('\r','') 
            
            scrollbar = self.OutputText.verticalScrollBar()
            #print("scrollbar.maximum = ", scrollbar.maximum())   
            scrollbarAtBottom = (scrollbar.value() >= (scrollbar.maximum() - 4))
            #print("scrollbarAtBottom = ", scrollbarAtBottom)   
            scrollbarPrevValue = scrollbar.value()
            #print("scrollbarPrevValue = ", scrollbarPrevValue)   
            
            # 获取到text光标
            textCursor = self.OutputText.textCursor()   
            
            if (textCursor.atEnd()==False):          
                # 滚动到底部
                textCursor.movePosition(textCursor.End)
                    
                # 设置光标到text中去
                self.OutputText.setTextCursor(textCursor)
             
                     
            # 串口接收到的字符串为b'123',要转化成unicode字符串才能输出到窗口中去
            #self.OutputText.insertPlainText(data2.decode('iso-8859-1'))
            self.OutputText.insertPlainText(data2)
            #self.OutputText.append(data2)
            
            if (scrollbarAtBottom):
                self.OutputText.ensureCursorVisible()
            else:
                self.OutputText.verticalScrollBar().setValue(scrollbarPrevValue)                      
            
        else:
            pass

    def openflie(self):
        global fileName
        new_fileName, save = QFileDialog.getSaveFileName(self,
                  "檔案儲存",
                  "./",
                  "All Files (*);;Text Files (*.txt)")
                  
        print("fileName =",new_fileName)
        if (new_fileName==""):
            print("fileName is empty")
        else:
            fileName = new_fileName

    def writeflie(self,log):
        global fileName
        f = open(fileName,'a') 
        data=str(log)
        f.write(data)
        f.close()        

    # 视图-浏览器字体颜色设置
    def browser_word_color(self):
        col = QColorDialog.getColor(self.OutputText.textColor(), self, "文字顏色設定")
        print("col = ", col)
        if col.isValid():
            self.OutputText.setTextColor(col)

    # 视图-浏览器背景颜色设置
    def browser_background_color(self):
        col = QColorDialog.getColor(self.OutputText.textColor(), self, "背景颜色设置")
        if col.isValid():
            self.OutputText.setStyleSheet(
                "background-color: rgb({}, {}, {});".format(col.red(), col.green(), col.blue()))


    def eventFilter(self, obj, event):
        #print(obj, event.type())
        #print("event.key === ",event.key()) 
        '''
        if event.type() == QEvent.KeyPress:
            print("=======KeyPress===========") 
            if event.key()== Qt.Key_Tab:
                print("event = ",event.text()) 
                input_s = event.text()
                input_s = (input_s).encode('utf-8')            
                ser.write(input_s)
                return False
        '''    
        if event.type() == QEvent.KeyPress:
            #print("event.key pressed()=", event.key())
            return True
        
        if event.type() == QEvent.KeyRelease:            
            global port_open
            if port_open==False:
                return True
                         
            if (self.OutputText.hasFocus()==True):
                #print("event.key = ",event.key()) 
                # 获取到text光标
                textCursor = self.OutputText.textCursor()                            
                # 滚动到底部
                textCursor.movePosition(textCursor.End)                        
                # 设置光标到text中去
                self.OutputText.setTextCursor(textCursor)
                                   
                print("event.key()=", event.key())
                
                if event.key()==16777220: # Enter key
                    print("event.key()=", event.key())
                    #input_s = ""
                    input_s = ('\n').encode('utf-8')
                elif event.key()==16777249:  # ^C
                    print("event.key()=", event.key())
                    return True
                #elif event.key()==67:  # 
                #    print("event.key()=", event.key())
                #    return True                    
                else:
                    #key = eventQKeyEvent.key()

                    print("event = ",event.text()) 
                    print("event.key = ",event.key()) 
                    input_s = event.text()
                    input_s = (input_s).encode('utf-8')            
                
                ser.write(input_s)
                return True

        return False # will not be filtered out (ie. event will be processed)
        
        

    '''
    def event(self, event):     
        print("event.type()=")  
        #if port_open == False:
        #    return True
               
        if (event.type()==QEvent.KeyPress):
            if (self.OutputText.hasFocus()==True):       
                print("event = ",event.text()) 
                return True
        
        if (event.type()==QEvent.KeyRelease) and (event.key()==Qt.Key_Space):
            print("Key_Space")  
            
            
        #if (event.type()==QEvent.ShortcutOverride):
        #    print("event.key()",event.key())  
            

        #return  event(self, event)
    '''    
        
        
#if (self.OutputText.hasFocus()==True):       
    '''
    def keyReleaseEvent(self, event):#eventQKeyEvent
        global port_open
        if port_open==False:
            return
                     
        if (self.OutputText.hasFocus()==True):
            # 获取到text光标
            textCursor = self.OutputText.textCursor()                            
            # 滚动到底部
            textCursor.movePosition(textCursor.End)                        
            # 设置光标到text中去
            self.OutputText.setTextCursor(textCursor)
            
            if event.key()== Qt.Key_Tab: # Tab key
                print("=================Key_Tab=====================") 
                print("event = ", event.key()) 
                print("=================Key_Tab=====================")             
        
            if event.key()==16777220: # Enter key
                #input_s = ""
                input_s = ('\n').encode('utf-8')
            #elif event.key()== Qt.Key_Tab: # Tab key
            #    print("=================Key_Tab=====================") 
            #    print("event = ", event.key()) 
            #    print("=================Key_Tab=====================") 
            else:
                #key = eventQKeyEvent.key()

                print("event = ",event.text()) 
                input_s = event.text()
                input_s = (input_s).encode('utf-8')            
            
            ser.write(input_s)
    '''        
    

    
    '''
    def keyPressEvent(self, event):
        print("keyPressEvent")  
        
        print("OutputText.hasFocus = ",self.OutputText.hasFocus())    
        if (self.OutputText.hasFocus()==True):
            if event.key() == Qt.Key_Space:
                print("Key_Space")    
            else:
                print("event = ",event.text()) 
    '''                        
           
class Sub(QWidget,config.Ui_Configure):
    print("class Sub")    
    def __init__(self):
        print("Sub__init__")    
        super().__init__()
        self.setupUi(self)
           
        self.pushButton.clicked.connect(self.configureClose)
            
        self.baud_list = ["4800", "9600", "14400", "19200", "38400", "57600", "115200"] 
        
        self.BaudRate.addItems(self.baud_list)
        self.Parity_list = ["none", "odd", "even", "mark", "space"]
        self.Parity.addItems(self.Parity_list)        
        self.databit_list = ["5", "6", "7", "8"]
        self.databit.addItems(self.databit_list)     
        self.stopbit_list = ["1", "2"]
        self.stopbit.addItems(self.stopbit_list)   
        
        print("baud_index = ", self.baud_list.index(str(ser.baudrate))) 
      
        self.BaudRate.setCurrentIndex(self.baud_list.index(str(ser.baudrate)))  

        if (ser.parity == serial.PARITY_NONE):
            self.Parity.setCurrentIndex(0)
        elif(ser.parity == serial.PARITY_ODD):
            self.Parity.setCurrentIndex(1)
        elif(ser.parity == serial.PARITY_EVEN):
            self.Parity.setCurrentIndex(2)
        elif(ser.parity == serial.PARITY_MARK):
            self.Parity.setCurrentIndex(3)
        elif(ser.parity == serial.PARITY_SPACE):
            self.Parity.setCurrentIndex(4)
        
        self.databit.setCurrentIndex(ser.bytesize-5)        
        self.stopbit.setCurrentIndex(ser.stopbits-1)     
           
        
    def configureClose(self): 
           
        ser.port =  self.comport.currentText()
        print("ser.port = ",ser.port)  
        
        ser.baudrate = self.baud_list[self.BaudRate.currentIndex()]        
        
        parity = self.Parity.currentIndex()
        if (parity==0):
            ser.parity = serial.PARITY_NONE;
        elif(parity==1):
            ser.parity = serial.PARITY_ODD;
        elif(parity==2):
            ser.parity = serial.PARITY_EVEN;
        elif(parity==3):
            ser.parity = serial.PARITY_MARK;
        elif(parity==4):
            ser.parity = serial.PARITY_SPACE;
                
        databits = self.databit.currentIndex()
        if (databits==0):
            ser.bytesize = serial.FIVEBITS;
        elif(databits==1):
            ser.bytesize = serial.SIXBITS;
        elif(databits==2):
            ser.bytesize = serial.SEVENBITS;
        elif(databits==3):
            ser.bytesize = serial.EIGHTBITS;
            
        stopbits =self.stopbit.currentIndex()
        if (stopbits == 0):
            ser.stopbits = serial.STOPBITS_ONE;
        elif (stopbits == 1):
            ser.stopbits = serial.STOPBITS_TWO;        
        
        
        if (ser.isOpen()): # open success
            print("opened already")  
        else:
            window.com_open()

        config_w = configparser.ConfigParser()
        config_w.read('config.ini')
        
        config_w.set('SystemSettings','comport',ser.port)
        config_w.set('SystemSettings','baudrate',str(ser.baudrate))
        config_w.set('SystemSettings','parity',ser.parity)
        config_w.set('SystemSettings','databits',str(ser.bytesize))
        config_w.set('SystemSettings','stopbits',str(ser.stopbits))
        
        config_w.write(open('config.ini',"w+"))

    
        wig.close()
 
    def comport_clear(self): 
        self.comport.clear()

    def comport_scan(self): 
        print("comport_scan")    
        # 处理串口值
        port_list = list(serial.tools.list_ports.comports())  
        port_str_list = []  # 用来存储切割好的串口号
        for i in range(len(port_list)):
            # 将串口号切割出来
            lines = str(port_list[i])
            str_list = lines.split(" ")
            #print("comport_list === ", str_list[0]) 
            self.comport.addItem(str_list[0])  
            
        AllItems = [self.comport.itemText(i) for i in range(self.comport.count())]    
        try:
            port_index = AllItems.index(str(ser.port))
            self.comport.setCurrentIndex(port_index)
        except:
            print("comport is not in list")
            


if __name__ == '__main__':
    global port_open
    global Timestamp_flag
    global log_to_file_flag
    global first_line_character
    global fileName
    
    CURRENT_PACKAGE_DIRECTORY = os.path.abspath('.')    
    
    print (CURRENT_PACKAGE_DIRECTORY)
    
    Log_DIRECTORY = CURRENT_PACKAGE_DIRECTORY + '\\log' 
    
    if os.path.exists(Log_DIRECTORY):
        print("Log目錄已存在。")   	
    else:       
        # 使用 try 建立目錄
        try:
            os.makedirs(Log_DIRECTORY)
        # 檔案已存在的例外處理
        except FileExistsError:
            print("Log目錄已存在。")   
    
    
    now = datetime.datetime.now()
    fileName = now.strftime("%Y-%m")+".txt"   
    
    fileName=os.path.join(Log_DIRECTORY,fileName)
    
    print (fileName)
    
    port_open = False
    Timestamp_flag = True
    log_to_file_flag = False
    first_line_character = True
        
    port_list = None
    port_str_list = []  
        
    ser = serial.Serial()
     
    config = configparser.ConfigParser()       
    config_w = configparser.ConfigParser()    
    config_w['SystemSettings'] = {}        
    config.read('config.ini')
    
    try:    
        ser.port = config['SystemSettings']['comport'] 
    except:
        ser.port = "COM1"
        config_w['SystemSettings']['comport'] = "COM1"
        config_w.write(open('config.ini',"w+"))  
    
    try:
        ser.baudrate = int(config['SystemSettings']['baudrate'])
    except:
        ser.baudrate = 115200
        config_w['SystemSettings']['baudrate'] = "115200"  
        config_w.write(open('config.ini',"w+"))  
    
    try:
        ser.parity = config['SystemSettings']['parity']
    except:
        ser.parity = "N"
        config_w['SystemSettings']['parity'] = "N"
        config_w.write(open('config.ini',"w+"))  
        
    try:    
        ser.bytesize = int(config['SystemSettings']['databits'])
    except:
        ser.bytesize = 8
        config_w['SystemSettings']['databits'] = "8"
        config_w.write(open('config.ini',"w+"))  
    
    try:
        ser.stopbits = int(config['SystemSettings']['stopbits'])
    except:
        ser.stopbits = 1
        config_w['SystemSettings']['stopbits'] = "1"
        config_w.write(open('config.ini',"w+"))  
          
    
    print("ser.port =",ser.port )   
    print("ser.baudrate =",ser.baudrate )      
    print("ser.parity =",ser.parity )   
    print("ser.bytesize =",ser.bytesize )       
    print("ser.stopbits =",ser.stopbits ) 

    app = QtWidgets.QApplication(sys.argv)
    window = Main()
    wig = Sub()   
    window.show()


    sys.exit(app.exec_())

