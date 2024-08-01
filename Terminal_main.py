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
import struct
import binascii
import inspect
from PyQt5 import QtCore
import pandas as pd


Maxlines = 10
MaxlinesInputed = 0
SelectedCol = 100
SelectedRow = 100


class Main(QWidget, ui.Ui_MainWindow):
    print("class Main")    
    def __init__(self):
        global modbus_mode
        global cmd_format
        global Maxlines

        
        print("Main__init__")   
        super().__init__()
        self.setupUi(self)
        self.Connect.setAutoFillBackground(True)
        self.Debug.setAutoFillBackground(True)

        self.Configure.clicked.connect(self.Configure_click)

         
        self.Connect.clicked.connect(self.Connect_click)
        self.Clear.clicked.connect(self.clear_click)
        self.Timestamp.clicked.connect(self.Timestamp_click)
        self.log.clicked.connect(self.openflie)
        #self.command.returnPressed.connect(self.command_function) #Press Enter callback 
        self.Debug.clicked.connect(self.Debug_click)
        #self.MBCMD.returnPressed.connect(self.MBCMD_function) #Press Enter callback 
        self.RegReadOne.clicked.connect(self.RegReadOne_click)
        self.RegReadAll.clicked.connect(self.RegReadAll_click)
        self.RegWriteAll.clicked.connect(self.RegWriteAll_click)
        
        self.tabWidget.currentChanged.connect(self.on_tab_changed)
        
        
        self.Connect.setStyleSheet("background-color: rgb(255, 0, 0);")  
        self.Timestamp.setStyleSheet("background-color: rgb(0, 255, 0);")  
        self.OutputText.setFont(QFont('Consolas', 9))
        self.setWindowIcon(QIcon(':img/Icon.ico')) #若ICON在同一層目錄
        
        self.Configure.setFocusPolicy(Qt.ClickFocus)#NoFocus   ClickFocus
        self.Clear.setFocusPolicy(Qt.ClickFocus)
        self.Timestamp.setFocusPolicy(Qt.ClickFocus)
        self.log.setFocusPolicy(Qt.ClickFocus)
        self.Connect.setFocusPolicy(Qt.ClickFocus)
        self.Debug.setFocusPolicy(Qt.ClickFocus)
        self.OutputText.setFocusPolicy(Qt.ClickFocus)
        self.command.setFocusPolicy(Qt.ClickFocus)        
        self.frame.setFocusPolicy(Qt.ClickFocus)         
        self.setFocusPolicy(Qt.ClickFocus)       
        
        self.OutputText.installEventFilter(self)
        self.command.installEventFilter(self)
      
        # 定时器接收数据
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.data_receive)      
         
        #self.ReadUARTThread = threading.Thread(target=self.ReadUART)
        #self.ReadUARTThread.start()
          
        #self.le  = MyLineEdit()
        
        #self.le.signalTabPressed[str].connect(self.update)
        
        self.CmdType.currentIndexChanged.connect(self.CmdTypeChange)
        self.CmdType_list = ["ASCII", "HEX"]
        self.CmdType.addItems(self.CmdType_list)   
        if (modbus_mode == "ascii"):
            cmd_format = "ASCII"
            self.CmdType.setCurrentIndex(0)
        else:
            cmd_format = "HEX"
            self.CmdType.setCurrentIndex(1)       
         
        
        print("baudrate = ",ser.baudrate)   
        self.com_open()
        self.OutputText.setReadOnly(False)
        
        self.OutputText.setStyleSheet("background-color: rgb(0, 0, 0);""color: rgb(255, 255, 255);")  
        self.MBOutputText.setStyleSheet("background-color: rgb(0, 0, 0);""color: rgb(255, 255, 255);")  
        
        self.printlinefilefunc()
        
        self.lines = [""] * Maxlines  # 初始化一個包含50個空字串的列表
        self.current_index = 0  # 初始化當前索引   
        self.print_index = 0  # 初始化當前索引
        

        # 读取Device.xlsx中的Device分頁
        self.loadDevices()  
        # 连接 Device 选择变化的信号到处理函数
        self.Device.currentIndexChanged.connect(self.load_functions)        
        
        # 在Reg標籤頁添加16x16表格
        self.tableWidget = QTableWidget(16, 16, self.tab_3)
        self.tableWidget.setGeometry(QtCore.QRect(10, 100, 500, 500))  # 調整表格位置和大小
        self.tableWidget.setObjectName("tableWidget")
             
        # 設置表格的標題
        self.tableWidget.setHorizontalHeaderLabels([f'0x{i:X}' for i in range(16)])
        self.tableWidget.setVerticalHeaderLabels([f'0x{i*16:02X}' for i in range(16)])
        

        # 新增 1x16 的表格
        self.smallTableWidget = QTableWidget(1, 16, self.tab_3)
        self.smallTableWidget.setGeometry(QtCore.QRect(10, 550, 700, 30))  # 调整表格位置和大小
        # 设置表格的标题（bit15 到 bit0）
        self.smallTableWidget.setHorizontalHeaderLabels([f'bit{i}' for i in range(15, -1, -1)])
        self.smallTableWidget.setVerticalHeaderLabels(['Data'])  # 显示行标题       
        
        # 添加布局以使表格随窗口缩放
        self.layout = QVBoxLayout(self.tab_3)
        self.layout.addWidget(self.label)
        self.layout.addWidget(self.tableWidget)
        self.layout.addWidget(self.smallTableWidget)
        
        # 设置伸缩因子以调整高度比例
        self.layout.setStretch(0, 3)  # label占比1
        self.layout.setStretch(1, 30)  # tableWidget占比9
        self.layout.setStretch(2, 3)  # smallTableWidget占比1      
        
        self.label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.tableWidget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.smallTableWidget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        self.tab_3.setLayout(self.layout)     

        # 調整每一列的寬度為40
        for i in range(16):
            self.tableWidget.setColumnWidth(i, 40)
            self.tableWidget.horizontalHeader().setSectionResizeMode(i,QHeaderView.Fixed); #fixed width 
            self.smallTableWidget.setColumnWidth(i, 40)
            self.smallTableWidget.horizontalHeader().setSectionResizeMode(i,QHeaderView.Fixed); #fixed width 

        # 設置內容置中tableWidget
        for row in range(16):
            for col in range(16):
                item = self.tableWidget.item(row, col) or QTableWidgetItem("")
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.tableWidget.setItem(row, col, item)

        # 設置內容置中smallTableWidget
        for col in range(16):
            item = self.smallTableWidget.item(0, col) or QTableWidgetItem("")
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.smallTableWidget.setItem(0, col, item)
                
                
        # Set the last two rows as read-only and grey them out
        for row in range(14, 16):
            for col in range(16):
                item = QTableWidgetItem('')
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item.setBackground(Qt.gray)
                self.tableWidget.setItem(row, col, item)
            
        self.resize(800, 700)  # 窗口的宽度和高度
        
        # Connect signals to slots
        self.tableWidget.currentCellChanged.connect(self.cell_moved)
        #self.tableWidget.cellClicked.connect(self.cell_clicked)   
        

    def cell_clicked(self, row, col):
        print(f'Clicked cell at row {row + 1}, column {col + 1}')

    def cell_moved(self, currentRow, currentCol, previousRow, previousCol):
        global SelectedCol
        global SelectedRow
        if currentRow is not None and currentCol is not None:
            SelectedCol = currentCol
            SelectedRow = currentRow
            print(f'Moved to cell at row {SelectedRow + 1}, column {SelectedCol + 1}')
            
    def printlinefilefunc(self):
        callerframerecord = inspect.stack()[1]
        frame = callerframerecord[0]
        info = inspect.getframeinfo(frame)
        print("info = ",info.filename,info.function,info.lineno )
        print("Function = ",info.function)
        print("Line = ",info.lineno)
        
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

    def Debug_click(self):
        global Debug_mode
        print("Debug_click")
        if port_open == False:
            return        
        if (Debug_mode=="off"):
            Debug_mode = "on"
            input_s = "$ DEBUG ON"
            self.Debug.setText("DEBUG ON")
            self.Debug.setStyleSheet("background-color: rgb(255, 0, 0);")  
        else:
            Debug_mode = "off"
            input_s = "$ DEBUG OFF"
            self.Debug.setText("DEBUG OFF")
            self.Debug.setStyleSheet("background-color: rgb(200, 200, 200);")  
        input_s = (input_s + '\n').encode('utf-8')       
        ser.write(input_s) 

    def RegReadOne_click(self):
        print("RegReadOne_click")

    def RegReadAll_click(self):
        print("RegReadAll_click")
        
    def RegWriteAll_click(self):
        print("RegWriteAll_click")        
            
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
        self.Debug.setStyleSheet("background-color: rgb(200, 200, 200);")  
        self.Debug.setText("DEBUG OFF")        
    
    def command_function(self):
        global modbus_mode
        global Maxlines
        global MaxlinesInputed
        print ("port_open = ",port_open)
        
        if port_open == False:
            return

        # 获取到text光标
        textCursor = self.OutputText.textCursor()   
                     
        # 滚动到底部
        textCursor.movePosition(textCursor.End)
                    
        # 设置光标到text中去
        self.OutputText.setTextCursor(textCursor)
        
        print ("modbus_mode =", modbus_mode)
        if (modbus_mode == "ascii"):
            input_s = self.command.text()
            self.command.clear()
            input_s = (input_s + '\n').encode('utf-8')
            print ("input_s =", input_s)        
            ser.write(input_s)
        else:
            # 获取用户输入的十六进制数据
            input_s = self.command.text()
            hex_values = input_s.replace(' ', '').split()
            bytes_data = bytes.fromhex(''.join(hex_values))
            self.command.clear()
            # 发送Modbus请求
            self.send_modbus_request(ser, bytes_data)
            input_s = (input_s + '\n').encode('utf-8')    
        
        input_s_new = input_s[:-1]
        #存储到lines数组中               
        self.lines[self.current_index] = input_s_new.decode('utf-8')  
        self.current_index = (self.current_index + 1) % Maxlines  # 保证索引在0到49之间循环     
        self.print_index = self.current_index

        MaxlinesInputed = MaxlinesInputed+1
        if (MaxlinesInputed>Maxlines):
            MaxlinesInputed = Maxlines+1      
        
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
        global modbus_mode
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
            #print("data_receive=",data)
            #current_index = self.tabWidget.currentIndex()

            if (modbus_mode == "rtu"):                                
                response_data_hex = ' '.join(['{:02X}'.format(byte) for byte in data])
                #self.OutputText.append(response_data_hex)
            
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
            if (modbus_mode == "rtu"):
                self.OutputText.append(response_data_hex)
            else:
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
        global Maxlines
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
        if (event.type() == QEvent.KeyPress and obj is self.command):
            print('key press in command:', (event.key(), event.text()))
            print('print_index:', self.print_index)
            
            if (event.key()==16777220) | (event.key()==16777221): # Enter key
                self.command_function()
            elif event.key()==16777235:  # up
                if (self.print_index>0):
                    self.print_index = self.print_index-1
                else:
                    if (MaxlinesInputed>=Maxlines):
                        self.print_index = Maxlines-1
                    else:
                        self.print_index = self.current_index-1
                print('print_index up:', self.print_index)  
                self.command.setText(self.lines[self.print_index])  
            elif event.key()==16777237:  # down
                
                print('print_index dw:', self.print_index)  
                print('current_index dw:', self.current_index)  
                if (MaxlinesInputed>=Maxlines):
                    self.print_index = (self.print_index + 1) % Maxlines
                else:
                    if (self.print_index<self.current_index):
                        self.print_index = self.print_index+1
                    
                self.command.setText(self.lines[self.print_index])             
                
            return False
        elif (event.type() == QEvent.KeyPress and obj is self.OutputText):
            print('key press in OutputText:', (event.key(), event.text()))
            
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
                    print("event.key()= Enter", event.key())
                    #input_s = ""
                    #input_s = ('\n').encode('utf-8')
                    input_s = ('\r').encode('utf-8')
                    
                elif event.key()==16777249:  # ^C
                    print("event.key()=", event.key())
                    return True
                elif event.key()==16777235:  # up
                    '''
                    # 获取到text光标
                    textCursor = self.OutputText.textCursor()   
                                 
                    # 滚动到底部
                    textCursor.movePosition(textCursor.End)
                    # 移动光标到行首
                    self.OutputText.moveCursor(QTextCursor.StartOfLine, QTextCursor.MoveAnchor)
                    # 重新设置值
                    self.OutputText.insertPlainText("")   
                    textCursor.removeSelectedText()
                    self.OutputText.setTextCursor(textCursor)      
                    '''

                    
                    cursor = self.OutputText.textCursor()
                    self.OutputText.moveCursor(QTextCursor.StartOfLine)
                    cursor.movePosition(QTextCursor.End)                    
                    cursor.movePosition(QTextCursor.StartOfLine, QTextCursor.KeepAnchor, 1)
                    cursor.removeSelectedText()
                    self.OutputText.setTextCursor(cursor)    
                    
                    #self.OutputText.moveCursor(QTextCursor.PreviousBlock, QTextCursor.MoveAnchor)
                    #self.OutputText.setTextCursor(textCursor)  
                     

                    #return True       
                    input_s = ( bytes.fromhex('1b5b41').decode('utf-8')).encode('utf-8')       
                elif event.key()==16777237:  # down
                    cursor = self.OutputText.textCursor()
                    self.OutputText.moveCursor(QTextCursor.StartOfLine)
                    cursor.movePosition(QTextCursor.End)                    
                    cursor.movePosition(QTextCursor.StartOfLine, QTextCursor.KeepAnchor, 1)
                    cursor.removeSelectedText()
                    self.OutputText.setTextCursor(textCursor)    
                    
                    input_s = ( bytes.fromhex('1b5b42').decode('utf-8')).encode('utf-8')       
                elif event.key()==16777234:  # left
                    input_s = ( bytes.fromhex('1b5b44').decode('utf-8')).encode('utf-8')       
                elif event.key()==16777236:  # right
                    input_s = ( bytes.fromhex('1b5b44').decode('utf-8')).encode('utf-8')           
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
        
    def calculate_crc(self,data):
        crc = 0xFFFF
        for byte in data:
            crc ^= byte
            for _ in range(8):
                if crc & 0x0001:
                    crc >>= 1
                    crc ^= 0xA001
                else:
                    crc >>= 1
        return crc        

    def send_modbus_request(self,serial_port, request_data):

        # 计算CRC
        crc = self.calculate_crc(request_data)
        print("RTU CRC:", hex(crc))
        crc_bytes = struct.pack('<H', crc)
        
        print("write:", request_data + crc_bytes)
        #self.MBOutputText.insertPlainText(request_data + crc_bytes)  
        
        # 將字節串轉換為十六進制字符串
        request_hex = binascii.hexlify(request_data + crc_bytes).decode('utf-8')

        # 印出要傳送的資料到 MBOutputText
        request_str = " ".join(request_hex[i:i+2] for i in range(0, len(request_hex), 2))
        self.OutputText.append(request_str)        

        self.timer.start(50)
        # 发送请求数据
        serial_port.write(request_data + crc_bytes)
        

    def receive_modbus_response(self,serial_port):
        response_received = threading.Event()

        def read_response():
            nonlocal response_received
            response_data = serial_port.read(1024)
            if response_data:
                print("response：", response_data)
                response_data_hex = ' '.join(['{:02X}'.format(byte) for byte in response_data])
                self.OutputText.append(response_data_hex)
            else:
                print("no response 1")

            response_received.set()

        response_thread = threading.Thread(target=read_response)
        response_thread.start()

        # 等待響應，最多等待1秒
        response_received.wait(timeout=1)

        if not response_received.is_set():
            print("no response 2")
        
    def MBCMD_function(self):
        #ser2 = serial.Serial('COM10', 19200, parity=serial.PARITY_EVEN, bytesize=8, stopbits=1, timeout=1)

        print ("port_open = ",port_open)
        if port_open == False:
            return 
        
        # 获取到text光标
        MBtextCursor = self.MBOutputText.textCursor()   
                     
        # 滚动到底部
        MBtextCursor.movePosition(MBtextCursor.End)
                    
        # 设置光标到text中去
        self.MBOutputText.setTextCursor(MBtextCursor)
        
        # 获取用户输入的十六进制数据
        input_text = self.MBCMD.text()
        hex_values = input_text.replace(' ', '').split()
        bytes_data = bytes.fromhex(''.join(hex_values))

        #self.MBCMD.clear()
        
        # 发送Modbus请求
        self.send_modbus_request(ser, bytes_data)

        # 接收Modbus响应
        #self.receive_modbus_response(ser)        
        
    def on_tab_changed(self, index):
        print("Switched to tab:", index)
        '''
        self.com_close()
        if (0==index):
            ser.timeout = 0.01
            self.timer.start(2)
        elif (1==index):
            ser.timeout = 0.1
        else:    
            ser.timeout = 0.01
            self.timer.start(2)
        self.com_open()
        '''
        
    def CmdTypeChange(self, index):
        global cmd_format
        print("Switched to index:", index)        
        if (0==index):
            cmd_format = "ASCII"
        elif (1==index):
            cmd_format = "HEX"

    def loadDevices(self):
        # 读取Device.xlsx中的Device分頁
        df = pd.read_excel('Device.xlsx', sheet_name='Device')
        
        # 检查列名
        if 'Device' in df.columns:
            devices = df['Device'].dropna().tolist()
        else:
            print('Device 列未找到')
        
        self.Device.addItems(devices)
        self.Device.setCurrentIndex(0)
        self.load_functions()

    def load_functions(self):
        # 获取当前选中的设备名称
        selected_device = self.Device.currentText()

        if selected_device:
            # 读取相应设备分页的内容
            df = pd.read_excel('Device.xlsx', sheet_name=selected_device)

            # 获取A2开始的内容
            functions = df.iloc[:, 0].dropna().tolist()
            self.Function.clear()
            self.Function.addItems(functions)        
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
        global modbus_mode
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
        self.modbus_list = ["ascii", "rtu"]
        self.ModbusType.addItems(self.modbus_list)          
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
        
        if (modbus_mode == "ascii"):
            self.ModbusType.setCurrentIndex(0) 
        else:
            self.ModbusType.setCurrentIndex(1)     
           
        
    def configureClose(self): 
        global modbus_mode
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
        
        modbus_index = self.ModbusType.currentIndex()
        if (modbus_index == 0):
            modbus_mode = "ascii"
        elif (modbus_index == 1):
            modbus_mode = "rtu" 
        
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
        config_w.set('SystemSettings','modbus',str(modbus_mode))

        
        config_w.write(open('config.ini',"w+"))
        
        if (modbus_mode == "ascii"):
            ser.timeout = 0.01
        else:
            ser.timeout = 0.1
        
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
    global modbus_mode
    global cmd_format
    global Debug_mode
    
    Debug_mode = "off"
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
        
    if (ser.port==None)|(ser.port==""):
        ser.port = "COM1"
    
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

    try:
        modbus_mode = config['SystemSettings']['modbus'] 
    except:
        modbus_mode = "ascii"
        config_w['SystemSettings']['modbus'] = "ascii"
        config_w.write(open('config.ini',"w+"))  

    if (modbus_mode == "ascii"):
        ser.timeout = 0.01
    else:
        ser.timeout = 0.1
    
    print("ser.port =",ser.port )   
    print("ser.baudrate =",ser.baudrate )      
    print("ser.parity =",ser.parity )   
    print("ser.bytesize =",ser.bytesize )       
    print("ser.stopbits =",ser.stopbits ) 
    print("modbus_mode =",modbus_mode ) 
    
    
    #ser = serial.Serial('COM10', 19200, parity=serial.PARITY_EVEN, bytesize=8, stopbits=1, timeout=1)

    app = QtWidgets.QApplication(sys.argv)
    window = Main()
    wig = Sub()   
    window.show()


    sys.exit(app.exec_())

