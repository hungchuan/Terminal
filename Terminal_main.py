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
import time
from PyQt5.QtCore import QDateTime
import argparse
import binascii
from collections import defaultdict
import ast

CONFIG_FILE = "config.ini"

DEFAULT_CONFIG = {
    "SystemSettings": {
        "comport": "COM3",
        "baudrate": "19200",
        "parity": "E",
        "databits": "8",
        "stopbits": "1",
        "modbus": "rtu",
        "tabwidgetindex": "0"
    },
    "ProgramSettings": {
        "program_start_addr": "0x84008",
        "last_file": "",
        "modbus_id": "81"
    }
}


Maxlines = 10
MaxlinesInputed = 0
SelectedCol = 100
SelectedRow = 100
MB_ids = ['81']

class Main(QWidget, ui.Ui_MainWindow):
    print("class Main")    
    def __init__(self):
        global modbus_mode
        global cmd_format
        global Maxlines
        global Program_start_addr
        global TabWidgetIndex
        global response_data_hex
        global MB_ids
        global chunk_size

        parser = argparse.ArgumentParser(description="Example script with debug option")
        parser.add_argument("--debug", action="store_true", help="Enable debug mode")
        args = parser.parse_args()
        
        # setup the debug mode 
        self.debug_mode = args.debug  
        if self.debug_mode:
            print("=====  debug on =====")
        else:            
            print("=====  debug off =====")
        
        self.Progfile_path = None
        self.n=0
        self.bin_data = None
        self.bin_data_Bank0 = None
        self.bin_data_Bank1 = None
        
        self.target_start_addr = Program_start_addr #0x84008  # 設定 MCU 的目標地址
        
        if (modbus_mode == "ascii"):
            chunk_size = 256
        else:
            chunk_size = 128        
        
        
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
        self.ProgOpen.clicked.connect(self.ProgOpenFile)
        self.ProgUpgrade.clicked.connect(self.ProgStart)
        self.Scan.clicked.connect(self.scan_modbus_ids)
        
        #self.progress_bar = QtWidgets.QProgressBar()
        #self.progressBar.setProperty("value", 25)
        
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
         
        # MB_ID list init 
        self.MB_ID.addItems(MB_ids) 
        self.MB_ID.currentIndexChanged.connect(self.MBID_Change)
        self.MB_ID.setCurrentIndex(int(MB_list_sel))  
        
        print("baudrate = ",ser.baudrate)   
        self.com_open()
        self.OutputText.setReadOnly(False)
        
        self.OutputText.setStyleSheet("background-color: rgb(0, 0, 0);""color: rgb(255, 255, 255);")  
        self.ProgOutputText.setStyleSheet("background-color: rgb(0, 0, 0);""color: rgb(255, 255, 255);")  
        
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
        self.ProgUpgrade.setEnabled(False)
        self.ProgCancel.setEnabled(False)
        self.tabWidget.setCurrentIndex(int(TabWidgetIndex))
        print("currentIndex = ",self.tabWidget.currentIndex())
        
        self.ProgFileName.setText(Program_last_file)
        
        

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
        if port_open == False:
            return
            
    def RegReadAll_click(self):
        print("RegReadAll_click")
        if port_open == False:
            return
        input_s = "a 3a00 100"    
        input_s = (input_s + '\n').encode('utf-8')       
        ser.write(input_s)
        
    def RegWriteAll_click(self):
        print("RegWriteAll_click")        
        if port_open == False:
            return
            
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
            input_s = self.command.text()
            # 如果 input_s 為空則直接 return
            if not input_s:
                self.command.clear()
                return            
            hex_values = input_s.replace(' ', '').split()
            # 只保留有效的十六進制數字 (過濾掉無效輸入)
            valid_hex_values = [h for h in hex_values if all(c in "0123456789ABCDEFabcdef" for c in h)]

            if not valid_hex_values:
                print("Error: No valid hexadecimal numbers found.")
                self.OutputText.append("Error: No valid hexadecimal numbers found.")
                self.command.clear()
                return
            else:
                try:
                    bytes_data = bytes.fromhex(''.join(valid_hex_values))
                    self.command.clear()
                    # Send Modbus
                    self.send_modbus_request(ser, bytes_data)
                    input_s = (input_s + '\n').encode('utf-8')
                except ValueError as e:
                    self.command.clear()
                    print(f"Error converting to bytes: {e}")
                    self.OutputText.append(f"Error converting to bytes: {e}")
                    return

            
            #bytes_data = bytes.fromhex(''.join(hex_values))
            #self.command.clear()
            # 发送Modbus请求
            #self.send_modbus_request(ser, bytes_data)
            #input_s = (input_s + '\n').encode('utf-8')    
        
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
        if (0==self.tabWidget.currentIndex()):        
            self.OutputText.clear()
        elif (1==self.tabWidget.currentIndex()):  
            self.ProgOutputText.clear()
  
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
        global response_data_hex
        
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
                print('data_receive:response_data_hex :', response_data_hex) 
            else:
                self.OutputText.insertPlainText(data2)     
                        
            #self.OutputText.append(data2)
            
            if (scrollbarAtBottom):
                self.OutputText.ensureCursorVisible()
            else:
                self.OutputText.verticalScrollBar().setValue(scrollbarPrevValue)                      
            
        else:
            pass
            
            
    def ProgOpenFile(self):
        global Program_last_file
        # Open file dialog to select BIN or HEX file
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        initial_dir = os.path.dirname(Program_last_file) if Program_last_file else ""
        Progfile_path, _ = QFileDialog.getOpenFileName(self, "Select Firmware File", initial_dir, "Binary/Hex Files (*.bin *.hex)", options=options)

        if Progfile_path:
            Program_last_file = Progfile_path
            if Program_last_file.endswith(".hex"):
                print("It's a HEX file")
                
                # Parse HEX file and store data in bin_data_bank0 and bin_data_bank1
                self.bin_data_bank0, self.bin_data_bank1 = self.parse_hex_file(Program_last_file)
                
                # Calculate and print checksum
                checksum = sum(self.bin_data_bank0) & 0xFF  
                print("checksum0 =", hex(checksum))
                checksum = sum(self.bin_data_bank1) & 0xFF  
                print("checksum1 =", hex(checksum))

                # **Check if bin_data_bank0 is empty**
                if self.bin_data_bank0[:4] == b'\xFF\xFF\xFF\xFF':
                    print("bin_data_bank0 is not available")
                    self.bin_data = self.bin_data_bank1
                else:
                    self.bin_data = self.bin_data_bank0
                '''
                # Print debug message
                # Get the first 64 bytes and last 64 bytes
                first_64_bytes = self.bin_data_bank0[:64]
                last_64_bytes = self.bin_data_bank0[-64:]
                    
                # Print first 64 bytes and last 64 bytes in hexadecimal format
                print("bank0 First 64 bytes:", first_64_bytes.hex(" "))
                print("bank0 Last 64 bytes:", last_64_bytes.hex(" "))

                # Get the first 64 bytes and last 64 bytes
                first_64_bytes = self.bin_data_bank1[:64]
                last_64_bytes = self.bin_data_bank1[-64:]
                    
                # Print first 64 bytes and last 64 bytes in hexadecimal format
                print("bank1 First 64 bytes:", first_64_bytes.hex(" "))
                print("bank1 Last 64 bytes:", last_64_bytes.hex(" "))
                '''

            else:
                # Read binary file
                with open(Progfile_path, "rb") as file:
                    self.bin_data = file.read()
                    '''
                    # Print debug message
                    # Get the first 64 bytes and last 64 bytes
                    first_64_bytes = self.bin_data[:64]
                    last_64_bytes = self.bin_data[-64:]
                    
                    # Print first 64 bytes and last 64 bytes in hexadecimal format
                    print("First 64 bytes:", first_64_bytes.hex(" "))
                    print("Last 64 bytes:", last_64_bytes.hex(" "))
                    '''

                    # Calculate and print checksum
                    checksum = sum(self.bin_data) & 0xFF  
                    print("checksum =", hex(checksum))                    
            
            self.Progfile_path = Progfile_path
            self.ProgFileName.setText(f"{Progfile_path}")
            self.ProgUpgrade.setEnabled(True)
            set_config_value(config, "ProgramSettings", "last_file", Program_last_file)
            save_config(config)

        else:
            # If no file is selected, disable the upgrade button
            self.ProgFileName.setText("")
            self.ProgUpgrade.setEnabled(False)
            
            
    def parse_hex_file(self, hex_file):
        """ Parse Intel HEX file and convert it to binary data (address * 2, 16-bit Little-Endian) """
        
        # Memory mapping
        memory = defaultdict(lambda: b'\xFF\xFF')  # Default fill with 0xFFFF
        base_address = 0  # Extended address (Segment / Linear Address)
        
        # Target address range (multiplied by 2)
        BANK0_START = 0x086808 * 2
        BANK1_START = 0x0A6808 * 2

        # Read HEX file
        with open(hex_file, "r") as f:
            for line in f:
                if not line.startswith(":"):
                    continue  # HEX records should start with `:`
                
                line = line.strip()
                byte_count = int(line[1:3], 16)
                address = int(line[3:7], 16) * 2 + base_address  # Multiply address by 2
                record_type = int(line[7:9], 16)
                data = line[9:9 + byte_count * 2]
                
                if record_type == 0x00:  # Data record
                    raw_data = binascii.unhexlify(data)

                    # **Ensure 16-bit Little-Endian storage**
                    for i in range(0, len(raw_data), 2):
                        chunk = raw_data[i:i+2]
                        if len(chunk) < 2:
                            chunk += b'\xFF'  # Ensure 16-bit length
                        memory[address + i] = chunk[::-1]  # Little-Endian byte order

                elif record_type == 0x02:  # Extended Segment Address Record
                    base_address = (int(data, 16) << 4) * 2

                elif record_type == 0x04:  # Extended Linear Address Record
                    base_address = (int(data, 16) << 16) * 2

                elif record_type == 0x01:  # End of File Record
                    break  # Stop parsing
        
        # Get the EndAddress
        if memory:
            EndAddress = max(memory.keys())  # Get the last byte's address
        else:
            EndAddress = BANK0_START  # If HEX file is empty, set to BANK0_START

        # Calculate File_Length
        if EndAddress > BANK1_START:
            File_Length = EndAddress - BANK1_START + 2
        else:
            File_Length = EndAddress - BANK0_START + 2

        # Set correct bank range
        BANK0_END = BANK0_START + File_Length
        BANK1_END = BANK1_START + File_Length

        # Initialize bin_data_bank0 & bin_data_bank1
        bin_data_bank0 = bytearray([0xFF] * (BANK0_END - BANK0_START))
        bin_data_bank1 = bytearray([0xFF] * (BANK1_END - BANK1_START))

        # Fill in data
        for addr in range(BANK0_START, BANK0_END, 2):
            if addr in memory:
                bin_data_bank0[addr - BANK0_START: addr - BANK0_START + 2] = memory[addr]

        for addr in range(BANK1_START, BANK1_END, 2):
            if addr in memory:
                bin_data_bank1[addr - BANK1_START: addr - BANK1_START + 2] = memory[addr]

        # Save bin files for verify 
        '''
        with open("bin_data_bank0.bin", "wb") as f:
            f.write(bin_data_bank0)

        with open("bin_data_bank1.bin", "wb") as f:
            f.write(bin_data_bank1)
        '''
        
        return bytes(bin_data_bank0), bytes(bin_data_bank1)
                  
    def ProgStart_ascii(self):
        global chunk_size
        print("ProgStart_ascii")        
        # 記錄開始時間
        start_time = QTime.currentTime()  
             
        """開始燒錄流程"""
        if not self.bin_data:
            QMessageBox.warning(self, "No File", "Please load a BIN file first.")
            return
            
        if port_open == False:
            QMessageBox.warning(self, "COM error", "Please check COM port.")
            return        
                   
        command = ('$ Download mode. Must jump bootloader.\r').encode('utf-8')
        self.ProgOutputText.append(f"$ Download mode. Must jump bootloader.")   
        ser.write(command)
        
        if not self.wait_for_response("BIOS", timeout=2):
            QMessageBox.warning(self, "Upload Failed", "Failed to enter bootloader mode.")
            return
            
        # 傳送準備命令
        checksum = sum(self.bin_data) & 0xFF  # 計算 Checksum
        size = len(self.bin_data)
        command = f"pprog {self.target_start_addr:05X} {size:05X} {(self.target_start_addr + size) & 0xFFFFF:05X}\r"
        self.ProgOutputText.append(command) 
        command = command.encode()           
        print("command =",command)

        
        ser.write(command)

        if not self.wait_for_response("a", timeout=5):
            QMessageBox.warning(self, "Upload Failed", "MCU did not acknowledge start command.")
            return            

        self.Connect.setEnabled(False)  #Disable connect button
        self.Configure.setEnabled(False)  #Disable Configure button
        
        
        # 傳送資料
        self.progressBar.setValue(0)
        bytes_sent = 0
        checksum = 0;
        while bytes_sent < size:
            chunk = self.bin_data[bytes_sent:bytes_sent + chunk_size]
            checksum = (sum(chunk) + checksum) & 0xFF  # 計算 Checksum
            
            chunk = chunk.ljust(256, b'\x00') #補滿到 256 bytes

            ser.write(chunk)
                 
            if not self.wait_for_checksum(checksum, timeout=5):
                QMessageBox.warning(self, "Upload Failed", "Checksum mismatch.")
                self.Connect.setEnabled(True)  #Enable connect button
                self.Configure.setEnabled(True)  #Disable Configure button                   
                return
            
            bytes_sent += len(chunk)
            #bytes_sent = size
            self.progressBar.setValue(int((bytes_sent / size) * 100))

            QApplication.processEvents()  # update UI
        
        # 記錄結束時間並顯示
        end_time = QTime.currentTime()
        self.ProgOutputText.append(f"start time: {start_time.toString('HH:mm:ss')}")          
        self.ProgOutputText.append(f"end time: {end_time.toString('HH:mm:ss')}")
        # 計算執行時間
        elapsed_time = start_time.msecsTo(end_time) / 1000  # Millisecond to Second
        self.ProgOutputText.append(f"elapsed time: {elapsed_time:.2f} seconds")       
        
        QMessageBox.information(self, "Upload Complete", "Firmware upload completed successfully.")
        self.Connect.setEnabled(True)  #Enable connect button
        self.Configure.setEnabled(True)  #Disable Configure button        
            
       
    def ProgStart_RTU(self):
        print("ProgStart_RTU ")
        global response_data_hex
        response_data_hex = None   
        # 記錄開始時間
        start_time = QTime.currentTime()  
             
        """開始燒錄流程"""
        if not self.bin_data:
            QMessageBox.warning(self, "No File", "Please load a BIN file first.")
            return
            
        if port_open == False:
            QMessageBox.warning(self, "COM error", "Please check COM port.")
            return        

        if self.debug_mode:
            self.writeflie(" ============== ProgStart_RTU ================ ")            
            self.writeflie("\n") 
        
        # Step 1: Send Firmware update start (0x81)
        firmware_update_cmd = ModbusID_HEX + b'\x6E\x81\x10WIC-001LF100LA' + b'\x00' * (16 - len("WIC-001LF100LA"))
        response = self.send_modbus_request(ser, firmware_update_cmd)              
        
        # 2. read the response
        response = self.read_modbus_response(expected_length = 6,timeout=2)
               
        # Step 3: Check response       
        if not response or response[3] == 1:
            QMessageBox.warning(self, "Upload Failed", "0x81:Command not acceptable.")
            return
        time.sleep(3)

        # Step 4: Send Update Status (0x84)
        update_status_cmd = ModbusID_HEX + b'\x6E\x84'
        response = self.send_modbus_request(ser, update_status_cmd)
        
        # 5. read the response
        response = self.read_modbus_response(expected_length = 7,timeout=2) 
        # Step 6: Check response      
        if not response or (response[4] & 0x40) == 0:
            QMessageBox.warning(self, "Upload Failed", "0x84:Update not allowed.")
            return
  
        # Check executing bank
        if not response or (response[4] & 0x80) == 0:
            self.writeflie("\n") 
            self.writeflie(" ============== Executing at Bank0, send bank1 ================ ")            
            self.writeflie("\n")             
            self.bin_data = self.bin_data_bank1 # Executing at Bank0, send bank1 
        else:
            self.writeflie("\n") 
            self.writeflie(" ============== Executing at Bank1, send bank0 ================ ")            
            self.writeflie("\n")              
            self.bin_data = self.bin_data_bank0 # Executing at Bank1, send bank0 
 
        # 傳送準備命令
        checksum = sum(self.bin_data) & 0xFF  # 計算 Checksum
        size = len(self.bin_data)
 
        self.Connect.setEnabled(False)  #Disable connect button
        self.Configure.setEnabled(False)  #Disable Configure button
        
        """==================================================================="""
        
        # 傳送資料
        self.progressBar.setValue(0)
        bytes_sent = 0
        checksum = 0
        block_index = 0
        last_block = 0
        
        while bytes_sent < size:
            chunk = self.bin_data[bytes_sent:bytes_sent + chunk_size]
            checksum = (sum(chunk) + checksum) & 0xFF  # 計算 Checksum
            
            chunk = chunk.ljust(128, b'\x00') #補滿到 128 bytes
            
            if ((bytes_sent + chunk_size) > size):
                last_block = 1
            
            #request = bytes([0x51, 0x6E, 0x82, 0x85, 0x00, 0x00, last_block, block_index & 0xFF, (block_index >> 8) & 0xFF]) + chunk
            request = ModbusID_HEX + bytes([0x6E, 0x82, 0x85, 0x00, 0x00, last_block,block_index & 0xFF, (block_index >> 8) & 0xFF]) + chunk
            response = self.send_modbus_request(ser, request)         
            #read the response
            response = self.read_modbus_response(expected_length = 6,timeout=2)
                   
            #Check response       
            if not response or response[3] == 1:
                QMessageBox.warning(self, "Upload Failed", "0x82:Command not acceptable.")
                self.Connect.setEnabled(True)  #Enable connect button
                self.Configure.setEnabled(True)  #Disable Configure button  
                return
            
            bytes_sent += len(chunk)
            block_index += 1
            #bytes_sent = size
            self.progressBar.setValue(int((bytes_sent / size) * 100))

            QApplication.processEvents()  # 更新 UI
        
        # Switch to new firmware 
        firmware_update_cmd = ModbusID_HEX + b'\x6E\x83\x08\x07'
        response = self.send_modbus_request(ser, firmware_update_cmd)          
        
        # 記錄結束時間並顯示
        end_time = QTime.currentTime()
        self.ProgOutputText.append(f"start time: {start_time.toString('HH:mm:ss')}")          
        self.ProgOutputText.append(f"end time: {end_time.toString('HH:mm:ss')}")
        # 計算執行時間
        elapsed_time = start_time.msecsTo(end_time) / 1000  # 毫秒轉換為秒
        self.ProgOutputText.append(f"elapsed time: {elapsed_time:.2f} seconds")       
        
        QMessageBox.information(self, "Upload Complete", "Firmware upload completed successfully.")
        self.Connect.setEnabled(True)  #Enable connect button
        self.Configure.setEnabled(True)  #Disable Configure button   
        
        if self.debug_mode:
            self.writeflie(" ============== ProgStart_RTU End ================ ")            
            self.writeflie("\n") 

    def ProgStart(self):
        print ("modbus_mode =", modbus_mode)
        if (modbus_mode == "ascii"):
            self.ProgStart_ascii();
        else:
            self.ProgStart_RTU();

    def read_modbus_response(self, expected_length, timeout=5):
        """等待指定回應字串，超時返回 False"""
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            num = ser.inWaiting()
            print ("num =", num)
            if (num>=expected_length):
                response = ser.read(num)
                print ("response[0] =", response[0])
                print ("response[1] =", response[1])
                print ("response[2] =", response[2])
                print ("response[3] =", response[3])
                print ("response[4] =", response[4])
                print ("response[5] =", response[5])
                #print ("response[6] =", response[6])
                
                if self.debug_mode:
                    response_data_hex = ' '.join(['{:02X}'.format(byte) for byte in response])
                    self.ProgOutputText.append(response_data_hex)
                    # Print to file for debug        
                    self.writeflie(response_data_hex)
                    self.writeflie("\n")

                return response
            time.sleep(0.01)
        return None
  
    def wait_for_response(self, expected_response, timeout=5):
        """等待指定回應字串，超時返回 False"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            response = ser.read(1024).decode(errors='ignore')
            print('response :', (expected_response, response))
            if expected_response in response:
                return True
            time.sleep(0.1)
        return False
        
    def wait_for_checksum(self, expected_checksum, timeout=5):
        """等待單一 byte 的 checksum 驗證，設置超時"""
        start_time = time.time()

        while time.time() - start_time < timeout:
            byte = ser.read(1)  # 每次讀取 1 byte
            if byte:
                received_checksum = byte[0]  # 取得接收的 byte 值
                print('checksum :', (hex(expected_checksum), hex(received_checksum)))
                return received_checksum == expected_checksum

            time.sleep(0.01)  # 短暫延遲，避免過多 CPU 使用

        print("wait_for_checksum timeout ")
        return False  # 超時未收到 checksum
           
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
        global TabWidgetIndex
        print("TabWidgetIndex:", TabWidgetIndex)
        # 计算CRC
        crc = self.calculate_crc(request_data)
        print("RTU CRC:", hex(crc))
        crc_bytes = struct.pack('<H', crc)
        
        print("write:", request_data + crc_bytes)
        #self.ProgOutputText.insertPlainText(request_data + crc_bytes)  
        
        if self.debug_mode:
            # 將字節串轉換為十六進制字符串
            request_hex = binascii.hexlify(request_data + crc_bytes).decode('utf-8')
            # 印出要傳送的資料到 ProgOutputText 或是 OutputText
            request_str = " ".join(request_hex[i:i+2] for i in range(0, len(request_hex), 2))
            if (0==int(TabWidgetIndex)):
                self.OutputText.append(request_str)
            elif (1==int(TabWidgetIndex)):
                self.ProgOutputText.append(request_str)
                
            # Print to file for debug        
            self.writeflie(request_str)
            self.writeflie("\n")         

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
        MBtextCursor = self.ProgOutputText.textCursor()   
                     
        # 滚动到底部
        MBtextCursor.movePosition(MBtextCursor.End)
                    
        # 设置光标到text中去
        self.ProgOutputText.setTextCursor(MBtextCursor)
        
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
        global TabWidgetIndex
        print("Switched to tab:", index)
        #config_w['SystemSettings']['TabWidgetIndex'] = str(index)
        #config_w.write(open('config.ini',"w+"))
        set_config_value(config, "SystemSettings", "TabWidgetIndex", str(index))
        save_config(config)
        
        TabWidgetIndex = index
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
            
    def MBID_Change(self, index):
        global ModbusID_HEX
        print("MBID_Change:", index)
        print(type(ModbusID_HEX)) 
        ModbusID_HEX = int(MB_ids[index]).to_bytes(1, 'big')  # Transfer to Hex        
        set_config_value(config, "ProgramSettings", "modbus_list_sel", str(index))
        save_config(config)        

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

    def scan_modbus_ids(self, scan_timeout=0.04):
        global TabWidgetIndex
        global ModbusID_HEX
        global MB_ids
               
        found_ids = []
        for modbus_id in range(1, 256):
            request = bytes([modbus_id]) + b'\x6E\x84'
            response = self.send_modbus_request(ser, request)            
            response = self.read_modbus_response(expected_length = 7,timeout = 0.04)
	            
            if response and response[0] == modbus_id:
                print("Found = ",modbus_id)  
                found_ids.append(str(modbus_id))
                if (0==int(TabWidgetIndex)):
                    self.OutputText.append(f"Find ID {modbus_id}")    
                elif (1==int(TabWidgetIndex)):
                    self.ProgOutputText.append(f"Find ID {modbus_id}")                  
            
            else:
                if (0==int(TabWidgetIndex)):
                    self.OutputText.append(f"ID {modbus_id} not found")    
                elif (1==int(TabWidgetIndex)):
                    self.ProgOutputText.append(f"ID {modbus_id} not found")
            
            QApplication.processEvents() #update UI
            #found_ids.append(str(modbus_id))
        
        if not found_ids:
            print("Not found, use default ID 81") 
            found_ids.append("81")  # default is 81        
        print("found_ids = ",found_ids) 
        self.MB_ID.clear()
        self.MB_ID.addItems(found_ids)
        set_config_value(config, "ProgramSettings", "found_ids", found_ids)
        self.MB_ID.setCurrentIndex(0)  
        
        MB_ids = found_ids
        #MB_ids = ast.literal_eval(MB_ids)
        
        #ModbusID_HEX = int(found_ids[0]).to_bytes(1, 'big')  # Transfer to Hex
        ModbusID_HEX = int(MB_ids[0]).to_bytes(1, 'big')  # Transfer to Hex        
        
        return found_ids
    
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
        global chunk_size
        
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
            ser.timeout = 0.01
            chunk_size = 256
        elif (modbus_index == 1):
            modbus_mode = "rtu" 
            ser.timeout = 0.1
            chunk_size = 128
        
        if (ser.isOpen()): # open success
            print("opened already")  
        else:
            window.com_open()      
        
        set_config_value(config, "SystemSettings", "comport", ser.port)
        set_config_value(config, "SystemSettings", "baudrate", str(ser.baudrate))
        set_config_value(config, "SystemSettings", "parity", ser.parity)
        set_config_value(config, "SystemSettings", "databits", str(ser.bytesize))
        set_config_value(config, "SystemSettings", "stopbits", str(ser.stopbits))
        set_config_value(config, "SystemSettings", "modbus", str(modbus_mode))
        save_config(config)      
       
        wig.close()
 
    def comport_clear(self): 
        self.comport.clear()

    def comport_scan(self): 
        print("comport_scan")    
        # 处理串口值
        port_list = list(serial.tools.list_ports.comports())  

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
           


def load_config():
    """讀取或創建 config.ini 配置文件"""
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE):
        config.read_dict(DEFAULT_CONFIG)
        save_config(config)  # save
    else:
        config.read(CONFIG_FILE)
    return config

def get_config_value(config, section, key, default):
    """嘗試讀取配置，若不存在則使用預設值，並更新 config"""
    if not config.has_section(section):
        config[section] = {}

    if key not in config[section]:
        config[section][key] = default  # load default
        save_config(config)  
    return config[section][key]

def set_config_value(config, section, key, value):
    """設定 config 內的某個值並寫回 config.ini"""
    if not config.has_section(section):
        config[section] = {}
    config[section][key] = str(value)  # Make sure all data is stored in string format
    save_config(config)

def save_config(config):
    #save config.ini
    with open(CONFIG_FILE, "w") as f:
        config.write(f)
        
        
if __name__ == '__main__':
    global port_open
    global Timestamp_flag
    global log_to_file_flag
    global first_line_character
    global fileName
    global modbus_mode
    global cmd_format
    global Debug_mode
    global response_data_hex
    global ModbusID_HEX

    
    Debug_mode = "off"
    CURRENT_PACKAGE_DIRECTORY = os.path.abspath('.')    
    
    print (CURRENT_PACKAGE_DIRECTORY)
    
    Log_DIRECTORY = CURRENT_PACKAGE_DIRECTORY + '\\log' 
    
    if os.path.exists(Log_DIRECTORY):
        print("Log folder exist")   	
    else:       
        try:
            os.makedirs(Log_DIRECTORY)
        except FileExistsError:
            print("Log folder exist")     
    
    
    now = datetime.datetime.now()
    fileName = now.strftime("%Y-%m")+".txt"   
    
    fileName=os.path.join(Log_DIRECTORY,fileName)
    
    print (fileName)
    
    port_open = False
    Timestamp_flag = True
    log_to_file_flag = False
    first_line_character = True
        
    port_list = None
        
    ser = serial.Serial()
    
    config = load_config()

    # Read SystemSettings
    ser.port = get_config_value(config, "SystemSettings", "comport", "COM1")
    ser.baudrate = int(get_config_value(config, "SystemSettings", "baudrate", "115200"))
    ser.parity = get_config_value(config, "SystemSettings", "parity", "N")
    ser.bytesize = int(get_config_value(config, "SystemSettings", "databits", "8"))
    ser.stopbits = int(get_config_value(config, "SystemSettings", "stopbits", "1"))
    modbus_mode = get_config_value(config, "SystemSettings", "modbus", "ascii")
    TabWidgetIndex = int(get_config_value(config, "SystemSettings", "tabwidgetindex", "0"))

    # Read ProgramSettings
    Program_start_addr = get_config_value(config, "ProgramSettings", "program_start_addr", "0x84008")
    Program_last_file = get_config_value(config, "ProgramSettings", "last_file", "")
    #ModbusID = get_config_value(config, "ProgramSettings", "modbus_id", "81")
    MB_list_sel = get_config_value(config, "ProgramSettings", "modbus_list_sel", "0")
    MB_ids = get_config_value(config, "ProgramSettings", "found_ids", "['81']")
    MB_ids = ast.literal_eval(MB_ids)
       
    # Save config
    save_config(config)      
    
    ModbusID_HEX = int(MB_ids[int(MB_list_sel)]).to_bytes(1, 'big')  # Transfer to Hex
      

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
    print("Program_start_addr =",Program_start_addr )           
    
    #ser = serial.Serial('COM10', 19200, parity=serial.PARITY_EVEN, bytesize=8, stopbits=1, timeout=1)

    app = QtWidgets.QApplication(sys.argv)
    window = Main()
    wig = Sub()   
    window.show()


    sys.exit(app.exec_())

