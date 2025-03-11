pyuic5 -x example.ui -o example_ui.py 

	

C:\Users\Ryan\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\Scripts\pyuic5 -x Terminal_ui3.ui -o Terminal_ui3.py 
C:\Users\Ryan\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\Scripts\pyuic5 -x configure.ui -o configure.py 

pyinstaller -F -w --onefile --icon=D:\Ryan\Project\Git\Terminal\Icon.ico Terminal_main.py
pyinstaller -D -w --icon=D:\Work\Tools\MODBUS\Terminal\Icon.ico Terminal_main.py
pyinstaller -F -w --icon=D:\Work\Tools\MODBUS\Terminal\Icon.ico Terminal_main.py
pyinstaller -w --icon=D:\Work\Tools\MODBUS\Terminal\Icon.ico Terminal_main.py

conda create --name py37 python=3.7

git remote set-url origin git@github.com:hungchuan/Terminal.git
