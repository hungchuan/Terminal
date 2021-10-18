pyuic5 -x example.ui -o example_ui.py 

pyuic5 -x Terminal_ui3.ui -o Terminal_ui.py 

pyinstaller -F -w --onefile --icon=D:\Ryan\Project\Git\Terminal\QT\Icon.ico Terminal_main.py