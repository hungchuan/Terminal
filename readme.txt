pyuic5 -x example.ui -o example_ui.py 

pyuic5 -x Terminal_ui3.ui -o Terminal_ui.py 

pyinstaller -F -w --onefile --icon=D:\Ryan\Project\Git\Terminal\Icon.ico Terminal_main.py
pyinstaller -D -w --icon=D:\Ryan\Project\Git\Terminal\Icon.ico Terminal_main.py

conda create --name py37 python=3.7

git remote set-url origin git@github.com:hungchuan/Terminal.git
