import os


UserLogin = os.getlogin()
if UserLogin == "":
       path = ''
else:
       path = 'C:\\Users\\' + UserLogin + '\\Downloads\\'
