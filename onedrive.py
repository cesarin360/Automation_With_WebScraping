from easy_one_drive import easy_one_drive as od
from pathlib import Path
from read_and_write import read_json as read

""" Actualiza el archivo en onedrive para empresas """

class onedrive():
    def upload_file():
        PATH = read('config.json', 'login')
        s = od.easy_one_drive(PATH['user'], PATH['password'], driver_path=PATH['driver'])
        try:
            if Path(PATH['pathd']).is_file():
                s.logging_in()
                s.delete_file(PATH['Folder'], PATH['Deletefile'], 20)  
                s.upload_file(PATH['pathd'], 30)
                s.logging_out()
                return True
            else:
                print("file not found")
                return False
        except Exception:
            print('Upload File Error', Exception)
            s.logging_out()
            s.driver.quit()
            return False