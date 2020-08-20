import subprocess
import sys
import foxpro_extractor
import mysql_injector
import input_to_template

def install_packages():
    print('installing packages...')
    packages = ["dbfread", "numpy","pandas", "python-dotenv", "xlrd", "xlwt","xlutils"]
    for package in packages:
        print('installing ' + package)
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(package + ' installed')

if __name__ == "__main__":
    # install_packages()
    
    foxproData = foxpro_extractor.getExcel()

    # mysqldata = mysql_injector.master_peserta()
