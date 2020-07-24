import subprocess
import sys
import foxpro_extractor
import mysql_extractor

if __name__ == "__main__":
    print('installing packages...')
    packages = ["dbfread", "numpy","pandas", "python-dotenv", "xlrd", "xlwt"]
    for package in packages:
        print('installing ' + package)
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(package + ' installed')
    
    foxproData = foxpro_extractor.getExcel()

