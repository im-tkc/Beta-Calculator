@echo off
pyinstaller src\run.py --hidden-import YahooToExcel --hidden-import PricePoint --onefile --paths=src/
