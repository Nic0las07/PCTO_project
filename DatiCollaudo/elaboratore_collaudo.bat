@echo off
REM Esegui gli script nell’ordine richiesto

python src\csv_formatter.py
python src\xlsx_data_loader.py

pause