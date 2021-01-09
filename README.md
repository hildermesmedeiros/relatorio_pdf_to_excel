### PyQt App to convert a especif pdf file to excel table - app for Licenciamento SMIHC

# To compile it I used pyinstaller, ie:
>pyinstaller --add-data "D:\Anaconda\envs\QTenv\Lib\site-packages\tabula\tabula-1.0.4-jar-with-dependencies.jar;tabula" --onedir .\Relatorio.py

## It depends on:
* PyQt5
* Pandas
* Tabula
