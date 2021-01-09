import sys
from PyQt5.QtWidgets import *
import pandas as pd
#from pandas import DataFrame
import numpy as np
import tabula
#import pdb

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Conversor PDF de relatório para EXCEL")
        self.setGeometry(350,150,500,400)
        self.UI()

    def UI(self):
        vbox=QVBoxLayout()
        hbox=QHBoxLayout()
        fileButton=QPushButton("Abrir arquivo")
        fileButton.clicked.connect(self.openFile)
        vbox.addLayout(hbox)
        hbox.addStretch()
        hbox.addWidget(fileButton)
        hbox.addStretch()
        self.setLayout(vbox)

        self.show()
    def openFile(self):
        url = QFileDialog.getOpenFileName(self,"Escolha o arquivo","","*pdf")
        fileUrl = url[0]
        if fileUrl:
            try:
                lista = tabula.read_pdf(fileUrl, guess=False, area = (51.0,0.0,701.0,601.0), pages = "all", columns= ( 80.0, 180,420.0,  605.0) )
                df = pd.DataFrame(columns=['Nr Documento', 'Nr Processo', 'Endereço', 'Dt Emissão'])
                for i in range(len(lista)):
                    df = df.append(lista[i], ignore_index=True)

                df = df.replace(np.nan, '')

                indexs = []
                for index, _ in df.iterrows():
                    if df['Nr Documento'].at[index] != '' and df['Nr Documento'].at[index] != '[RETIFICADA PE':
                        indexs.append(index)
                    else:
                        indexs.append(-1)

                df['obs'] = ''

                count = 0
                for i, _ in df.iterrows():
                    if i == 0:
                        pass
                    else:
                        count = count + 1

                    if count == indexs[i]:
                        u = True
                        index = i
                    elif count != indexs[i]:
                        u = False

                    if u == False and df['Nr Documento'].at[i] != '[RETIFICADA PE':
                        df['Endereço'].at[index] = str(df['Endereço'].at[index]) + ' ' + str(df['Endereço'].at[i])
                    elif u == False and df['Nr Documento'].at[i] == '[RETIFICADA PE':
                        df['obs'].at[index] = str(df['Nr Documento'].at[i]) + str(df['Nr Processo'].at[i]) + str(
                            df['Endereço'].at[i])

                df = df[df['Nr Documento'] != '']
                df = df[df['Nr Documento'] != '[RETIFICADA PE']

                url2 = QFileDialog.getSaveFileName(self, "Salvar", "", "*xlsx")
                saveUrl = url2[0]
                #print(saveUrl)
                if saveUrl:
                    if '.xlsx' in saveUrl:
                        writer = pd.ExcelWriter(saveUrl, engine='xlsxwriter')
                    else:
                        writer = pd.ExcelWriter(saveUrl+".xlsx", engine='xlsxwriter')

                    df.to_excel(writer, sheet_name='Relatório', index=False, na_rep='NaN')

                    for column in df:

                        column_length = max(df[column].astype(str).map(len).max(), len(column))
                        col_idx = df.columns.get_loc(column)
                        writer.sheets['Relatório'].set_column(col_idx, col_idx, column_length)

                    writer.save()
            except Exception as e:
                print(x)
                print(e)
                raise


def main():
    App=QApplication(sys.argv)
    window = Window()
    sys.exit(App.exec_())

if __name__=='__main__':
    main()