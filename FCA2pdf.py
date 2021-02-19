import pathlib
import openpyxl
import PyPDF2  # 外部ライブラリ
import os
from win32com import client  # 外部ライブラリ

"""
シートごとにpdfファイルを作成する.
これ以降のpdf統合コードをクオートで囲み，Acrobat Pro DCなどで統合を行うことで，
統合後に自動で型番のしおりが付く．
全三行程をばらばらに書いたのはこのため．
"""
path_CR_input = str(input('コストレポートの最上階層のフォルダ名を入力:'))
path_SS_input = str(input('実行するシステムアセンブリのフォルダ名を入力:'))
path_SS = pathlib.Path(os.path.join('..' ,path_CR_input, path_SS_input))

xlApp = client.Dispatch('Excel.Application')  # PythonでVBAを利用するためのコード
all_list = []
for folder_obj in path_SS.iterdir():
    for pass_obj in folder_obj.iterdir():
        if pass_obj.match('*.xlsx'):
            try:
                fca = xlApp.workbooks.open(str(pass_obj.resolve()))
                """
                VBAを用いてWorkbookを開く．
                workbooks.open()関数は引数に絶対パスを指定する必要がある.
                resolveメソッドは絶対パスを返す．
                開いたworkbookオブジェクトは変数fcaに代入する

                """
                no_list = []
                for sheet in fca.Worksheets:
                    if sheet.Range('A5').value == 'P/N Base':
                    # Partsシートの場合
                        try:
                            comp_no = str(int(sheet.Range('B5').value))
                            # 型番を取得
                        except:
                            # Int型に変換出来なかった場合
                            print('以下のパーツでP/N Baseに誤り')
                            print(sheet.Range('B4').value)
                    elif sheet.Range('A4').value == 'P/N Base':
                    # Assemblyシートの場合
                        comp_no = str(sheet.Range('B4').value)
                        # 型番を取得
                    no_list.append(comp_no)
                    file_name = comp_no + '.pdf'
                    # ファイル名は型番.pdfとしている
                    pdf_path = pathlib.Path(folder_obj) / file_name
                    try:
                        sheet.ExportAsFixedFormat(0, str(pdf_path.resolve()))
                    except:
                        pass
                fca.Close()
                all_list.append(no_list)
            except:
                print(pass_obj + 'にてエラー発生')

xlApp.Quit()

"""
シートごとに作成したpdfを統合する.
ただし，PyPDF2はしおりの生成が(おそらく)出来ないので，
型番によるしおりを自動生成したい場合はこれ以下をクオートで囲み，
統合はAcrobat Pro DCなどで行うとよい
"""

counter = 0
for folder_obj in path_SS.iterdir():
    merger = PyPDF2.PdfFileMerger()
    for i in range(len(all_list[counter])):
        file_no = str(all_list[counter][i])
        file_name_check = str(file_no + '.pdf')
        for pass_obj in folder_obj.iterdir():
            if pass_obj.match(file_name_check):
                merger_path = pathlib.Path(pass_obj)
                merger.append(str(merger_path.resolve()))
                d = PyPDF2.PdfFileReader(str(merger_path.resolve())).documentInfo
                merger.addMetadata(d)
    counter += 1
    merger.write(str(pathlib.Path(os.path.join(folder_obj, 'FCA統合(名前を変更).pdf'))))
    merger.close()


"""
最後に型番.pdfの形で出力したファイルを消去している．
残したい場合は以下をクオートで囲む．
"""
counter = 0
for folder_obj in path_SS.iterdir():
    for pass_obj in folder_obj.iterdir():
        if pass_obj.match('*.pdf'):
            for i in range(len(all_list[counter])):
                file_no = str(all_list[counter][i])
                file_name_check = str(file_no + '.pdf')
                if pass_obj.match(file_name_check):
                    try:
                        fca_pdf_path = pathlib.Path(pass_obj)
                        os.remove(str(fca_pdf_path.resolve()))
                    except:
                        pass
    counter += 1
