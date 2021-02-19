"""
一行当たりの文字数がPEP8を無視しています．申し訳ないです．
読みにくかったら，改行してください．
PEP8では一行当たりの文字数は79文字以下が推奨されています．
また，インデントもTabキーを使用しています．
(本モジュールはVisual Studioで編集しています．)
"""

import openpyxl  # 外部ライブラリ
import pathlib
import os

path_CR_input = str(input('コストレポートの最上階層のフォルダ名を入力:'))
path_SS_input = str(input('実行するシステムアセンブリのフォルダ名を入力:'))
path_SS = pathlib.Path(os.path.join('..' ,path_CR_input, path_SS_input))
# システムアセンブリのパスを生成

car_number = int(input('車番を入力(半角2桁):'))

for folder_obj in path_SS.iterdir():
    for pass_obj in folder_obj.iterdir():  
    # システムアセンブリフォルダ内のファイルをワイルドカード検索
        if pass_obj.match('*.xlsx'):
        # 検索したファイルの内，Excelファイルのみを抽出
            fca = openpyxl.load_workbook(pass_obj)
            sheet_name_list = fca.get_sheet_names()
            # get_sheet_names関数は使えなくなる可能性あり
            for i, sh in enumerate(fca):
                if sh['B1'].value != 'Kyoto University':
                    print(sheet_name_list[i] + 'のUniversityの欄に誤りがあります．')
                if sh['A5'].value == 'Suffix' and sh['B5'].value != 'AA':
                    print(sheet_name_list[i] + 'のSuffixの欄に誤りがあります．')
                if sh['A6'].value == 'Suffix' and sh['B6'].value != 'AA':
                    print(sheet_name_list[i] + 'のSuffixの欄に誤りがあります．')
                if sh['K1'].value != car_number:
                    print(sheet_name_list[i] + 'のCarNumberの欄に誤りがあります．')

                # 以下sub_totalのSUM関数が全ての合計となっているか，ItemOrderが1から順に並んでいるかの確認
                # 同じような処理が続くので関数にすることも考えたが，引数が多くなるのでやめました
                for m in range(1, 500):
                    if sh['B'+str(m)].value == 'Material':
                        for j in range(m+1, 500):
                            if sh['A' + str(j)].value == None and sh['B' + str(j)].value == None:
                                if sh['N' + str(j)].value != '=SUM(N' + str(m+1) + ':N'+ str(j-1) + ')':
                                    print(sheet_name_list[i] + 'のMaterialのsub_totalの関数に誤りがあります．')
                                break
                            elif sh['A' + str(j)].value != 'MA' + str(j-m):
                                print(sheet_name_list[i] + 'のMaterialのItemOrderに誤りがあります．')
                        break
                
                for m in range(1, 500):
                    if sh['B' + str(m)].value == 'Process':
                        for j in range(m+1, 500):
                            if sh['A' + str(j)].value == None and sh['B' + str(j)].value == None:
                                if sh['I' + str(j)].value != '=SUM(I' + str(m+1) + ':I'+ str(j-1) + ')':
                                    print(sheet_name_list[i] + 'のProcessのsub_totalの関数に誤りがあります．')
                                break
                            elif sh['A' + str(j)].value != 'PR' + str(j-m):
                                print(sheet_name_list[i] + 'のProcessのItemOrderに誤りがあります．')
                        break
                
                for m in range(1, 500):
                    if sh['B' + str(m)].value == 'Fastener':
                        for j in range(m+1, 500):
                            if sh['A' + str(j)].value == None and sh['B' + str(j)].value == None:
                                if sh['J' + str(j)].value != '=SUM(J' + str(m+1) + ':J'+ str(j-1) + ')':
                                    print(sheet_name_list[i] + 'のFastenerのsub_totalの関数に誤りがあります．')
                                break
                            elif sh['A' + str(j)].value != 'FA' + str(j-m):
                                print(sheet_name_list[i] + 'のFastenerのItemOrderに誤りがあります．')
                        break

                for m in range(1, 500):
                    if sh['B' + str(m)].value == 'Tooling':
                        for j in range(m+1, 500):
                            if sh['A' + str(j)].value == None and sh['B' + str(j)].value == None:
                                if sh['I' + str(j)].value != '=SUM(I' + str(m+1) + ':I'+ str(j-1) + ')':
                                    print(sheet_name_list[i] + 'のToolingのsub_totalの関数に誤りがあります．')
                                break
                            elif sh['A' + str(j)].value != 'TO' + str(j-m):
                                print(sheet_name_list[i] + 'のToolingのItemOrderに誤りがあります．')
                        break
                
                # Multiplaierの確認，本モジュールではコストテーブルを参照しない
                for m in range(1, 500):
                    if sh['G' + str(m)].value == 'Multiplier':
                        for j in range(m+1, 500):
                            if sh['G' + str(j)].value == None and sh['B' + str(j)].value == None:
                                break
                            elif sh['G' + str(j)].value == 'Aluminum' and sh['H' + str(j)].value != 1:
                                print(sheet_name_list[i] + 'のMultiplier(Aluminum)に誤りがあります．')
                            elif sh['G' + str(j)].value == 'Steel' and sh['H' + str(j)].value != 3:
                                print(sheet_name_list[i] + 'のMultiplier(Steel)に誤りがあります．')
                            elif sh['G' + str(j)].value == 'Cast Iron' and sh['H' + str(j)].value != 2.50:
                                print(sheet_name_list[i] + 'のMultiplier(Cast Iron)に誤りがあります．')
                            elif sh['G' + str(j)].value == 'Foam' and sh['H' + str(j)].value != 0.33:
                                print(sheet_name_list[i] + 'のMultiplier(Foam)に誤りがあります．')
                            elif sh['G' + str(j)].value == 'Composite' and sh['H' + str(j)].value != 2:
                                print(sheet_name_list[i] + 'のMultiplier(Composite)に誤りがあります．')
                            elif sh['G' + str(j)].value == 'Plastic' and sh['H' + str(j)].value != 0.50:
                                print(sheet_name_list[i] + 'のMultiplier(Plastic)に誤りがあります．')
                            elif sh['G' + str(j)].value == 'Stainless Steel' and sh['H' + str(j)].value != 3.75:
                                print(sheet_name_list[i] + 'のMultiplier(Stainless Steel)に誤りがあります．')
                            if sh['E' + str(j)].value == 'cm^4':
                                print(sheet_name_list[i] + 'のProcessの単位に誤りがあります．')
                            if sh['B' + str(j)].value == 'Machining Setup, Install and remove' and sh['D' + str(j)].value != 1.30:
                                print(sheet_name_list[i] + 'のMachining Setup, Install and removeの金額に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Machining Setup, Change' and sh['D' + str(j)].value != 0.65:
                                print(sheet_name_list[i] + 'のMachining Setup, Changeの金額に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Laser Cut' and sh['D' + str(j)].value != 0.01:
                                print(sheet_name_list[i] + 'のLaser Cutの金額に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Machining' and sh['D' + str(j)].value != 0.04:
                                print(sheet_name_list[i] + 'のMachiningの金額に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Weld' and sh['D' + str(j)].value != 0.15:
                                print(sheet_name_list[i] + 'のWeldの金額に誤りがあります．')
                        break
                
                # Materialの確認，本モジュールではコストテーブルを参照しない
                for m in range(1, 500):
                    if sh['B' + str(m)].value == 'Material':
                        for j in range(m+1, 500):
                            if sh['B' + str(j)].value == None and sh['A' + str(j)].value == None:
                                break
                            elif sh['B' + str(j)].value == 'Aluminum, Premium' and sh['D' + str(j)].value != 4.2:
                                print(sheet_name_list[i] + 'のMaterial(Aluminum, Premiumの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Aluminum, Normal' and sh['D' + str(j)].value != 4.2:
                                print(sheet_name_list[i] + 'のMaterial(Aluminum, Normalの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Iron' and sh['D' + str(j)].value != 1:
                                print(sheet_name_list[i] + 'のMaterial(Ironの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Plastic, Nylon' and sh['D' + str(j)].value != 3.3:
                                print(sheet_name_list[i] + 'のMaterial(Plastic, Nylonの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Rubber' and sh['D' + str(j)].value != 3.3:
                                print(sheet_name_list[i] + 'のMaterial(Rubberの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Steel, Alloy' and sh['D' + str(j)].value != 2.25:
                                print(sheet_name_list[i] + 'のMaterial(Steel, Alloyの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Carbon Fiber, 1 Ply' and sh['D' + str(j)].value != 200:
                                print(sheet_name_list[i] + 'のMaterial(Carbon Fiber, 1 Plyの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Steel, Mild' and sh['D' + str(j)].value != 2.25:
                                print(sheet_name_list[i] + 'のMaterial(Steel, Mildの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Steel, Stainless' and sh['D' + str(j)].value != 2.25:
                                print(sheet_name_list[i] + 'のMaterial(Steel, Stainlessの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Honeycomb, Aluminum' and sh['D' + str(j)].value != 50:
                                print(sheet_name_list[i] + 'のMaterial(Honeycomb, Aluminumの金額)に誤りがあります．')
                            elif sh['B' + str(j)].value == 'Honeycomb, Nomex' and sh['D' + str(j)].value != 125:
                                print(sheet_name_list[i] + 'のMaterial(Honeycomb, Nomexの金額)に誤りがあります．')
                    
