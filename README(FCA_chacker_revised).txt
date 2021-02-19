FCA_chacker_revised
(ver. Python 3.8.4)

FCAとコストテーブルを照らし合わせて，誤りがあった場合，リストアップします．
※FCAチェッカーの簡易版と組み合わせて使って下さい．

## Description
2020年度の電子提出版コストレポートのディレクトリ構造に合わせて作成したBOM作成プログラムです．
Python3.8.4, Openpyxl3.03にて動作します．
実行すると以下の入力が求められます．
・最上階層のフォルダ名：FCAとBOMがまとめて入っているフォルダです
・実行するシステムアセンブリのフォルダ名
・CostTable5種類のファイル名

##VS. FCA_chacker_light
本プログラムは, FCA_chacker_lightと異なり，コストテーブルを参照します．
そのため，ほぼ全てのコストテーブルの値を用いる項目についてチェックが行えます．
ただし, FCA_chacker_lightと異なり，大学名，車番等の確認も行えません．

## Requirement
・Python3の導入が必須です
・外部ライブラリopenpyxlを使用しています．導入が必須です．
　コマンドプロンプトを起動し，"pip install openpyxl"と入力するだけでインストールできます．

##Usage
python_prgのフォルダごと，コストレポートの最上階層と同じ階層にコピーしてください．
同時にCost_TablesというフォルダにCostTable5種類を格納してください．
example)
--------015_KyotoUniversity_FSAEJ_CR---------015_A1_BR
 　|                                                         |-----015_A2_EN
　 |                                                         |-----015_A3_FR
    |                                                         |            ：
　 |　　　　　　　　　　　　　　　　　　　　　   |-----015_KyotoUniversity_FSAEJ_CR_BOM.xlsx
　 |
    |---python_prg-------------------------------------FCA_chacker_revised.py
    |                                                          |------README(FCA_chacker_revised).txt
    |---Cost_Tables-------------------------------------Cost_Table 5種類
次にIDLEを起動し，File→OpenからFCA_chacker_revised.pyを開いてください．
(プログラムを動かすのみであればパワーシェルから開いても可能ですが，コードの修正はできません)
最後に，F5キーを押すと実行されます．

## Check_list
以下の項目をチェックします．
・MaterialのMatrial名とUnitCostの誤り
・ProcessのProcess名とUnitCost，Unitの誤り
・ProcessのMultiplier名とMultiVal．の誤り
・FastenerのFastener名とUnitCostの誤り
・ToolingのTooling名とUnitCostの誤り

## Attention
・2020年度のCostTableに沿って動作します．
・最初に入力した内容以外は2020年度フォーマットに対して動作します．
　フォーマットの変更によるFCAのsub_totalの列の変更，BOMの空欄部の列の変更などで動作しなくなります．
　ただし，フォーマットの多少の変更であれば，コードを数文字修正するのみで正しく動作します．
　各コードで何を行っているかは#にてコード中に注釈してあるので，修正してみてください．
・ディレクトリ構造はフォルダ名以外は2020年度の構造(以下)を厳密に維持してください．
　最上階層→BOMと各システムアセンブリのFCAおよび裏付け資料を含むフォルダ→サブシステムのフォルダ→FCAと裏付け資料
　システムアセンブリのフォルダ名を入力すると，そのフォルダ内に含まれるサブシステムのフォルダを全てリスト化します．
　さらにリスト内のサブシステムフォルダに存在する拡張子が.xlsxのファイルを順に全て開いて読み取っていきます．
　このため，フォルダ内にFCA以外のExcelファイルが存在すると誤作動が起きます．
・openpyxlは拡張子が.xlsxのファイルしか読み取ることができません．
　マクロが組み込まれたExcelファイルは取り扱うことができません．
 (そもそも.xlsx形式以外での提出は認められていません)
・CostTableの変更には対応可能です．(ただし，CostTableのフォーマット変更は不可)
・UnitCostに計算式を入れる場合はチェックを行いません．注意してください．

## Update
修正・改良したらここにログを書き込みましょう．

## Author
野口晴臣
080-7826-1613
医学部医学科4回生(2020年度)