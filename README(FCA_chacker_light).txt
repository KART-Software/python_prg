
FCA_chacker_light
(ver. Python 3.8.4)

FCAからの値をチェックし誤りがあった場合，リストアップします．
※FCAチェッカーの簡易版です．MaterialのUnitCostなどは主要なもののみ確認していきます．

## Description
2020年度の電子提出版コストレポートのディレクトリ構造に合わせて作成したBOM作成プログラムです．
Python3.8.4, Openpyxl3.03にて動作します．
実行すると以下の入力が求められます．
・最上階層のフォルダ名：FCAとBOMがまとめて入っているフォルダです
・実行するシステムアセンブリのフォルダ名
・車番(半角二桁)

##VS. FCA_chacker_revised
本プログラムは, FCA_chacker_revisedと異なり，コストテーブルを参照しません．
そのため，動作は軽いです．
また, FCA_chacker_revisedと異なり，大学名，車番等の確認も行います．

## Requirement
・Python3の導入が必須です
・外部ライブラリopenpyxlを使用しています．導入が必須です．
　コマンドプロンプトを起動し，"pip install openpyxl"と入力するだけでインストールできます．

##Usage
python_prgのフォルダごと，コストレポートの最上階層と同じ階層にコピーしてください．
example)
--------015_KyotoUniversity_FSAEJ_CR---------015_A1_BR
 　|                                                         |-----015_A2_EN
　 |                                                         |-----015_A3_FR
    |                                                         |            ：
　 |　　　　　　　　　　　　　　　　　　　　　   |-----015_KyotoUniversity_FSAEJ_CR_BOM.xlsx
　 |
    |---python_prg-------------------------------------FCA_chacker_light.py
                                                               |------README(FCA_chacker_light).txt

次にIDLEを起動し，File→OpenからFCA_chacker_light.pyを開いてください．
(プログラムを動かすのみであればパワーシェルから開いても可能ですが，コードの修正はできません)
最後に，F5キーを押すと実行されます．

## Check_list
以下の項目をチェックします．
・University
・Suffix
・Car_Number
・Material，Process，Fastener, ToolingにおいてItemOrderが順に並んでいるか
　また，sub_totalの関数が全ての合計となっているか
・以下のMaterialのUnitCost
	Aluminum, Premium
	Aluminum, Normal
	Steel, Alloy
	Steel, Mild
	Steel, Stainless
	Iron
	Rubber
	Plastic, Nylon
	Carbon Fiber, 1 Ply
	Honeycomb, Aluminum
	Honeycomb, Nomex
・以下のProcessのUnitCost
	Machining Setup, Install and remove
	Machining Setup, Change
	Laser Cut
	Machining
	Weld
・以下のProcessのMultiplier
	Aluminum
	Steel
	Cast Iron
	Foam
	Composite
	Plastic
	Stainless, Steel
・単位にcm^4が含まれていないか

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
・FCA内のsub_totalを検索する方法は，FCA内のB列からMaterialなどの文字列を検索し，そこから下のセルを順に見ていきます．
  そして，空白セルに達した時点で，その行の特定の列のセルをsub_totalとして読み取っています．
・本プログラムではget_sheet_names関数を使用しています．verの更新でこの関数が削除される可能性があります．

## Update
修正・改良したらここにログを書き込みましょう．

## Author
野口晴臣
080-7826-1613
医学部医学科4回生(2020年度)