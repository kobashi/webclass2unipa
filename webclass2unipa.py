import csv
import pandas as pd
import openpyxl
import sys

# WebClassの出欠データをユニバーサルパスポートの出欠データに差し込むスクリプト
#
# 使用方法：
# 0. csv,pandas,xlrd,xlwt,openpyxl のライブラリをpython環境にインストールする
# 1. WebClassから出席ファイルをダウンロードしファイル名 attendance.csv （既定）として、
#  　このスクリプトと同じフォルダに置く
# 2. attendance.csv には15回分（回数固定）の出席データが記録されているものとする
#    点数 10 が記録された学籍番号を出席とし、それ以外を欠席として処理する。
#    WebClassのコースに未登録の学籍番号は、全欠席として処理される。
# 3. ユニバーサルパスポートから対応する科目の出欠登録データをダウンロードして
#  　このスクリプトと同じフォルダに置く
#    ヘッダー行を含んだ形式をダウンロードすること。
# 4. コマンドプロンプトから以下を実行
#
# cmd> yourpath/python webclass2unipa.py ユニバーサルパスポート出席ファイル.xlsx
#
# 5. ユニバーサルパスポートにアップロード用のファイル attendance.xlxs が書き出される。
# 6. attendance.xlxsをアップロードする。

file_name ='dummy.xlsx'
args = sys.argv
if 2 <= len(args):
    file_name = args[1]
else:
    print('Arguments are too short')

# csvからデータを抜き出してxlsxに差し込む程度の処理にはpandasは機能過剰だった。
# xlrd と xlwt が有れば事足りる。

# pandasのDataFrameに xlsx を読み込んでいる。
# from universal passport attends xlsx
# 先頭行をヘッダーに用いる。インデックスは0からの連番とする。
df = pd.read_excel('./' + file_name, header=0, index_col=None,converters={2:str,8:int})
# 先頭行を読み飛ばしてヘッダーを用いない。インデックスは0からの連番とする。
# df = pd.read_excel('./kobashi.kazuhide_Atb003Exc01_20200219101753001.xlsx', header=None, index_col=None,skiprows=[0],converters={2:str,8:int})
# print(df)

csv_file = open("./attendance.csv", "r", encoding="ms932", errors="", newline="" )
# from webclass attends list
f = csv.reader(csv_file, delimiter=",", doublequote=True, lineterminator="\r\n", quotechar='"', skipinitialspace=True)

# 見出しなどの不要な行を読み飛ばす
title = next(f)
date  = next(f)
blank = next(f)
header = next(f)
# print(header)
count1 = next(f)
count2 = next(f)
count3 = next(f)

# csvから出席データを辞書（keyは学籍番号）としてに抽出
attends_dict = {}
for row in f:
    # 学籍番号と15回分の出席リスト（開講回数は15回に固定）
    id = row[1]
    attends = row[2:17]
    print(id, attends)
    attends_dict[id] = attends

# csvから抽出したデータを確認表示
for key in attends_dict.keys():
    for col in attends_dict[key]:
        print(key,col)

# pandasのdfから学籍番号と出欠のカラムにアクセスしてcsvの出席データを差し込む
lec_num = 0
old_id = 0
for i,id in zip(df.index, df[df.columns[2]]):
    if old_id != id :
        lec_num = 0
        old_id = id
    else:
        lec_num += 1
    attends_list = attends_dict.get(id)
    if attends_list != None : #webclassに登録有り
        temp = attends_list[lec_num]
        if temp == '10' :  #webclassの出席記録
            df.iat[i,8] = 0 #出席
        else:
            df.iat[i,8] = 3 #欠席
    else:
        df.iat[i,8] = 3 #webclassに未登録の場合欠席扱い
    print(i,id,df.iat[i,8])

# ファイルに書き戻す
df.to_excel('./attendance.xlsx', index=False, header=False)