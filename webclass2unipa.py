import csv
import pandas as pd
import openpyxl
import sys

# WebClassの出欠データをユニバーサルパスポートの出欠データに差し込むスクリプト
#
# 使用方法：
# 0. csv,pandas,xlrd,xlwt,openpyxl のライブラリをpython環境にインストールする
# 1. WebClassから出席ファイルをダウンロードし　例）ファイル名 src_attendance.csv として、
#  　このスクリプトと同じフォルダに置く
# 2. src_attendance.csv には15回分（回数固定）の出席データが記録されているものとする
# 3. 出席として変換する webclass 側のデータを変数 attend_code にセットする。
#    それ以外は欠席として変換される。
#    attend_code の規定値は '出'
#    WebClassのコースに未登録の学籍番号は、全欠席として処理される。
# 4. ユニバーサルパスポートから対応する科目の出欠登録データをダウンロードして
#  　このスクリプトと同じフォルダに置く
#    ヘッダー行を含んだ形式をダウンロードすること。
#    例）ファイル名 dst_attendance.xlsx
# 5. コマンドプロンプトから以下を実行
#
# cmd> yourpath/python webclass2unipa.py src_attendance.csv dst_attendance.xlsx
#
#    csvやxlsxのファイル名にパスを含むと出力ファイルの生成に失敗するので避ける。
# 6. ユニバーサルパスポートにアップロードするためのファイル upload_dst_attendance.xlsx が書き出される。
# 7. upload_dst_attendance.xlsxをアップロードする。

src_file_name ='dummy.csv'
dst_file_name ='dummy.xlsx'
upload_file_name = 'upload_dummy.xlsx'
attend_code ='出'
args = sys.argv
if 3 <= len(args):
    src_file_name = args[1]
    dst_file_name = args[2]
else:
    print('Arguments are too short')
upload_file_name = 'upload_' + dst_file_name

# csvからデータを抜き出してxlsxに差し込む程度の処理にはpandasは機能過剰だった。
# xlrd と xlwt が有れば事足りる。

# pandasのDataFrameに xlsx を読み込んでいる。
# from universal passport attends xlsx
# 先頭行をヘッダーに用いる。インデックスは0からの連番とする。
df = pd.read_excel('./' + dst_file_name, header=0, index_col=None,converters={2:str,8:int})
# 先頭行を読み飛ばしてヘッダーを用いない。インデックスは0からの連番とする。
# df = pd.read_excel('./kobashi.kazuhide_Atb003Exc01_20200219101753001.xlsx', header=None, index_col=None,skiprows=[0],converters={2:str,8:int})
# print(df)

csv_file = open(src_file_name, "r", encoding="ms932", errors="", newline="" )
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
        if temp == attend_code :  #webclassの出席記録
            df.iat[i,8] = 0 #出席
        else:
            df.iat[i,8] = 3 #欠席
    else:
        df.iat[i,8] = 3 #webclassに未登録の場合欠席扱い
    print(i,id,df.iat[i,8])

# ファイルに書き戻す

print('src file: ' + src_file_name)
print('dst file: ' + dst_file_name)
print('upload file: ' + upload_file_name)

df.to_excel(upload_file_name, index=False, header=False)