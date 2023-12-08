# webアプリケーション用のライブラリ
from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename

# 統計等のライブラリ
import numpy as np
import pandas as pd
import scipy as sp
from scipy.stats import poisson
from scipy.stats import binom
import math

# 可視化ライブラリ
import matplotlib.pyplot as plt
import seaborn as sns

# フォーム情報をインポート
from forms import ExcelUploadForm,SheetNameForm

# ポアソン分布による金額単位サンプリングによるサンプル数算定の関数
def sample_poisson(N, pm, ke, alpha, audit_risk, internal_control='依拠しない'):
    k = np.arange(ke+1)
    pt = pm/N
    n = 1
    while True:
        mu = n*pt
        pmf_poi = poisson.cdf(k, mu)
        if pmf_poi.sum() < alpha:
            break
        n += 1
    if audit_risk == 'SR':
        n = math.ceil(n)
    if audit_risk == 'RMM-L':
        n = math.ceil(n/10*2)
    if audit_risk == 'RMM-H':
        n = math.ceil(n/2)
    if internal_control == '依拠する':
        n = math.ceil(n/3)
    return n

def audit_sampling(file_name, sheet_name, amount, row_number):
    # 読み込み用のシート名(.xlsxまで入れる)
    # file_name = '母集団.xlsx'
    # sheet_name = '製)原材料仕入'
    # amount = '金額'
    # row_nbumebr = '行番号

    sample_data = pd.read_excel(file_name, sheet_name=sheet_name, header=row_number-1)

    # 金額がマイナスなので、それを修正
    #sample_data[amount] = sample_data[amount]*-1

    # 母集団の金額が正しいかチェック
    total_amount = sample_data[amount].sum()
    print(total_amount)



    # 変動パラメータの設定

    # 母集団の金額合計
    N =  total_amount
    # 手続実施上の重要性
    pm = 12155185
    # ランダムシード　(サンプリングの並び替えのステータスに利用、任意の数を入力)
    random_state = 2
    # 監査リスク
    audit_risk = 'RMM-L'
    # 内部統制
    internal_control = '依拠しない'


    # 予想虚偽表示金額（変更不要）
    ke = 0
    alpha = 0.05

    # サンプルサイズnの算定
    n = sample_poisson(N, pm, ke, alpha, audit_risk, internal_control)
    print(n)

    # サンプリングシートに記載用の、パラメータ一覧
    sampling_param = pd.DataFrame([['母集団合計', N],
                                ['手続実施上の重要性', pm],
                                ['リスク', audit_risk],
                                ['内部統制', internal_control],
                                ['random_state', random_state]])

    # 母集団をまずは降順に並び替える（ここで並び替えるのは、サンプル出力の安定のため安定のため）
    sample_data = sample_data.sort_values(amount, ascending=False)

    # 母集団をシャッフル
    shuffle_data = sample_data.sample(frac=1, random_state=random_state) #random_stateを使って乱数を固定化する
    shuffle_data.head()

    # サンプリング区間の算定
    m = N/n
    print(m)

    # 列の追加
    shuffle_data['cumsum'] = shuffle_data[amount].cumsum() # 積み上げ合計
    shuffle_data['group'] = shuffle_data['cumsum']//m # サンプルのグループ化
    shuffle_data.head()

    result_data = shuffle_data.loc[shuffle_data.groupby('group')['cumsum'].idxmin(), ]
    result_data

    file_name = '{}サンプル.xlsx'.format(sheet_name)
    # result_data.to_excel(file_name, encoding="shift_jis", index=False)

    writer = pd.ExcelWriter(file_name)

    # 全レコードを'全体'シートに出力
    sample_data.to_excel(writer, sheet_name = '母集団', index=False)
    # サンプリング結果を、サンプリングシートに記載
    result_data.to_excel(writer, sheet_name = 'サンプリング結果', index=False)
    # サンプリングの情報追記
    sampling_param.to_excel(writer, sheet_name = 'サンプリングパラメータ', index=False, header=None)

    # Excelファイルを保存して閉じる
    writer.close()


# =====================================================================
# ルーティング
# =====================================================================
app = Flask(__name__)

# ファイル入力画面
@app.route("/input1", methods=['GET', 'POST'])
def input1_page():
    form = ExcelUploadForm(request.form)
    if request.method == 'POST':
        file = request.files['excel_file']
        filename = secure_filename(file.filename)
        file.save("tmp/" + filename)
        return redirect(url_for("input2_page", filename=filename))
    else:
        return render_template("input1.html", form=form)

# シート名入力画面
@app.route("/input2", methods=["GET", "POST"])
def input2_page():
    form = SheetNameForm(request.form)
    filename = request.args.get('filename', None)

    # Excelファイルを読み込む
    xls = pd.ExcelFile("tmp/" + filename)
    # シート名のリストを取得
    sheet_names = xls.sheet_names

    return render_template("input2.html", filename=filename,sheet_names=sheet_names,form=form)


# =====================================================================
# 実行
# =====================================================================
if __name__ == "__main__":
    app.run(debug=True)



