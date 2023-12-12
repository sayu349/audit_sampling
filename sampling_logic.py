# =========================================================================
# ライブラリ
# =========================================================================
# 計算に必要なライブラリ
import numpy as np
import pandas as pd
from scipy.stats import poisson
import math

# 一時ファイル作成に必要なライブラリ
import io


# =========================================================================
# ポアソン分布による金額単位サンプリングによるサンプル数算定の関数
# =========================================================================
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


# =========================================================================
# 監査サンプリング
# =========================================================================
def audit_sampling(xlsx_or_csv, file, amount_column_name, row_number, sheet_name=None):
    if xlsx_or_csv == "xlsx":
        sample_data = pd.read_excel(
                                    file,
                                    sheet_name=sheet_name,
                                    header=row_number-1
                                    )
    else:
        sample_data = pd.read_csv(
                                    file,
                                    encoding="UTF-8",
                                    header=row_number-1,
                                    thousands=","
                                )

    # 母集団の金額合計算出
    print(sample_data[amount_column_name])
    total_amount_column_name = sample_data[amount_column_name].sum()
    # 母集団の金額合計
    N =  total_amount_column_name
    # 手続実施上の重要性
    pm = 12155185
    # ランダムシード(サンプリングの並び替えのステータスに利用、任意の数を入力)
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

    # サンプリングシートに記載用の、パラメータ一覧
    sampling_param = pd.DataFrame(
                                    [
                                        ['母集団合計', N],
                                        ['手続実施上の重要性', pm],
                                        ['リスク', audit_risk],
                                        ['内部統制', internal_control],
                                        ['random_state', random_state]
                                    ]
                                )

    # 母集団をまずは降順に並び替える（ここで並び替えるのは、サンプル出力の安定のため安定のため）
    sample_data = sample_data.sort_values(amount_column_name, ascending=False)

    # 母集団をシャッフル
    shuffle_data = sample_data.sample(frac=1, random_state=random_state) #random_stateを使って乱数を固定化する

    # サンプリング区間の算定
    m = N/n

    # 列の追加
    shuffle_data['cumsum'] = shuffle_data[amount_column_name].cumsum() # 積み上げ合計
    shuffle_data['group'] = shuffle_data['cumsum']//m # サンプルのグループ化

    result_data = shuffle_data.loc[shuffle_data.groupby('group')['cumsum'].idxmin(), ]

    # メモリ内のバイトストリームにExcelファイルを作成する
    file_stream = io.BytesIO()

    # result_data.to_excel(file_name, encoding="shift_jis", index=False)
    writer = pd.ExcelWriter(file_stream)

    # 全レコードを'全体'シートに出力
    sample_data.to_excel(writer, sheet_name = '母集団', index=False)
    # サンプリング結果を、サンプリングシートに記載
    result_data.to_excel(writer, sheet_name = 'サンプリング結果', index=False)
    # サンプリングの情報追記
    sampling_param.to_excel(writer, sheet_name = 'サンプリングパラメータ', index=False, header=None)

    # Excelファイルを保存して閉じる
    writer.close()

    # バイトストリームをリセット
    file_stream.seek(0)

    return file_stream