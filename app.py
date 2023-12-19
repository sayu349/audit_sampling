# ============================================================
# webアプリケーション用のライブラリ
# ============================================================
from flask import Flask
from flask import render_template
from flask import request
from flask import send_file


# ============================================================
# 監査サンプリング関数のインポート
# ============================================================
from sampling_logic import audit_sampling


# ============================================================
# Flask
# ============================================================
app = Flask(__name__)


# ============================================================
# ルーティング
# ============================================================
# home
@app.route("/")
def home_page():
    return render_template("home.html")

# #xlsxファイルの読み込み
@app.route("/sampling_xlsx", methods=["POST","GET"])
def sampling_xlsx_page():
    # POST
    if request.method == "POST":
        # xlsxファイル取得
        file = request.files["fileInput"]
        # シート名取得
        sheetNameSelectBox = request.form["sheetNameSelectBox"]
        # 行番号
        rowNumberInput = int(request.form["rowNumberInput"])
        # 金額列名
        columnNameSelectBox = request.form["columnNameSelectBox"]
        # シード値取得
        randomState = int(request.form["randomState"])
        # 許容逸脱金額・手続実施上の重要性
        pm = int(request.form["pm"])
        # 監査リスク
        auditRisk = request.form["auditRisk"]
        # 内部統制
        internalControl = request.form["internalControl"]
        # 関数実行
        file_stream = audit_sampling(
                                    file= file,
                                    sheet_name= sheetNameSelectBox,
                                    xlsx_or_csv = "xlsx",
                                    amount_column_name = columnNameSelectBox,
                                    row_number = rowNumberInput,
                                    random_state=randomState,
                                    pm = pm,
                                    audit_risk = auditRisk,
                                    internal_control= internalControl
                                    )
        return send_file(
                        file_stream,
                        download_name=f'{sheetNameSelectBox}サンプル.xlsx',
                        as_attachment=True,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
    # GET
    else:
        return render_template("sampling_xlsx.html")


# csvの読み込み
@app.route("/sampling_csv", methods=["POST","GET"])
def sampling_csv_page():
    # POST
    if request.method == "POST":
        # csvファイル取得
        file = request.files["fileInput"]
        # ファイル名取得
        fileName = file.filename
        fileName = fileName.replace(".csv", "")
        # 行番号取得
        rowNumberInput = int(request.form["rowNumberInput"])
        # 金額列名取得
        columnNameSelectBox = request.form["columnNameSelectBox"]
        # シード値取得
        randomState = int(request.form["randomState"])
        # 許容逸脱金額・手続実施上の重要性
        pm = int(request.form["pm"])
        # 監査リスク
        auditRisk = request.form["auditRisk"]
        # 内部統制
        internalControl = request.form["internalControl"]
        # 関数実行
        file_stream = audit_sampling(
                                    file= file,
                                    xlsx_or_csv = "csv",
                                    amount_column_name = columnNameSelectBox,
                                    row_number = rowNumberInput,
                                    random_state= randomState,
                                    pm = pm,
                                    audit_risk = auditRisk,
                                    internal_control= internalControl
                                    )
        return send_file(
                        file_stream,
                        download_name=f'{fileName}サンプル.xlsx',
                        as_attachment=True,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
    # GET
    else:
        return render_template("sampling_csv.html")

# 404エラーが発生した場合の処理
@app.errorhandler(404)
def error_404(error): # errorは消さない！
    return render_template('404.html')


# ============================================================
# 実行
# ============================================================
"""
if __name__== '__main__':
    app.run(debug=True, host="192.168.0.81", port=80)
"""

if __name__== '__main__':
    app.run(debug=True)