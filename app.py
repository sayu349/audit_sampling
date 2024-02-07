# ============================================================
# webアプリケーション用のライブラリ
# ============================================================
from flask import Flask
from flask import render_template
from flask import request
from flask import send_file
# ============================================================


# ============================================================
# 監査サンプリング関数のインポート
# ============================================================
# 属性サンプリング関数
from sampling_logic import attribute_sampling

# 金額単位サンプリング
from sampling_logic import audit_sampling


# ============================================================
# Flask
# ============================================================
app = Flask(__name__)


# ============================================================
# ルーティング
# ============================================================
# ホーム
@app.route("/")
def home_page():
    return render_template("home.html")


# ============================================================
# 属性サンプリング・xlsx版
# ============================================================
@app.route("/attribute_sampling_xlsx", methods=["POST","GET"])
def attribute_sampling_xlsx_page():
    # POST
    if request.method == "POST":
        # xlsxファイル
        file = request.files["fileInput"]
        # シート名
        sheetNameSelectBox = request.form["sheetNameSelectBox"]
        # ヘッダー行
        rowNumberInput = int(request.form["rowNumberInput"])
        # シード値
        randomState = int(request.form["randomState"])
        # 許容逸脱率の上限
        pt = float(request.form["pt"])
        # 予想逸脱金額と優位水準が任意してされた場合
        if "ke" in request.form and "alpha" in request.form:
            ke = int(request.form["ke"])
            alpha = float(request.form["alpha"])
        # デフォルトの予想逸脱金額と優位水準を使う
        else:
            ke = 0
            alpha = 0.05
        try:
            # 属性サンプリング関数実行
            file_stream = attribute_sampling(
                                            file = file,
                                            sheet_name = sheetNameSelectBox,
                                            xlsx_or_csv ="xlsx",
                                            row_number = rowNumberInput,
                                            random_state = randomState,
                                            pt = pt,
                                            ke = ke,
                                            alpha = alpha)
            # 成功した場合
            return send_file(
                            file_stream,
                            download_name=f'{sheetNameSelectBox}サンプル.xlsx',
                            as_attachment=True,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
        except:
            # 失敗した場合
            return render_template("error.html", return_page ="attribute_xlsx")
    # GET
    else:
        return render_template("attribute_sampling_xlsx.html")


# ============================================================
# 属性サンプリング・csv版
# ============================================================
@app.route("/attribute_sampling_csv", methods=["POST","GET"])
def attribute_sampling_csv_page():
    # POST
    if request.method == "POST":
        # xlsxファイル
        file = request.files["fileInput"]
        # ファイル名
        fileName = file.filename
        fileName = fileName.replace(".csv", "")
        # ヘッダー行
        rowNumberInput = int(request.form["rowNumberInput"])
        # シード値
        randomState = int(request.form["randomState"])
        # 許容逸脱率の上限
        pt = float(request.form["pt"])
        # 予想逸脱金額と優位水準が任意してされた場合
        if "ke" in request.form and "alpha" in request.form:
            ke = int(request.form["ke"])
            alpha = float(request.form["alpha"])
        # デフォルトの予想逸脱金額と優位水準を使う
        else:
            ke = 0
            alpha = 0.05
        # 属性サンプリング関数実行
        file_stream = attribute_sampling(
                                        file = file,
                                        xlsx_or_csv ="csv",
                                        row_number = rowNumberInput,
                                        random_state = randomState,
                                        pt = pt,
                                        ke = ke,
                                        alpha = alpha)
        # 成功した場合
        return send_file(
                        file_stream,
                        download_name=f'{fileName}サンプル.xlsx',
                        as_attachment=True,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
    # GET
    else:
        return render_template("attribute_sampling_csv.html")


# ============================================================
# 金額単位サンプリング・xlsx版
# ============================================================
@app.route("/money_sampling_xlsx", methods=["POST","GET"])
def sampling_xlsx_page():
    # POST
    if request.method == "POST":
        # xlsxファイル
        file = request.files["fileInput"]
        # シート名
        sheetNameSelectBox = request.form["sheetNameSelectBox"]
        # ヘッダー行
        rowNumberInput = int(request.form["rowNumberInput"])
        # 金額単位サンプリングを用いる列
        columnNameSelectBox = request.form["columnNameSelectBox"]
        # シード値
        randomState = int(request.form["randomState"])
        # 手続実施上の重要性
        pm = int(request.form["pm"])
        # 監査リスク
        auditRisk = request.form["auditRisk"]
        # 内部統制
        internalControl = request.form["internalControl"]

        # 予想逸脱金額と優位水準が任意してされた場合
        if "ke" in request.form and "alpha" in request.form:
            ke = int(request.form["ke"])
            alpha = float(request.form["alpha"])
        # デフォルトの予想逸脱金額と優位水準を使う
        else:
            ke = 0
            alpha = 0.05

        try:
            # 属性サンプリング関数実行
            file_stream = audit_sampling(
                                        file= file,
                                        sheet_name= sheetNameSelectBox,
                                        xlsx_or_csv = "xlsx",
                                        amount_column_name = columnNameSelectBox,
                                        row_number = rowNumberInput,
                                        random_state=randomState,
                                        pm = pm,
                                        audit_risk = auditRisk,
                                        ke = ke,
                                        alpha = alpha,
                                        internal_control= internalControl
                                        )
            # 成功した場合
            return send_file(
                            file_stream,
                            download_name=f'{sheetNameSelectBox}サンプル.xlsx',
                            as_attachment=True,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
        except:
            # 失敗した場合
            return render_template("error.html", return_page ="money_xlsx")
    # GET
    else:
        return render_template("money_sampling_xlsx.html")


# ============================================================
# 金額単位サンプリング・csv版
# ============================================================
@app.route("/money_sampling_csv", methods=["POST","GET"])
def sampling_csv_page():
    # POST
    if request.method == "POST":
        # csvファイル
        file = request.files["fileInput"]
        # ファイル名
        fileName = file.filename
        fileName = fileName.replace(".csv", "")
        # ヘッダー行
        rowNumberInput = int(request.form["rowNumberInput"])
        # 金額単位サンプリングを用いる列
        columnNameSelectBox = request.form["columnNameSelectBox"]
        # シード値
        randomState = int(request.form["randomState"])
        # 手続実施上の重要性
        pm = int(request.form["pm"])
        # 監査リスク
        auditRisk = request.form["auditRisk"]
        # 内部統制
        internalControl = request.form["internalControl"]
        # 予想逸脱金額と優位水準が任意してされた場合
        if "ke" in request.form and "alpha" in request.form:
            ke = int(request.form["ke"])
            alpha = float(request.form["alpha"])
        # デフォルトの予想逸脱金額と優位水準を使う
        else:
            ke = 0
            alpha = 0.05
        try:
            # 属性サンプリング関数実行
            file_stream = audit_sampling(
                                        file= file,
                                        xlsx_or_csv = "csv",
                                        amount_column_name = columnNameSelectBox,
                                        row_number = rowNumberInput,
                                        random_state= randomState,
                                        pm = pm,
                                        audit_risk = auditRisk,
                                        ke = ke,
                                        alpha = alpha,
                                        internal_control= internalControl
                                        )
            # 成功した場合
            return send_file(
                            file_stream,
                            download_name=f'{fileName}サンプル.xlsx',
                            as_attachment=True,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
        except:
            # 失敗した場合
            return render_template("error.html", xlsx_or_csv ="money_csv")
    # GET
    else:
        return render_template("money_sampling_csv.html")


# ============================================================
# エラーハンドリング
# ============================================================
# サンプリング中におけるエラー
@app.route("/sampling_error")
def sampling_error_page():
    return render_template("error.html")

# 404エラー
@app.errorhandler(404)
def error_404(error): # errorは消さない！
    return render_template('404.html')



# ============================================================
# 実行
# ============================================================
if __name__== '__main__':
    app.run(debug=True)
