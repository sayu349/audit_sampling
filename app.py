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
# #xlsxファイルの読み込み
@app.route("/", methods=["POST","GET"])
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
        # 関数実行
        file_stream = audit_sampling(
                                    file= file,
                                    sheet_name= sheetNameSelectBox,
                                    xlsx_or_csv = "xlsx",
                                    amount_column_name = columnNameSelectBox,
                                    row_number = rowNumberInput
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
        # 関数実行
        file_stream = audit_sampling(
                                    file= file,
                                    xlsx_or_csv = "csv",
                                    amount_column_name = columnNameSelectBox,
                                    row_number = rowNumberInput
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


# ============================================================
# 実行
# ============================================================
if __name__== '__main__':
    app.run(debug=True)