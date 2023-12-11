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
import sampling_logic


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
        # audit_sampling_xlsx(xlsxファイル, シート名, カラム名, 行番号)
        sampling_logic.audit_sampling_xlsx(file, sheetNameSelectBox, columnNameSelectBox, rowNumberInput)
        return send_file(f"{sheetNameSelectBox}サンプル.xlsx")
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
        # 行番号取得
        rowNumberInput = int(request.form["rowNumberInput"])
        # 金額列名取得
        columnNameSelectBox = request.form["columnNameSelectBox"]
        # "xxx.csv" ⇒ "xxx"（.csvを取り除く）
        fileName = fileName.replace(".csv", "")
        # 関数実行
        # audit_sampling_csv(csvファイル, csvファイル名, カラム名, 行番号)
        sampling_logic.audit_sampling_csv(file, fileName, columnNameSelectBox, rowNumberInput)
        return send_file(f"{fileName}サンプル.xlsx")
    # GET
    else:
        return render_template("sampling_csv.html")


# ============================================================
# 実行
# ============================================================
if __name__== '__main__':
    app.run(debug=True, host="192.168.0.81", port=80)