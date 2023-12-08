from wtforms import Form, IntegerField, FileField, StringField, SelectField, SubmitField
from wtforms.validators import DataRequired
from flask_wtf.file import FileAllowed

"""
ファイルアップロードフォーム
input1.htmlにて使用
"""
class ExcelUploadForm(Form):
    # エクセルファイル入力
    excel_file = FileField("Excelファイル (.xlsx)", validators=[
        DataRequired(),
        FileAllowed(["xlsx"], "Excelファイルのみ!")
    ])

    # 決定ボタン
    submit = SubmitField("テーブル選択へ")

"""
シート名入力画面
input2.htmlにて使用
"""
class SheetNameForm(Form):
    # シート名選択
    sheet_name = SelectField('シート名を選択', choices=[])

    # 決定ボタン
    submit = SubmitField("要素を含む行の選択へ")