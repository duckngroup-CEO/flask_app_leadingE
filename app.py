from cmath import tan
import os
from re import S
import shutil
from flask import Flask, request, redirect, url_for, render_template, Markup,session,Response, make_response,send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from PIL import Image
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import japanize_matplotlib
matplotlib.use('agg')
import openpyxl as xl
import datetime
import make_invoice as invo

# 変数の宣言
UPLOAD_FOLDER = 'original_files'
ALLOWED_EXTENSIONS = {'xlsx'}
XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

# ファイルパスの指定
ORIGINAL_FILEPATH = "static/original_files/original_data.xlsx"
OUTPUT_FILEPATH = "static/output_files/output_data/output_data.xlsx"

INVOICE_TEMPLATES = "static/invoice_templates/invoice_template1.xlsx"
OUTPUT_INVOICES_DIRPATH = "static/output_files/output_invoices"

# Flaskのインスタンス化とUPLOAD_FOLDERの定義
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
sales_data = 'データなし'
app.secret_key = 'secret'

# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# 共通で使う関数の定義
# allowed_file()関数の定義
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# エクセルを読み込む関数
def read_excel(filepath):
    try:
        global sales_data
        sales_data = pd.read_excel(filepath, sheet_name='売上')
        return sales_data
    except Exception as e:
        print('Error:エクセルファイルの読み込みに失敗しました。', e)

# エクセルファイルを吐き出す関数
def output_excel(df, filepath):
    df.to_excel(filepath, sheet_name="売上", index=False, header=True)
    return df

# インデックスがあるかどうかチェックする関数
def check_index(index):
    sales_data = read_excel(ORIGINAL_FILEPATH)
    index_list = sales_data.index.to_list()
    print(index_list)
    if index in index_list:
        return True
    else:
        return False

# zipファイルに圧縮する関数
def make_zip():
    shutil.make_archive('static/output_files/output_zip/請求書', format='zip', root_dir='static/output_files/output_invoices')

# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# Topページのルーティングとindex()関数の定義
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")
# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# 判定結果ページのルーティングとresult()関数の定義
@app.route("/result", methods=["GET", "POST"])
def result():
    if request.method == "POST":
        # ファイルの存在と形式を確認
        if "file" not in request.files:
            print("File doesn't exist!")
            return redirect(url_for("index"))
        file = request.files["file"]
        if not allowed_file(file.filename):
            print(file.filename + ": File not allowed!")
            return redirect(url_for("index"))

        # ファイルの保存
        if os.path.isdir(UPLOAD_FOLDER):
            shutil.rmtree(UPLOAD_FOLDER)
        os.mkdir(UPLOAD_FOLDER)
        filename = secure_filename(file.filename)  # ファイル名を安全なものに
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        print(filepath)

        try:
            global sales_data
            sales_data = read_excel(ORIGINAL_FILEPATH)
            return render_template("result.html", result=Markup(sales_data.to_html()))
        except Exception as e:
            print('Error:エクセルファイルの読み込みに失敗しました。', e)
    else:
        return redirect(url_for("index"))
# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# トップページへの遷移
@app.route("/result_top", methods=["GET","POST"])
def result_top():
    sales_data = read_excel(ORIGINAL_FILEPATH)
    return render_template("result.html", result=Markup(sales_data.to_html()))

# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# データ消去後のページへ遷移
@app.route("/result_deleted", methods=["POST"])
def delete():
    if request.method == "POST":
        got_id = int(request.form.get("got_id", False))
        if check_index(got_id)==True:
            sales_data = read_excel(ORIGINAL_FILEPATH)
            sales_data = sales_data.drop([got_id])
            output_excel(sales_data,OUTPUT_FILEPATH)
            return render_template("deleted.html", result=Markup(sales_data.to_html()))
        else:
            sales_data = read_excel(ORIGINAL_FILEPATH)
            return render_template("id_error.html", result=Markup(sales_data.to_html()))
        

# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# データ追加後のページへ遷移
@app.route("/result_post", methods=["POST"])
def create():
    if request.method == "POST":
        sales_date = request.form.get("sales_date", False)
        cus_id = request.form.get("cus_id", False)
        pro_name = request.form.get("pro_name", False)
        tanka = int(request.form.get("tanka", False))
        suryo = int(request.form.get("suryo", False))
        total = int(request.form.get("total", False))
        sales_data = read_excel(ORIGINAL_FILEPATH)
        max_index = max(sales_data.index)+1
        sales_data.loc[max_index] = [sales_date, cus_id, pro_name, tanka, suryo, total]
        output_excel(sales_data,OUTPUT_FILEPATH)
        return render_template("created.html", result=Markup(sales_data.to_html()))

# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# データ編集ページへ遷移
@app.route("/result_updated", methods=["POST"])
def update():
    if request.method == "POST":
        got_id = int(request.form.get("got_id", False))
        if check_index(got_id)==True:
            sales_data = read_excel(ORIGINAL_FILEPATH)
            sales_data = sales_data.iloc[got_id:got_id+1]
            session["id"]=got_id
            return render_template("edit_page.html", result=Markup(sales_data.to_html()))
        else:
            sales_data = read_excel(ORIGINAL_FILEPATH)
            return render_template("id_error.html", result=Markup(sales_data.to_html()))

# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# 実際にデータ編集後、データ編集後のページへ遷移
@app.route("/result_edit", methods=["POST"])
def update_exe():
    if request.method == "POST":
        got_id = int(session.get("id")) 
        sales_data = read_excel(ORIGINAL_FILEPATH)
        # sales_data = sales_data.iloc[got_id:got_id+1]
        sales_data["売上日"][got_id] = request.form.get("sales_date", False)
        sales_data["顧客名"][got_id] = request.form.get("cus_id", False)
        sales_data["商品名"][got_id] = request.form.get("pro_name", False)
        sales_data["単価"][got_id] = int(request.form.get("tanka", False))
        sales_data["数量"][got_id] = int(request.form.get("suryo", False))
        sales_data["合計"][got_id] = int(request.form.get("total", False))
        output_excel(sales_data,OUTPUT_FILEPATH)
        return render_template("updated.html", result=Markup(sales_data.to_html()))
# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# グラフ分析ページへ遷移
@app.route("/graph", methods=["GET", "POST"])
def graph():
    sales_data_= read_excel(ORIGINAL_FILEPATH)
    # 売上高×日付別
    fig = plt.figure(figsize=(10,5))
    plt.barh(sales_data['売上日'], sales_data['合計'],color='pink')
    plt.xticks(rotation=0)
    # plt.yticks([0,1000000,100000])

    fig.savefig("./static/img/img_sales_date.png")
    image_path = "./static/img/img_sales_date.png"

    # 売上高×商品別
    fig2 = plt.figure(figsize=(10,5))
    plt.barh(sales_data['商品名'], sales_data['合計'], color='red')
    plt.xticks(rotation=0)
    fig2.savefig("./static/img/img_sales_products.png")
    image_path2 = "./static/img/img_sales_products.png"

    # 売上高×顧客名
    fig3 = plt.figure(figsize=(10,5))
    plt.barh(sales_data['顧客名'], sales_data['合計'],color='blue')
    plt.xticks(rotation=0)
    fig3.savefig("./static/img/img_sales_customer.png")
    image_path3 = "./static/img/img_sales_customer.png"

    # グラフを辞書型に格納
    images = {
        'date': image_path,
        'products': image_path2,
        'customers': image_path3,
    }

    return render_template("graph.html", graphs=images)
# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# 請求書作成
@app.route("/invoice_make_done", methods=["GET", "POST"])
def make_invoice():
    if request.method == "POST":
        sales_data = read_excel(OUTPUT_FILEPATH)
        invo.make_invoice()
        make_zip()
        return render_template("invoice_make_done.html")
# ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
# メイン関数の実行
if __name__ == "__main__":
    app.run(debug=False)