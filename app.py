import random
import requests
from google import genai
from google.genai import types
import json
import re #単語を抜き出すのに使う
import nltk #nltk:単語を原型にするのに使う↓
nltk.download("wordnet")#←初回のみ使います
from nltk.stem import WordNetLemmatizer
from flask import Flask,render_template,request,url_for,redirect,send_file
from sqlite3 import connect,Row
import sqlite3
import os
from openpyxl import Workbook
from io import BytesIO

# ==========================================
# データベース
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE = os.path.join(BASE_DIR, "database.db")
# これで、WindowsでもPythonAnywhereでも app.py と同じフォルダの database.db を使うようになります。

def init_db():
    conn = connect(DATABASE)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS spell_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            spell_name TEXT NOT NULL,
            description TEXT,
            translation TEXT,
            word_result TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP 
        )
    """)#DEFAULT CURRENT_TIMESTAMP → データを保存したとき、日時を自動で入れる

    conn.commit()
    conn.close()

#型を作る
app = Flask(__name__)

client = genai.Client(
    api_key=os.environ.get("GEMINI_API_KEY"),
    http_options=types.HttpOptions(timeout=30000)
)

@app.route("/")
def index():
    return render_template("index.html") 

@app.route("/work1_kekka", methods=["POST"])
def work1_kekka():
    namae = request.form.get("username")

    # ==========================================
    # 1. 呪文APIからデータ取得（1回だけ）
    # ==========================================
    try:
        url = "https://hp-api.onrender.com/api/spells"

        res = requests.get(url, timeout=(5, 10))
        res.raise_for_status()

        kekka = res.json()

        if not kekka:
            return "呪文データが取得できませんでした。", 503

    except (requests.RequestException, ValueError) as e:
        app.logger.warning("呪文APIの取得に失敗: %s", e)
        return "呪文データの取得に失敗しました。時間をおいて再度お試しください。", 503

    # ==========================================
    # 2. ランダムに呪文を選ぶ
    # ==========================================
    spell = random.choice(kekka)

    name = spell["name"]
    description = spell["description"]

    # ==========================================
    # 3. 日本語翻訳
    # ==========================================
    if name.lower() == "expecto patronum":

        description = (
            "The Patronus Charm is a powerful projection of hope and "
            "happiness that drives away Dementors; a corporeal Patronus "
            "takes the respective animal form of the caster, while a "
            "non-corporeal appears as a wisp of light."
        )

        honyaku = (
            "「守護霊の呪文（パトローナス・チャーム）」は、"
            "吸魂鬼（ディメンター）を追い払う、希望と幸福の強力な具現化です。"
            "実体を持つ守護霊は術者固有の動物の姿をとりますが、"
            "実体を持たない守護霊は光の筋のような姿で現れます。"
        )

    else:
        try:
            response = client.models.generate_content(
                model="gemini-3.6-flash",
                contents=(
                    "次の英語の文章を自然な日本語に翻訳してください。"
                    "翻訳文だけを出力してください。\n\n"
                    + description
                )
            )

            honyaku = response.text

            if not honyaku:
                honyaku = "日本語訳を取得できませんでした。"

        except Exception as e:
            app.logger.warning("Gemini翻訳に失敗: %s", e)
            honyaku = "日本語訳を取得できませんでした。"
    # ==========================================
    # 4. 英単語を抽出
    # ==========================================
    stop_words = {
        "that", "this", "with", "from", "into",
        "your", "they", "them", "have", "been",
        "will", "were", "when", "where",
        "which", "what"
    }

    kotoba = re.findall(r"[a-zA-Z]+", description)

    # 小文字にして重複を削除
    words = list(dict.fromkeys(
        word.lower() for word in kotoba
    ))

    # 4文字以上・除外単語以外
    words = [
        word for word in words
        if len(word) >= 4 and word not in stop_words
    ]

    # ==========================================
    # 5. 原形化して辞書APIで意味を取得
    # ==========================================
    lemmatizer = WordNetLemmatizer()

    result_lines = []

    # API通信に時間がかかりすぎないよう、最大8単語
    MAX_WORDS = 8

    for word in words[:MAX_WORDS]:

        # 動詞・名詞・形容詞として原形化
        base_list = [
            lemmatizer.lemmatize(word, pos="v"),
            lemmatizer.lemmatize(word, pos="n"),
            lemmatizer.lemmatize(word, pos="a"),
            word
        ]

        # 同じ原形を重複して調べない
        base_list = list(dict.fromkeys(base_list))

        jp = None

        # 1単語につき最大2回だけ問い合わせる
        for base_word in base_list[:2]:

            try:
                url = "https://api.excelapi.org/dictionary/enja"

                res = requests.get(
                    url,
                    params={"word": base_word},
                    timeout=(3, 4)
                )

                res.raise_for_status()

                if not res.text.strip():
                    continue

                jp = res.text.split("/")[0].strip()

                if jp:
                    break

            except requests.RequestException as e:
                app.logger.warning(
                    "辞書APIの取得に失敗 (%s): %s",
                    base_word,
                    e
                )
                continue

        if jp:
            result_lines.append(f"{word} ： {jp}")
        else:
            result_lines.append(f"{word} ： 意味を取得できませんでした")

    result = "\n".join(result_lines)

    # ==========================================
    # 6. 結果画面
    # ==========================================
    return render_template(
        "work1_kekka.html",
        val=[namae, name, description, honyaku, result],
        spell_type=name
    )

# ==========================================
# 呪文をDBへ保存
# ==========================================

@app.route("/save_spell", methods=["POST"])
def save_spell():

    username = request.form.get("username")
    spell_name = request.form.get("spell_name")
    description = request.form.get("description")
    translation = request.form.get("translation")
    word_result = request.form.get("word_result")

    conn = connect(DATABASE)

    try:
        conn.execute("""
            INSERT INTO spell_history
            (username, spell_name, description, translation, word_result)
            VALUES (?, ?, ?, ?, ?)
        """, (
            username,
            spell_name,
            description,
            translation,
            word_result
        ))
        conn.commit()
    except sqlite3.IntegrityError:
        # すでに同じデータ（UNIQUE制約違反）が存在する場合の処理
        conn.rollback()
        # 必要に応じてフラッシュメッセージなどを追加可能
    finally:
        conn.close()

    # 保存後、履歴ページへ移動
    return redirect(url_for("history"))

# ==========================================
# 履歴を見る
# ==========================================

@app.route("/history")
def history():

    conn = connect(DATABASE)

    conn.row_factory = Row

    spells = conn.execute("""
        SELECT *
        FROM spell_history
        ORDER BY created_at DESC
    """).fetchall()

    conn.close()

    return render_template(
        "history.html",
        spells=spells
    )

# エクセルにまとめてダウンロードさせる
@app.route("/export_excel")
def export_excel():

    conn = connect(DATABASE)
    conn.row_factory = Row

    spells = conn.execute("""
        SELECT spell_name, description,
               translation, word_result
        FROM spell_history
        ORDER BY created_at DESC
    """).fetchall()

    conn.close()

    # Excelを作成
    wb = Workbook()
    ws = wb.active
    ws.title = "呪文履歴"

    # 見出し
    ws.append([
        "呪文名",
        "英語説明",
        "日本語訳",
        "単語・意味",
    ])

    # データを追加
    for spell in spells:
        ws.append([
            spell["spell_name"],
            spell["description"],
            spell["translation"],
            spell["word_result"],
        ])

    # 列幅を調整
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 80

    # メモリ上にExcelを保存
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="spell_history.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================================
# アプリ起動時にDBを準備
# ==========================================

init_db()


if __name__ == "__main__":
    app.run(debug=True)









# #Flask型のメソッドを使ってルール付け
# @app.route("/") #/はTOPページのこと
# def index():
#     li=["大吉","中吉","小吉","凶"]
#     youbi=["月","火","水","木","金","土","日"]
#     kekka=random.choice(li)
#     date=date.today()
#     youbi2=youbi[date.weekday()]
#     #関数名はなんでもOK(後で使うのでわかりやすく)
#     #returnでrender?templateを使いどのhtmlを使うか指定
#     return render_template("index.html",val1=kekka,val2=date,val3=youbi2)

# @app.route("/prof")#スラッシュ忘れずに
# def prof():
#     return render_template("prof.html")

# @app.route("/work1")
# def work1():
#     return render_template("work1.html")

# @app.route("/work2")
# def work2():
#     return render_template("work2.html")

# @app.route("/work3")
# def work3():
#     return render_template("work3.html")

# @app.route("/work4")
# def work4():
#     return render_template("work4.html")

# @app.route("/work5")
# def work5():
#     return render_template("work5.html")

# @app.route("/work1_kekka",methods=["POST"])#データの送信: URLではなく、リクエストの本体（ボディ）にデータを入れて送信します。フォームから入力してもらったらPOSTがないとエラーに
# def work1_kekka():
#     li=["大吉","中吉","小吉","凶"]
#     kekka=random.choice(li)
#     name = request.form.get("username")
#     myouji = request.form.get("myouji")
#     return render_template("work1_kekka.html",val1=name,val2=myouji,val3=kekka)

# @app.route("/work2_kekka")
# def work2_kekka():
#     date = request.args.get('id','両方')#両方はデフォルト値
#     if date =="犬":
#         path = 'dog.jpg'
#     elif date =="猫":
#         path = 'cat.jpg'
#     else:
#         path = 'animal.jpg'
#     return render_template("work2_kekka.html",val1=date,val2=path)

# @app.route("/work3_kekka",methods=["POST"])
# def work3_kekka():
#     bangou = request.form.get("nyuuryoku") 
#     url=("https://api.openbd.jp/v1/get?isbn="+bangou)
#     res = requests.get(url)
#     kekka = json.loads(res.text)
#     title = kekka[0]["summary"]["title"]
#     pub = kekka[0]["summary"]["publisher"]
#     tyosya = kekka[0]["summary"]["author"]
#     return render_template("work3_kekka.html",val1=[title,pub,tyosya])

# @app.route("/work4_kekka")
# def work4_kekka():
#     tokyo="https://www.jma.go.jp/bosai/forecast/data/forecast/130000.json"
#     tiba="https://www.jma.go.jp/bosai/forecast/data/forecast/120000.json"
#     saitama="https://www.jma.go.jp/bosai/forecast/data/forecast/110000.json"
#     date = request.args.get('id','天気')
#     if date == "東京":
#         url=tokyo
#     elif date == "千葉":
#         url=tiba
#     else:
#         url=saitama
#     res = requests.get(url)
#     kekka = json.loads(res.text)

#     times = kekka[0]["timeSeries"][0]["timeDefines"]
#     areas = kekka[0]["timeSeries"][0]["areas"]

#     weather_list = []

#     for area in areas:
#         for j in range(len(times)):
#             weather_list.append({
#                 "date": times[j][:10],
#                 "area": area["area"]["name"],
#                 "code": area["weatherCodes"][j],
#                 "weather": area["weathers"][j],
#                 "gazou": f"https://tpf.weathernews.jp/wxicon/152/{area["weatherCodes"][j]}.png"
#             })

#     return render_template("work4_kekka.html",list=weather_list)

# @app.route("/work5_kekka",methods=["POST"])
# def work5_kekka():
#      youto = request.form.get("youto")
#      price = request.form.get("price")     
#      today = str(date.today())
#      con = connect("kakeibo.db")
#      #セキュリテイ上は良くない
#      # out_str = f"INSERT INTO data (youto, price, date) VALUES ({youto}, {price}, '{today}');"
#      # cur = con.execute(out_str)

#      #推奨の書き方
#      out_str = "INSERT INTO data (youto, price, date) VALUES (?, ?, ?)"
#      cur = con.execute(out_str, (youto, price, str(today)))
#      con.commit()
#      con.close()
#      return render_template("work5_kekka.html",val1=cur)



# #ルール付けが終わったら実行（一番最後）
# if __name__ =='__main__':
#     app.run(debug=True)
#     #公開の時はFalseがおすすめ