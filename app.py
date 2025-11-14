
from flask import Flask, render_template, request, redirect, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, timedelta
import os, csv
from openpyxl import Workbook

app = Flask(__name__)

db_path = os.path.join(os.path.dirname(__file__), "database.db")
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

class Client(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50))
    house_address = db.Column(db.String(200))
    register_address = db.Column(db.String(200))
    first_contact = db.Column(db.String(20))
    next_follow = db.Column(db.String(20))
    notes = db.Column(db.Text)

with app.app_context():
    db.create_all()

@app.route("/")
def home():
    return redirect("/list")

@app.route("/add", methods=["GET", "POST"])
def add_client():
    if request.method == "POST":
        now = datetime.now().strftime("%Y-%m-%d")

        chosen_next = request.form["next_follow"]
        if chosen_next.strip() == "":
            next_time = (datetime.now() + timedelta(days=14)).strftime("%Y-%m-%d")
        else:
            next_time = chosen_next

        c = Client(
            name=request.form["name"],
            house_address=request.form["house_address"],
            register_address=request.form["register_address"],
            first_contact=now,
            next_follow=next_time,
            notes=request.form["notes"]
        )

        db.session.add(c)
        db.session.commit()
        return redirect("/list")

    return render_template("add.html")

@app.route("/edit/<int:id>", methods=["GET", "POST"])
def edit_client(id):
    client = Client.query.get(id)

    if request.method == "POST":
        client.name = request.form["name"]
        client.house_address = request.form["house_address"]
        client.register_address = request.form["register_address"]
        client.notes = request.form["notes"]

        chosen_next = request.form["next_follow"].strip()
        if chosen_next != "":
            client.next_follow = chosen_next

        db.session.commit()
        return redirect("/list")

    return render_template("edit.html", client=client)

@app.route("/export_csv")
def export_csv():
    filename = "clients_export.csv"
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["客戶姓名", "房屋地址", "戶籍地址", "第一次開發", "下次跟進", "備註"])
        for c in Client.query.all():
            writer.writerow([c.name, c.house_address, c.register_address, c.first_contact, c.next_follow, c.notes])
    return send_file(filename, as_attachment=True)

@app.route("/export_excel")
def export_excel():
    filename = "clients_export.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "客戶資料"
    ws.append(["客戶姓名", "房屋地址", "戶籍地址", "第一次開發", "下次跟進", "備註"])
    for c in Client.query.all():
        ws.append([c.name, c.house_address, c.register_address, c.first_contact, c.next_follow, c.notes])
    wb.save(filename)
    return send_file(filename, as_attachment=True)

@app.route("/list")
def list_clients():
    search = request.args.get("q", "").strip()
    today_flag = request.args.get("today", "") == "1"
    today_str = datetime.now().strftime("%Y-%m-%d")

    query = Client.query

    if search:
        keyword = f"%{search}%"
        query = query.filter(
            (Client.name.like(keyword)) |
            (Client.house_address.like(keyword)) |
            (Client.register_address.like(keyword)) |
            (Client.notes.like(keyword))
        )

    clients = query.all()

    if today_flag:
        clients = [c for c in clients if c.next_follow == today_str]

    today_count = sum(1 for c in Client.query.all() if c.next_follow == today_str)

    return render_template("list.html",
        clients=clients, search=search,
        today_flag=today_flag, today_str=today_str, today_count=today_count )

@app.route("/delete/<int:id>")
def delete_client(id):
    c = Client.query.get(id)
    db.session.delete(c)
    db.session.commit()
    return redirect("/list")

if __name__ == "__main__":
    app.run(debug=True)
