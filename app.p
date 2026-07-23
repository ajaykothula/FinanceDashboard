from flask import Flask, render_template, request, redirect, send_file, session
from openpyxl import Workbook, load_workbook
import sqlite3

app = Flask(__name__)
app.secret_key = "finance_dashboard_secret"
DB = "database/finance.db"

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        conn = sqlite3.connect(DB)
        cur = conn.cursor()

        cur.execute(
            "SELECT * FROM users WHERE username=? AND password=?",
            (username, password)
        )

        user = cur.fetchone()
        conn.close()

        if user:
            session["user"] = username
            return redirect("/")

        return "❌ Invalid username or password"

    return render_template("login.html")

@app.route("/signup", methods=["GET", "POST"])
def signup():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        conn = sqlite3.connect(DB)
        cur = conn.cursor()

        try:
            cur.execute(
                "INSERT INTO users(username,password) VALUES(?,?)",
                (username, password)
            )
            conn.commit()
            conn.close()

            return redirect("/login")

        except sqlite3.IntegrityError:
            conn.close()
            return "Username already exists!"

    return render_template("signup.html")
@app.route("/logout")
def logout():

    session.clear()

    return redirect("/login")
def get_conn():
    return sqlite3.connect(DB)

@app.route("/")
def home():
    if "user" not in session:
        return redirect("/login")

    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    cur.execute("SELECT IFNULL(SUM(amount),0) FROM transactions WHERE type='Income'")
    income = cur.fetchone()[0]

    cur.execute("SELECT IFNULL(SUM(amount),0) FROM transactions WHERE type='Expense'")
    expense = cur.fetchone()[0]

    balance = income - expense

    cur.execute("SELECT COUNT(*) FROM transactions")
    count = cur.fetchone()[0]


    # ⭐ ADD BUDGET CODE HERE

    cur.execute("SELECT monthly_budget FROM budget LIMIT 1")
    row = cur.fetchone()

    if row:
        monthly_budget = row[0]
    else:
        monthly_budget = 0

    remaining = monthly_budget - expense
    if monthly_budget > 0:
        budget_percent = (expense / monthly_budget) * 100
    else:
         budget_percent = 0

    # Get transactions
    # Search transactions
    # Search + Date Filter
    search = request.args.get("search", "")
    from_date = request.args.get("from_date", "")
    to_date = request.args.get("to_date", "")

    query = "SELECT * FROM transactions WHERE 1=1"
    params = []

    if search:
       query += " AND category LIKE ?"
       params.append("%" + search + "%")

    if from_date:
       query += " AND date >= ?"
       params.append(from_date)

    if to_date:
       query += " AND date <= ?"
       params.append(to_date)

    query += " ORDER BY id"

    cur.execute(query, params)
    transactions = cur.fetchall()
    # Monthly chart data
    # Monthly chart data
    cur.execute("""
        SELECT
            strftime('%Y-%m', date),
            SUM(CASE WHEN type='Income' THEN amount ELSE 0 END),
            SUM(CASE WHEN type='Expense' THEN amount ELSE 0 END)
        FROM transactions
        GROUP BY strftime('%Y-%m', date)
        ORDER BY strftime('%Y-%m', date)
    """)

    monthly_data = cur.fetchall()
    cur.execute("""
         SELECT category, SUM(amount)
           FROM transactions
             WHERE type='Expense'
           GROUP BY category
          """)

    chart_data = cur.fetchall()
    cur.execute("""
    SELECT category, SUM(amount)
    FROM transactions
    WHERE type='Expense'
    GROUP BY category
    ORDER BY SUM(amount) DESC
    """)

    category_summary = cur.fetchall()
    conn.close()

    return render_template(
        "index.html",
        transactions=transactions,
        income=income,
        expense=expense,
        balance=balance,
        count=count,
        monthly_budget=monthly_budget,
        remaining=remaining,
        budget_percent=budget_percent,
        monthly_data=monthly_data,
        chart_data=chart_data,
        category_summary=category_summary
    )
@app.route("/add", methods=["POST"])
def add():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO transactions(date,category,type,amount,description)
        VALUES(?,?,?,?,?)
    """, (
        request.form["date"],
        request.form["category"],
        request.form["type"],
        request.form["amount"],
        request.form["description"]
    ))

    conn.commit()
    conn.close()

    return redirect("/")

@app.route("/delete/<int:id>")
def delete(id):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("DELETE FROM transactions WHERE id=?", (id,))

    conn.commit()
    conn.close()

    return redirect("/")

@app.route("/edit/<int:id>", methods=["GET", "POST"])
def edit(id):
    conn = get_conn()
    cur = conn.cursor()

    if request.method == "POST":
        cur.execute("""
            UPDATE transactions
            SET date=?,category=?,type=?,amount=?,description=?
            WHERE id=?
        """, (
            request.form["date"],
            request.form["category"],
            request.form["type"],
            request.form["amount"],
            request.form["description"],
            id
        ))

        conn.commit()
        conn.close()
        return redirect("/")

    cur.execute("SELECT * FROM transactions WHERE id=?", (id,))
    transaction = cur.fetchone()

    conn.close()

    return render_template("edit.html", transaction=transaction)
@app.route("/transactions")
def all_transactions():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT * FROM transactions ORDER BY id DESC")
    transactions = cur.fetchall()

    conn.close()

    return render_template(
        "transactions.html",
        transactions=transactions
    )
@app.route("/export")
def export():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT * FROM transactions")
    rows = cur.fetchall()

    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"

    ws.append([
        "ID",
        "Date",
        "Category",
        "Type",
        "Amount",
        "Description"
    ])
    for row in rows:
        ws.append(row)

    filename = "transactions.xlsx"
    wb.save(filename)

    return send_file(filename, as_attachment=True)
@app.route("/set_budget", methods=["POST"])
def set_budget():

    budget = request.form["budget"]

    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    cur.execute("DELETE FROM budget")

    cur.execute(
        "INSERT INTO budget(monthly_budget) VALUES(?)",
        (budget,)
    )

    conn.commit()
    conn.close()

    return redirect("/")
@app.route("/import", methods=["POST"])
def import_excel():
    file = request.files["file"]

    wb = load_workbook(file)
    ws = wb.active

    conn = get_conn()
    cur = conn.cursor()

    # Skip the header row
    # Skip the header row
    # Read Excel headers
# Read Excel headers
    headers = [cell.value for cell in ws[1]]

    # Existing columns in SQLite
    cur.execute("PRAGMA table_info(transactions)")
    existing = [c[1] for c in cur.fetchall()]

    # Add new columns if needed
    for col in headers:
        if col.lower() != "id" and col not in existing:
            cur.execute(f'ALTER TABLE transactions ADD COLUMN "{col}" TEXT')

    conn.commit()

    # Insert all rows dynamically
    for row in ws.iter_rows(min_row=2, values_only=True):

        values = list(row)

        columns = ",".join(headers)
        placeholders = ",".join(["?"] * len(values))

        cur.execute(
            f"INSERT INTO transactions ({columns}) VALUES ({placeholders})",
            values
        )

    conn.commit()
    conn.close()

    return redirect("/")
@app.route("/data-manager")
def data_manager():
    return render_template("data_manager.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
