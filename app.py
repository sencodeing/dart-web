from flask import Flask, request, render_template, send_file
from dart_parser import get_table
import pandas as pd
from io import BytesIO

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    table = None
    error = None

    corp = request.form.get("corp", "")
    section = request.form.get("section", "")
    pblntf_ty = request.form.get("pblntf_ty", "A")
    date = request.form.get("date", "")

    if request.method == "POST":
        try:
            df = get_table(corp, section, pblntf_ty, date)
            table = df.to_html(index=False)
        except Exception as e:
            error = str(e)

    return render_template(
        "index.html",
        table=table,
        error=error,
        corp=corp,
        section=section,
        pblntf_ty=pblntf_ty,
        date=date
    )


@app.route("/download", methods=["POST"])
def download():
    corp = request.form.get("corp")
    section = request.form.get("section")
    pblntf_ty = request.form.get("pblntf_ty")
    date = request.form.get("date")

    df = get_table(corp, section, pblntf_ty, date)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    output.seek(0)

    return send_file(output, download_name="result.xlsx", as_attachment=True)


if __name__ == "__main__":
    app.run()