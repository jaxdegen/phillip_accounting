from flask import Flask, jsonify, send_from_directory, request
import os
from datetime import datetime
import phillip_core

app = Flask(__name__)

@app.route("/")
def home():
    return send_from_directory(".", "index.html")


@app.route("/<path:path>")
def files(path):
    return send_from_directory(".", path)


@app.route("/api/test")
def test():
    return jsonify({
        "status": "Phillip backend is running"
    })


@app.route("/api/templates")
def templates():
    return jsonify({
        "templates": phillip_core.template_options
    })

@app.route("/api/spreadsheets")
def spreadsheets():
    folder = os.path.expanduser("~/Documents/Phillip Accounting")

    search_text = request.args.get("search", "").lower()

    if not os.path.exists(folder):
        return jsonify({
            "spreadsheets": []
        })

    files = []

    for filename in os.listdir(folder):
        if filename.endswith(".xlsx"):
            if search_text in filename.lower():
                files.append(filename)

    return jsonify({
        "spreadsheets": files
    })

@app.route("/api/transactions")
def transactions():
    folder = os.path.expanduser("~/Documents/Phillip Accounting")
    search_text = request.args.get("search", "").strip().lower()

    current_path = os.path.join(
        folder,
        "Welcome.xlsx"
    )

    if not search_text:
        return jsonify({
            "transactions": []
        })

    if not os.path.exists(current_path):
        return jsonify({
            "transactions": []
        })

    workbook = phillip_core.load_workbook(
        current_path,
        data_only=True
    )

    sheet = workbook.active
    results = []

    for row_number, row in enumerate(
        sheet.iter_rows(values_only=True),
        start=1
    ):
        row_text = " | ".join(
            str(value)
            for value in row
            if value is not None
        )

        if search_text in row_text.lower():
            results.append({
                "row": row_number,
                "text": row_text
            })

    workbook.close()

    return jsonify({
        "transactions": results
    })

@app.route("/api/spreadsheet/<path:filename>")
def spreadsheet_info(filename):
    folder = os.path.expanduser("~/Documents/Phillip Accounting")
    current_path = os.path.join(folder, filename)

    if not os.path.exists(current_path):
        return jsonify({
            "error": "Spreadsheet not found."
        }), 404

    workbook = phillip_core.load_workbook(
        current_path,
        data_only=True
    )

    sheet = workbook.active
    transaction_count = max(sheet.max_row - 1, 0)

    workbook.close()

    last_updated = datetime.fromtimestamp(
        os.path.getmtime(current_path)
    ).strftime("%B %d, %Y at %I:%M %p")

    return jsonify({
        "filename": filename,
        "last_updated": last_updated,
        "transactions": transaction_count
    })

@app.route("/api/transaction", methods=["POST"])
def transaction():
    data = request.get_json()

    text = data.get("text", "").strip()
    template = data.get("template", "")

    if not text:
        return jsonify({
            "error": "No transaction information provided."
        }), 400

    if template not in phillip_core.template_options:
        return jsonify({
            "error": "Invalid spreadsheet template."
        }), 400

    result = phillip_core.add_transaction(
        transaction=text,
        template=template,
        web=True
    )

    return jsonify({
        "status": "Transaction received",
        "template": template,
        "text": text,
        "result": result
    })

if __name__ == "__main__":
    app.run(debug=True)