"""
Jaiswal Group business dashboard — production server.

Serves dashboard.html, the data/ workbooks, and the AI bill-extraction APIs
(purchase + sales) behind a single Flask app protected by HTTP Basic Auth.

Internal-network use only: Basic Auth sends credentials on every request, so
this must run on a trusted LAN/VPN, or behind a TLS-terminating reverse proxy
if it is ever reached from outside your office network. See PRODUCTION.md.

Usage:
    pip install -r requirements.txt
    python app.py
    # serves on http://0.0.0.0:5000 via waitress (production WSGI server)
"""

import base64
import hmac
import io
import json
import os
import re

from dotenv import load_dotenv
from flask import Flask, Response, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
import openpyxl
import requests

load_dotenv()

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = "gemini-3.6-flash"
GEMINI_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
)

ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

PURCHASE_HEADERS = [
    "Purchaser Invoice No",
    "JG Invoice No",
    "Purchase Date",
    "Supplier Name",
    "Product Category",
    "Item Name",
    "HSN Code",
    "Quantity",
    "Rate (Per Unit)",
    "CGST",
    "SGST",
    "Total Amount",
]

SALES_HEADERS = [
    "Invoice No",
    "Customer Name",
    "Product Name",
    "Product Category",
    "HSN Code",
    "Quantity",
    "Rate",
    "CGST",
    "SGST",
    "Total Amount",
]

PURCHASE_PROMPT = """You are extracting a purchase bill into a strict schema. Read the attached PDF invoice/bill and return ONLY a JSON array (no markdown fences, no commentary) of line-item objects. Each object must have exactly these keys, in this order: "Purchaser Invoice No", "JG Invoice No", "Purchase Date", "Supplier Name", "Product Category", "Item Name", "HSN Code", "Quantity", "Rate (Per Unit)", "CGST", "SGST", "Total Amount".

Map each key to whatever label the bill actually uses for it — bills vary in wording, so match by meaning, not exact text:
- "Purchaser Invoice No" ← the bill's own invoice number, usually labelled "Original Invoice No.", "Invoice No.", or similar.
- "JG Invoice No" ← the internal purchase reference, usually labelled "Purchase No." or similar; if the bill has no such field, leave it "".
- "Purchase Date" ← the invoice/bill date, usually labelled "Purchase Date", "Invoice Date", "Bill Date", or similar. Return it exactly as printed on the bill (do not reformat).
- "Supplier Name" ← the selling/billing company's full name as printed at the top of the bill (e.g. "Adsun Lighting Private Limited West Bengal") — the vendor issuing the invoice, not the buyer.
- "Item Name" ← each row's item/product description, usually under a column labelled "ITEMS" or "Item Description".
- "HSN Code" ← each row's HSN/SAC code, usually under a column labelled "HSN" or "HSN/SAC".
- "Quantity" ← each row's quantity, usually under "QTY." or "Quantity".
- "Rate (Per Unit)" ← each row's unit price, usually under "RATE" or "Rate".
- "Product Category", "CGST", "SGST", "Total Amount" ← these are usually NOT printed on this style of bill. Leave each as an empty string "" unless the bill clearly and explicitly states that exact value — do not calculate, estimate, or infer them.

Rules:
- One object per line item on the bill.
- "Purchaser Invoice No", "JG Invoice No", "Purchase Date" and "Supplier Name" are invoice-level fields — repeat the same value on every line item row from that bill.
- "Quantity" and "Rate (Per Unit)" must be plain numbers (no currency symbols, no commas, no units) when present on the bill; use "" if genuinely absent — do not invent 0.
- Do not invent line items that are not on the bill.
- Read every page of the PDF — a bill may span multiple pages, and every line item across all pages must be included.
Return ONLY the JSON array."""

SALES_PROMPT = """You are extracting a sales bill into a strict schema. Read the attached PDF invoice/bill and return ONLY a JSON array (no markdown fences, no commentary) of line-item objects. Each object must have exactly these keys, in this order: "Invoice No", "Customer Name", "Product Name", "Product Category", "HSN Code", "Quantity", "Rate", "CGST", "SGST", "Total Amount".

Map each key to whatever label the bill actually uses for it — bills vary in wording, so match by meaning, not exact text:
- "Invoice No" ← the bill's own invoice number, usually labelled "Invoice No.", "Invoice Number", or similar.
- "Customer Name" ← the buyer's name/company, usually found under a "Bill To" (or "Billed To" / "Ship To") section.
- "Product Name" ← each row's item/product description, usually under a column labelled "ITEMS" or "Item Description".
- "HSN Code" ← each row's HSN/SAC code, usually under a column labelled "HSN" or "HSN/SAC".
- "Quantity" ← each row's quantity, usually under "QTY." or "Quantity".
- "Rate" ← each row's unit price, usually under "RATE" or "Rate".
- "Product Category", "CGST", "SGST", "Total Amount" ← these are usually NOT printed per line item on this style of bill. Leave each as an empty string "" unless the bill clearly and explicitly states that exact value for that field — do not calculate, estimate, or infer them.

Rules:
- One object per line item on the bill.
- "Invoice No" and "Customer Name" are invoice-level fields — repeat the same value on every line item row from that bill.
- "Quantity" and "Rate" must be plain numbers (no currency symbols, no commas, no units) when present on the bill; use "" if genuinely absent — do not invent 0.
- Do not invent line items that are not on the bill.
- Read every page of the PDF — a bill may span multiple pages, and every line item across all pages must be included.
Return ONLY the JSON array."""

app = Flask(__name__, static_folder=None)
CORS(app)


# ---------------------------------------------------------------- auth ----

def _unauthorized():
    return Response(
        "Authentication required.",
        401,
        {"WWW-Authenticate": 'Basic realm="Jaiswal Group Dashboard"'},
    )


@app.before_request
def require_basic_auth():
    if not ADMIN_USERNAME or not ADMIN_PASSWORD:
        # Fail closed: refuse to serve anything if auth isn't configured,
        # rather than silently running the dashboard open to anyone.
        return Response(
            "Server misconfigured: ADMIN_USERNAME / ADMIN_PASSWORD not set in .env.",
            500,
        )

    auth = request.authorization
    if auth is None:
        return _unauthorized()

    user_ok = hmac.compare_digest(auth.username or "", ADMIN_USERNAME)
    pass_ok = hmac.compare_digest(auth.password or "", ADMIN_PASSWORD)
    if not (user_ok and pass_ok):
        return _unauthorized()


# ------------------------------------------------------- static / data ----

@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "dashboard.html")


@app.route("/data/<path:filename>")
def data_file(filename):
    return send_from_directory(DATA_DIR, filename)


# --------------------------------------------------------- gemini call ----

def call_gemini(pdf_bytes: bytes, prompt: str, headers: list[str]) -> list[dict]:
    if not GEMINI_API_KEY:
        raise RuntimeError("No Gemini API key configured (set GEMINI_API_KEY in .env).")

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": "application/pdf",
                            "data": base64.b64encode(pdf_bytes).decode("ascii"),
                        }
                    },
                ]
            }
        ],
        "generationConfig": {"temperature": 0, "responseMimeType": "application/json"},
    }

    resp = requests.post(GEMINI_URL, params={"key": GEMINI_API_KEY}, json=payload, timeout=120)
    resp.raise_for_status()
    data = resp.json()

    text = "".join(part.get("text", "") for part in data["candidates"][0]["content"]["parts"])
    text = re.sub(r"^```json\s*|^```\s*|```\s*$", "", text.strip())

    rows = json.loads(text)
    if not isinstance(rows, list):
        raise ValueError("Gemini did not return a JSON array")

    return [{h: row.get(h, "") for h in headers} for row in rows]


def rows_to_workbook(rows: list[dict], headers: list[str], sheet_title: str) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])

    for col_cells in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(width + 2, 12), 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _extract_endpoint(prompt, headers, sheet_title, as_excel, filename_suffix):
    if "file" not in request.files:
        return jsonify({"error": "No PDF file uploaded."}), 400

    pdf_file = request.files["file"]
    if not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported."}), 400

    pdf_bytes = pdf_file.read()

    try:
        rows = call_gemini(pdf_bytes, prompt, headers)
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 400
    except requests.HTTPError as e:
        return jsonify({"error": f"Gemini API error: {e.response.text}"}), 502
    except (ValueError, KeyError, json.JSONDecodeError) as e:
        return jsonify({"error": f"Could not parse Gemini's response: {e}"}), 502

    if not rows:
        return jsonify({"error": "Gemini did not find any line items on this bill."}), 422

    if not as_excel:
        return jsonify({"headers": headers, "rows": rows})

    workbook_buf = rows_to_workbook(rows, headers, sheet_title)
    out_name = re.sub(r"\.pdf$", "", pdf_file.filename, flags=re.I) + filename_suffix
    return send_file(
        workbook_buf,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/extract-purchase-bill", methods=["POST"])
def extract_purchase_bill():
    return _extract_endpoint(PURCHASE_PROMPT, PURCHASE_HEADERS, "Purchase Register", False, "")


@app.route("/api/extract-purchase-bill/excel", methods=["POST"])
def extract_purchase_bill_excel():
    return _extract_endpoint(
        PURCHASE_PROMPT, PURCHASE_HEADERS, "Purchase Register", True, "_purchase_register.xlsx"
    )


@app.route("/api/extract-sales-bill", methods=["POST"])
def extract_sales_bill():
    return _extract_endpoint(SALES_PROMPT, SALES_HEADERS, "Sales Register", False, "")


@app.route("/api/extract-sales-bill/excel", methods=["POST"])
def extract_sales_bill_excel():
    return _extract_endpoint(
        SALES_PROMPT, SALES_HEADERS, "Sales Register", True, "_sales_register.xlsx"
    )


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "gemini_key_configured": bool(GEMINI_API_KEY)})


if __name__ == "__main__":
    from waitress import serve

    port = int(os.environ.get("PORT", 5000))
    print(f"Serving Jaiswal Group dashboard on http://0.0.0.0:{port} (internal network only)")
    serve(app, host="0.0.0.0", port=port)
