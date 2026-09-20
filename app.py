"""
Jaiswal Group business dashboard — production server.

Serves dashboard.html, the data/ workbooks, and the AI bill-extraction APIs
(purchase + sales) behind a single Flask app.

Usage:
    pip install -r requirements.txt
    python app.py
    # serves on http://0.0.0.0:5000 via waitress (production WSGI server)
"""

import base64
import io
import json
import mimetypes
import os
import re
import shutil
import subprocess
import tempfile

from dotenv import load_dotenv
from flask import Flask, Response, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
import openpyxl
import requests

load_dotenv()

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
# Current Google AI models for this key (legacy 1.5 / 2.0 / 2.5 return HTTP 404).
GEMINI_MODEL = "gemini-3.6-flash"
GEMINI_FALLBACKS = (
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-3.7-flash",
    "gemini-3.8-flash",
    "gemini-flash-latest",
)
GEMINI_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
)


def _gemini_model_chain(primary: str | None = None) -> list[str]:
    ordered = [primary or GEMINI_MODEL, *GEMINI_FALLBACKS]
    seen: set[str] = set()
    return [m for m in ordered if m and not (m in seen or seen.add(m))]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "Reports")
ASSETS_DIR = os.path.join(BASE_DIR, "data")

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
    "Billing Date",
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

SALES_PROMPT = """You are extracting a sales bill into a strict schema. Read the attached PDF invoice/bill and return ONLY a JSON array (no markdown fences, no commentary) of line-item objects. Each object must have exactly these keys, in this order: "Invoice No", "Billing Date", "Customer Name", "Product Name", "Product Category", "HSN Code", "Quantity", "Rate", "CGST", "SGST", "Total Amount".

Map each key to whatever label the bill actually uses for it — bills vary in wording, so match by meaning, not exact text:
- "Invoice No" ← the bill's own invoice number, usually labelled "Invoice No.", "Invoice Number", or similar.
- "Billing Date" ← the invoice/bill date, usually labelled "Billing Date", "Invoice Date", "Bill Date", "Date", or similar. Return it exactly as printed on the bill (do not reformat).
- "Customer Name" ← the buyer's name/company, usually found under a "Bill To" (or "Billed To" / "Ship To") section.
- "Product Name" ← each row's item/product description, usually under a column labelled "ITEMS" or "Item Description".
- "HSN Code" ← each row's HSN/SAC code, usually under a column labelled "HSN" or "HSN/SAC".
- "Quantity" ← each row's quantity, usually under "QTY." or "Quantity".
- "Rate" ← each row's unit price, usually under "RATE" or "Rate".
- "Product Category", "CGST", "SGST", "Total Amount" ← these are usually NOT printed per line item on this style of bill. Leave each as an empty string "" unless the bill clearly and explicitly states that exact value for that field — do not calculate, estimate, or infer them.

Rules:
- One object per line item on the bill.
- "Invoice No", "Billing Date", and "Customer Name" are invoice-level fields — repeat the same value on every line item row from that bill.
- "Quantity" and "Rate" must be plain numbers (no currency symbols, no commas, no units) when present on the bill; use "" if genuinely absent — do not invent 0.
- Do not invent line items that are not on the bill.
- Read every page of the PDF — a bill may span multiple pages, and every line item across all pages must be included.
Return ONLY the JSON array."""

app = Flask(__name__, static_folder=None)
CORS(app)


# ------------------------------------------------------- static / data ----

def _safe_read_bytes(path: str) -> bytes:
    """Read a report file even if Excel/OneDrive has a share lock on it."""
    try:
        with open(path, "rb") as f:
            return f.read()
    except PermissionError:
        pass

    # Windows share-lock fallbacks (Excel often blocks Python open + cmd copy)
    tmp_dir = tempfile.mkdtemp(prefix="jg_report_")
    tmp_path = os.path.join(tmp_dir, os.path.basename(path))
    src = os.path.abspath(path)
    dst = os.path.abspath(tmp_path)
    errors: list[str] = []
    try:
        # 1) PowerShell Copy-Item (works more often than cmd copy under Excel locks)
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    f"Copy-Item -LiteralPath '{src}' -Destination '{dst}' -Force",
                ],
                check=True,
                capture_output=True,
                text=True,
            )
            with open(tmp_path, "rb") as f:
                return f.read()
        except Exception as e:
            errors.append(f"powershell: {e}")

        # 2) Win32 CreateFile with full share mode
        try:
            import ctypes
            from ctypes import wintypes

            GENERIC_READ = 0x80000000
            FILE_SHARE_ALL = 0x00000007
            OPEN_EXISTING = 3
            FILE_ATTRIBUTE_NORMAL = 0x80
            INVALID = ctypes.c_void_p(-1).value

            kernel32 = ctypes.windll.kernel32
            handle = kernel32.CreateFileW(
                src,
                GENERIC_READ,
                FILE_SHARE_ALL,
                None,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                None,
            )
            if handle == INVALID:
                raise OSError(f"CreateFileW failed ({ctypes.GetLastError()})")

            try:
                size = wintypes.LARGE_INTEGER()
                if not kernel32.GetFileSizeEx(handle, ctypes.byref(size)):
                    raise OSError("GetFileSizeEx failed")
                nbytes = int(size.value)
                buf = (ctypes.c_char * nbytes)()
                read = wintypes.DWORD()
                offset = 0
                while offset < nbytes:
                    chunk = min(1024 * 1024, nbytes - offset)
                    ok = kernel32.ReadFile(
                        handle,
                        ctypes.byref(buf, offset),
                        chunk,
                        ctypes.byref(read),
                        None,
                    )
                    if not ok or read.value == 0:
                        raise OSError("ReadFile failed")
                    offset += read.value
                return bytes(buf)
            finally:
                kernel32.CloseHandle(handle)
        except Exception as e:
            errors.append(f"win32: {e}")

        raise PermissionError(
            f"Cannot read {os.path.basename(path)} — close it in Excel if open. "
            f"({' | '.join(errors)})"
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.route("/")
def index():
    resp = send_from_directory(BASE_DIR, "dashboard.html")
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    return resp


@app.route("/data/<path:filename>")
def data_file(filename):
    safe = os.path.normpath(filename).replace("\\", "/")
    if safe.startswith("../") or os.path.isabs(safe):
        return jsonify({"error": "Not found"}), 404
    path = os.path.join(DATA_DIR, safe)
    if not os.path.isfile(path):
        return jsonify({"error": "Not found"}), 404
    try:
        data = _safe_read_bytes(path)
    except PermissionError as e:
        return jsonify({"error": str(e)}), 503
    mime = mimetypes.guess_type(path)[0] or "application/octet-stream"
    resp = Response(data, mimetype=mime)
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Content-Disposition"] = f'inline; filename="{os.path.basename(path)}"'
    return resp


@app.route("/assets/<path:filename>")
def assets_file(filename):
    return send_from_directory(ASSETS_DIR, filename)


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

    models = _gemini_model_chain()
    last_err = None
    data = None
    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        resp = requests.post(url, params={"key": GEMINI_API_KEY}, json=payload, timeout=120)
        if resp.status_code >= 400:
            last_err = f"{model}: HTTP {resp.status_code} {resp.text[:220]}"
            continue
        data = resp.json()
        if not (data.get("candidates") or []):
            last_err = f"{model}: no candidates"
            data = None
            continue
        break
    if data is None:
        raise RuntimeError(last_err or "Gemini API error")

    parts = data["candidates"][0].get("content", {}).get("parts", [])
    text = "".join(part.get("text", "") for part in parts if isinstance(part, dict))
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


INSIGHTS_PROMPT = """You are the strategic advisor for Jaiswal Group, an electrical / lighting wholesale business in West Bengal (India).
You receive a JSON snapshot computed from their live Excel registers (sales, purchases, stock, customers/receivables, payments, expenses) covering ALL periods (no FY filter).

Write a short executive brief plus 5 to 7 sharp, actionable board-level insights. Each insight must:
- Be grounded ONLY in the numbers provided (do not invent parties, amounts, or facts).
- Prefer rupee amounts and percentages already in the snapshot.
- Name the concrete next action for the owner this week.
- Use Indian business tone (direct, practical).

Return ONLY a JSON object (no markdown) with exactly these keys:
- "brief": 2–3 sentences executive summary of the business pulse right now
- "cards": array of 5–7 objects, each with:
  - "tag": one of "ok", "warn", "bad"
  - "tagText": short label (e.g. Collections, Supplier risk, Inventory, Demand, Cash, Margin)
  - "title": one-line headline (max ~90 chars)
  - "body": 2–4 sentences of analysis with numbers from the snapshot
  - "act": one concrete next step (max ~120 chars)

Prioritise: cash stuck in receivables, silent customers with dues, supplier concentration, stockouts, sales vs purchase imbalance, payment collection patterns, expense gaps if expenses are empty.
Do not mention that you are an AI. Do not invent expense figures if expenses.total is 0.
Do not invent parties, SKUs, amounts, percentages, or dates that are not present in the snapshot JSON. If a register is empty (e.g. expenses = 0), say so plainly instead of fabricating spend."""


def call_gemini_text(prompt: str, user_payload: dict) -> dict:
    """Text-only Gemini call returning {brief, cards} for Strategic Insights."""
    if not GEMINI_API_KEY:
        raise RuntimeError("No Gemini API key configured (set GEMINI_API_KEY in .env).")

    models = _gemini_model_chain()
    errors = []
    body = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {"text": "BUSINESS SNAPSHOT JSON:\n" + json.dumps(user_payload, ensure_ascii=False)},
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.35,
            "responseMimeType": "application/json",
        },
    }

    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        try:
            resp = requests.post(url, params={"key": GEMINI_API_KEY}, json=body, timeout=90)
            if resp.status_code >= 400:
                # Skip retired models quickly; retry on transient overload
                msg = resp.text[:220].replace("\n", " ")
                errors.append(f"{model}: HTTP {resp.status_code} {msg}")
                if resp.status_code == 404:
                    continue
                if resp.status_code in (429, 503):
                    continue
                continue
            data = resp.json()
            candidates = data.get("candidates") or []
            if not candidates:
                errors.append(f"{model}: no candidates ({str(data)[:180]})")
                continue
            text = "".join(
                part.get("text", "") for part in candidates[0].get("content", {}).get("parts", [])
            )
            text = re.sub(r"^```json\s*|^```\s*|```\s*$", "", text.strip())
            parsed = json.loads(text)

            # Accept either {brief, cards} or a bare cards array (legacy)
            if isinstance(parsed, list):
                brief = ""
                cards = parsed
            elif isinstance(parsed, dict):
                brief = str(parsed.get("brief") or "").strip()
                cards = parsed.get("cards") or []
            else:
                errors.append(f"{model}: unexpected JSON type")
                continue

            if not isinstance(cards, list) or not cards:
                errors.append(f"{model}: empty or non-array cards")
                continue

            cleaned = []
            for c in cards[:8]:
                if not isinstance(c, dict):
                    continue
                tag = str(c.get("tag") or "ok").lower()
                if tag not in ("ok", "warn", "bad"):
                    tag = "warn"
                cleaned.append(
                    {
                        "tag": tag,
                        "tagText": str(c.get("tagText") or "Insight")[:40],
                        "title": str(c.get("title") or "Insight")[:120],
                        "body": str(c.get("body") or "")[:800],
                        "act": str(c.get("act") or "")[:160],
                    }
                )
            if cleaned:
                return {"brief": brief[:600], "cards": cleaned, "model": model}
            errors.append(f"{model}: no usable cards")
        except Exception as e:
            errors.append(f"{model}: {e}")
            continue

    raise RuntimeError(" | ".join(errors[-3:]) if errors else "Gemini insights failed")


@app.route("/api/insights", methods=["POST"])
def api_insights():
    """Generate Strategic Insights from a frontend-built business snapshot via Gemini."""
    payload = request.get_json(silent=True) or {}
    snapshot = payload.get("snapshot") or payload
    if not isinstance(snapshot, dict) or not snapshot:
        return jsonify({"error": "Missing business snapshot."}), 400

    try:
        result = call_gemini_text(INSIGHTS_PROMPT, snapshot)
    except RuntimeError as e:
        return jsonify({"error": str(e), "gemini_key_configured": bool(GEMINI_API_KEY)}), 502
    except Exception as e:
        return jsonify({"error": f"Insights failed: {e}"}), 502

    return jsonify(
        {
            "brief": result.get("brief") or "",
            "cards": result.get("cards") or [],
            "model": result.get("model") or GEMINI_MODEL,
            "source": "gemini",
        }
    )


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "gemini_key_configured": bool(GEMINI_API_KEY)})


if __name__ == "__main__":
    from waitress import serve

    port = int(os.environ.get("PORT", 5000))
    print(f"Serving Jaiswal Group dashboard on http://0.0.0.0:{port} (internal network only)")
    serve(app, host="0.0.0.0", port=port)
