# Running in production (internal network)

This dashboard is intended for use on your own trusted office network or VPN —
**not exposed to the public internet.** It's protected by HTTP Basic Auth, which
sends credentials on every request; that's fine on a trusted LAN, but if you
ever need to reach it from outside the office, put it behind a reverse proxy
that terminates TLS (HTTPS) first, rather than exposing it directly.

## 1. Configure `.env`

Create/edit `.env` in the project root with:

```
GEMINI_API_KEY=your-gemini-api-key
ADMIN_USERNAME=admin
ADMIN_PASSWORD=a-strong-password
```

`.env` is gitignored — never commit it. The server refuses to start serving
pages if `ADMIN_USERNAME`/`ADMIN_PASSWORD` aren't set (fails closed, not open).

## 2. Install dependencies

```
pip install -r requirements.txt
```

## 3. Run it

```
python app.py
```

This serves the dashboard, the `data/*.xlsx` workbooks, and the AI bill
extraction APIs all from one process, on `http://0.0.0.0:5000`, via
`waitress` (a production WSGI server — not Flask's dev server). Reach it from
any machine on the same network at `http://<this-machine's-LAN-IP>:5000`, and
sign in with the username/password from `.env`.

To use a different port, set `PORT` in `.env` or the environment before running.

## What changed from the dev setup

- The Gemini API key now lives only in `.env`, read server-side — it is never
  sent to the browser. The two former separate dev servers
  (`extract_purchase_bill.py` on :5000, `extract_sales_bill.py` on :5001) were
  merged into this single `app.py`.
- Every route (the dashboard page, the `data/` files, and the extraction APIs)
  requires HTTP Basic Auth.
