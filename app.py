import copy
import io
import json
import os
import re
import smtplib
import threading
import urllib.error
import urllib.request
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from functools import wraps

import pandas as pd
import plotly.graph_objects as go
from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from flask_wtf.csrf import CSRFProtect
from fpdf import FPDF
from werkzeug.security import check_password_hash, generate_password_hash

import data as cfg

app = Flask(__name__)
app.config["SECRET_KEY"] = cfg.SECRET_KEY
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["WTF_CSRF_TIME_LIMIT"] = 3600

csrf = CSRFProtect(app)
DATA_LOCK = threading.Lock()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CUSTOMERS_FILE = os.path.join(BASE_DIR, "customers.json")
LOCAL_STORE_FILE = os.path.join(BASE_DIR, "local_store.json")
ORDERS_EXCEL_FILE = os.path.join(BASE_DIR, "orders.xlsx")
ORDERS_CSV_FILE = os.path.join(BASE_DIR, "orders_backup.csv")
REVIEWS_CSV_FILE = os.path.join(BASE_DIR, "reviews_backup.csv")
INVENTORY_CSV_FILE = os.path.join(BASE_DIR, "inventory_backup.csv")


# ==============================================================================
# 0. AUTOMATIC ORDER CONFIRMATION EMAIL ENGINE (BREVO API + SMTP FALLBACK)
# ==============================================================================
def send_order_confirmation_email(order: dict):
    sender_email = os.getenv("SENDER_EMAIL", "").strip()
    brevo_key = os.getenv("BREVO_API_KEY", "").strip()
    smtp_login = os.getenv("BREVO_SMTP_LOGIN", "").strip() or sender_email

    if not sender_email or not brevo_key:
        return

    to_email = str(order.get("customerEmail", "")).strip()
    to_name = str(order.get("customerName", "Reader")).strip()
    if not to_email:
        return

    items = order.get("items", [])
    grand_total = round(sum(float(i["totalAmount"]) for i in items), 2)
    rows_html = "".join(
        f"""
        <tr>
          <td style="padding:10px;border-bottom:1px solid #e6dfd5;font-size:13px;color:#1b2a4a;">
            <b>{i['bookTitle']}</b><br><span style="font-size:11px;color:#64748b;">{i['author']}</span>
          </td>
          <td style="padding:10px;border-bottom:1px solid #e6dfd5;font-size:13px;text-align:center;">{i['quantity']}</td>
          <td style="padding:10px;border-bottom:1px solid #e6dfd5;font-size:13px;text-align:right;font-weight:700;">&#8377;{i['totalAmount']:,.2f}</td>
        </tr>
        """
        for i in items
    )

    subject = f"Order Confirmed ({order['orderId']}) — BookVerse"
    html_body = f"""
    <div style="font-family:Inter,Arial,sans-serif;background:#faf7f2;padding:24px;">
      <div style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #e6dfd5;border-radius:16px;overflow:hidden;">
        <div style="background:#1b2a4a;color:#ffffff;padding:22px 26px;">
          <span style="font-size:11px;color:#d9a05b;text-transform:uppercase;font-weight:700;">&#10003; Order Confirmed</span>
          <h1 style="margin:8px 0 4px;font-size:22px;">Thank you for ordering from BookVerse!</h1>
          <p style="margin:0;font-size:12px;color:#e6dfd5;">Order #{order['orderId']} &bull; {order['date']}</p>
        </div>
        <div style="padding:24px 26px;">
          <p style="font-size:14px;color:#1b2a4a;">Hi <b>{to_name}</b>, your books are being prepared for dispatch to <b>{order['city']}, {order['state']}</b>.</p>
          <table style="width:100%;border-collapse:collapse;margin:16px 0;">
            <thead>
              <tr style="background:#faf7f2;font-size:11px;text-transform:uppercase;color:#1b2a4a;">
                <th style="padding:8px;text-align:left;">Book</th>
                <th style="padding:8px;text-align:center;">Qty</th>
                <th style="padding:8px;text-align:right;">Amount</th>
              </tr>
            </thead>
            <tbody>{rows_html}</tbody>
          </table>
          <div style="background:#1b2a4a;color:#ffffff;padding:12px 16px;border-radius:10px;font-size:14px;font-weight:700;text-align:right;">
            Total ({order['paymentMethod']}): &#8377;{grand_total:,.2f}
          </div>
        </div>
      </div>
    </div>
    """

    if brevo_key.startswith("xkeysib-"):
        payload = {
            "sender": {"name": "BookVerse", "email": sender_email},
            "to": [{"email": to_email, "name": to_name}],
            "subject": subject,
            "htmlContent": html_body,
        }
        req = urllib.request.Request(
            "https://api.brevo.com/v3/smtp/email",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "accept": "application/json",
                "api-key": brevo_key,
                "content-type": "application/json",
                "User-Agent": "BookVerse-Store/1.0",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=10):
                return
        except Exception as e:
            print(f"Brevo API email skipped: {e}")
            return

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = f"BookVerse <{sender_email}>"
        msg["To"] = to_email
        msg.attach(MIMEText(html_body, "html"))
        with smtplib.SMTP("smtp-relay.brevo.com", 587, timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_login, brevo_key)
            server.sendmail(sender_email, to_email, msg.as_string())
    except Exception as e:
        print(f"Brevo SMTP email skipped: {e}")


# ==============================================================================
# 1. FIREBASE REALTIME DATABASE REST API (WITH OFFLINE LOCAL FALLBACK)
# ==============================================================================
def firebase_enabled() -> bool:
    return bool(cfg.FIREBASE_DB_URL)


def _normalize_fb_key(key: str) -> str:
    return re.sub(r"[.$#\[\]/]", "_", str(key).strip())


def firebase_request(path: str, method: str = "GET", payload=None, timeout: int = 5):
    if not firebase_enabled():
        return None
    url = f"{cfg.FIREBASE_DB_URL}/{path.lstrip('/')}.json"
    body = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=body, method=method)
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else None
    except (urllib.error.URLError, TimeoutError, ValueError, OSError):
        return None


def sync_customer(customer: dict):
    if not firebase_enabled():
        return False
    safe_customer = {k: v for k, v in customer.items() if k != "password"}
    key = _normalize_fb_key(safe_customer["customerId"])
    return firebase_request(f"customers/{key}", "PUT", safe_customer) is not None


def sync_order(order: dict):
    if not firebase_enabled():
        return False
    key = _normalize_fb_key(order["orderId"])
    order_header = {k: v for k, v in order.items() if k != "items"}
    ok1 = firebase_request(f"orders/{key}", "PUT", order_header) is not None
    items_payload = {
        f"{key}_{idx + 1}": {**item, "orderId": order["orderId"], "date": order["date"]}
        for idx, item in enumerate(order.get("items", []))
    }
    ok2 = firebase_request(f"order_items/{key}", "PUT", items_payload) is not None
    return ok1 and ok2


def sync_review(review: dict):
    if not firebase_enabled():
        return False
    key = _normalize_fb_key(review["reviewId"])
    return firebase_request(f"reviews/{key}", "PUT", review) is not None


def sync_inventory(books: list):
    if not firebase_enabled():
        return False
    payload = {_normalize_fb_key(b["id"]): b for b in books}
    return firebase_request("inventory", "PUT", payload) is not None


# ==============================================================================
# 2. LOCAL PERSISTENCE & INITIALIZATION
# ==============================================================================
def _load_json(path: str, default):
    if not os.path.exists(path):
        return copy.deepcopy(default)
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return copy.deepcopy(default)


def _save_json(path: str, payload):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)


def get_store() -> dict:
    default_store = {
        "books": copy.deepcopy(cfg.INITIAL_BOOKS),
        "orders": cfg.get_demo_orders(),
        "reviews": cfg.get_demo_reviews(),
    }
    store = _load_json(LOCAL_STORE_FILE, default_store)
    if not store.get("books"):
        store["books"] = copy.deepcopy(cfg.INITIAL_BOOKS)
    return store


def save_store(store: dict):
    _save_json(LOCAL_STORE_FILE, store)


def get_customers() -> list:
    return _load_json(CUSTOMERS_FILE, cfg.get_demo_customers())


def save_customers(customers: list):
    _save_json(CUSTOMERS_FILE, customers)


def _recalculate_book_ratings(store: dict):
    reviews = store.get("reviews", [])
    by_book = {}
    for rev in reviews:
        by_book.setdefault(rev["bookId"], []).append(float(rev["rating"]))
    for book in store["books"]:
        if book["id"] in by_book:
            vals = by_book[book["id"]]
            book["rating"] = round(sum(vals) / len(vals), 1)


# ==============================================================================
# 3. PANDAS ETL, DATA CLEANING & AUTOMATED 8-SHEET EXCEL WORKBOOK
# ==============================================================================
def clean_orders_df(orders: list) -> tuple[pd.DataFrame, pd.DataFrame]:
    header_rows = []
    item_rows = []
    for o in orders:
        items = o.get("items", [])
        order_rev = sum(float(i.get("unitPrice", 0)) * int(i.get("quantity", 0)) for i in items)
        order_cost = sum(float(i.get("costPrice", 0)) * int(i.get("quantity", 0)) for i in items)
        order_profit = order_rev - order_cost
        order_qty = sum(int(i.get("quantity", 0)) for i in items)
        margin = (order_profit / order_rev * 100.0) if order_rev > 0 else 0.0

        header_rows.append(
            {
                "orderId": str(o.get("orderId", "")).strip(),
                "date": str(o.get("date", "")),
                "customerId": str(o.get("customerId", "")),
                "customerName": str(o.get("customerName", "")).strip(),
                "customerEmail": str(o.get("customerEmail", "")).strip().lower(),
                "customerPhone": str(o.get("customerPhone", "")).strip(),
                "city": str(o.get("city", "Unknown")).strip() or "Unknown",
                "state": str(o.get("state", "Unknown")).strip() or "Unknown",
                "totalQuantity": order_qty,
                "totalAmount": round(order_rev, 2),
                "totalCost": round(order_cost, 2),
                "profit": round(order_profit, 2),
                "profitMarginPct": round(margin, 2),
                "paymentMethod": str(o.get("paymentMethod", "UPI")),
                "paymentStatus": str(o.get("paymentStatus", "Pending")),
                "orderStatus": str(o.get("orderStatus", "Confirmed")),
                "deliveryAddress": str(o.get("deliveryAddress", "")),
            }
        )
        for i in items:
            qty = max(0, int(i.get("quantity", 0)))
            unit_p = max(0.0, float(i.get("unitPrice", 0)))
            cost_p = max(0.0, float(i.get("costPrice", 0)))
            rev = round(unit_p * qty, 2)
            cost = round(cost_p * qty, 2)
            prof = round(rev - cost, 2)
            m_pct = round((prof / rev * 100.0) if rev > 0 else 0.0, 2)
            item_rows.append(
                {
                    "orderId": str(o.get("orderId", "")),
                    "date": str(o.get("date", "")),
                    "customerId": str(o.get("customerId", "")),
                    "customerName": str(o.get("customerName", "")),
                    "customerEmail": str(o.get("customerEmail", "")).lower(),
                    "city": str(o.get("city", "Unknown")),
                    "state": str(o.get("state", "Unknown")),
                    "bookId": str(i.get("bookId", "")),
                    "bookTitle": str(i.get("bookTitle", "")),
                    "author": str(i.get("author", "Unknown")),
                    "category": str(i.get("category", "General")),
                    "quantity": qty,
                    "unitPrice": unit_p,
                    "costPrice": cost_p,
                    "totalAmount": rev,
                    "totalCost": cost,
                    "profit": prof,
                    "profitMarginPct": m_pct,
                    "paymentMethod": str(o.get("paymentMethod", "UPI")),
                    "paymentStatus": str(o.get("paymentStatus", "Pending")),
                    "orderStatus": str(o.get("orderStatus", "Confirmed")),
                }
            )

    orders_df = pd.DataFrame(header_rows)
    items_df = pd.DataFrame(item_rows)

    if not orders_df.empty:
        orders_df["date"] = pd.to_datetime(orders_df["date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
        orders_df = orders_df[orders_df["orderId"] != ""]
    if not items_df.empty:
        items_df["date"] = pd.to_datetime(items_df["date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
        items_df = items_df[items_df["quantity"] > 0]

    return orders_df, items_df


def clean_reviews_df(reviews: list) -> pd.DataFrame:
    if not reviews:
        return pd.DataFrame(
            columns=[
                "reviewId", "bookId", "bookTitle", "customerName", "customerEmail",
                "rating", "reviewTag", "comment", "reviewDate", "verifiedBuyer",
            ]
        )
    df = pd.DataFrame(reviews)
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce").fillna(5).clip(1, 5).astype(int)
    df["customerEmail"] = df["customerEmail"].astype(str).str.strip().str.lower()
    df["reviewDate"] = pd.to_datetime(df["reviewDate"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    df["verifiedBuyer"] = df["verifiedBuyer"].apply(lambda x: "Yes" if str(x).lower() == "yes" else "No")
    return df


def clean_inventory_df(books: list, items_df: pd.DataFrame) -> pd.DataFrame:
    df = pd.DataFrame(books)
    if df.empty:
        return df
    df["price"] = pd.to_numeric(df["price"], errors="coerce").fillna(0.0)
    df["cost_price"] = pd.to_numeric(df["cost_price"], errors="coerce").fillna(0.0)
    df["stock_quantity"] = pd.to_numeric(df["stock_quantity"], errors="coerce").fillna(0).astype(int)
    df["initial_stock"] = pd.to_numeric(df.get("initial_stock", df["stock_quantity"]), errors="coerce").fillna(0).astype(int)

    if not items_df.empty:
        valid_items = items_df[items_df["orderStatus"] != "Cancelled"]
        sold = valid_items.groupby("bookId")["quantity"].sum().to_dict()
    else:
        sold = {}

    df["units_sold"] = df["id"].map(sold).fillna(0).astype(int)
    df["stock_value_cost"] = (df["stock_quantity"] * df["cost_price"]).round(2)
    df["stock_value_retail"] = (df["stock_quantity"] * df["price"]).round(2)
    df["low_stock_flag"] = df["stock_quantity"].apply(
        lambda q: "Out of Stock" if q <= 0 else ("Low Stock" if q <= 5 else "Healthy")
    )
    return df


def clean_customers_df(customers: list) -> pd.DataFrame:
    if not customers:
        return pd.DataFrame(
            columns=["customerId", "fullName", "email", "phone", "city", "gender", "ageGroup", "createdAt"]
        )
    safe = [{k: v for k, v in c.items() if k != "passwordHash"} for c in customers]
    df = pd.DataFrame(safe)
    df["email"] = df["email"].astype(str).str.strip().str.lower()
    df["city"] = df["city"].fillna("Unknown").astype(str).str.strip()
    df["gender"] = df["gender"].fillna("Prefer not to say").astype(str)
    df["ageGroup"] = df["ageGroup"].fillna("26–40").astype(str)
    return df


def run_etl_pipeline():
    """Normalize data with Pandas, recalculate KPIs, write 8-sheet Excel & 3 CSV backups."""
    store = get_store()
    customers = get_customers()

    orders_df, items_df = clean_orders_df(store.get("orders", []))
    reviews_df = clean_reviews_df(store.get("reviews", []))
    inv_df = clean_inventory_df(store.get("books", []), items_df)
    cust_df = clean_customers_df(customers)

    active_items = items_df[items_df["orderStatus"] != "Cancelled"] if not items_df.empty else pd.DataFrame()

    if not active_items.empty:
        book_summary = (
            active_items.groupby(["bookId", "bookTitle", "category"], as_index=False)
            .agg(unitsSold=("quantity", "sum"), revenue=("totalAmount", "sum"), totalCost=("totalCost", "sum"), grossProfit=("profit", "sum"))
        )
        book_summary["profitMarginPct"] = book_summary.apply(
            lambda r: round((r["grossProfit"] / r["revenue"] * 100.0), 2) if r["revenue"] > 0 else 0.0, axis=1
        )
    else:
        book_summary = pd.DataFrame(columns=["bookId", "bookTitle", "category", "unitsSold", "revenue", "totalCost", "grossProfit", "profitMarginPct"])

    if not active_items.empty:
        cat_summary = (
            active_items.groupby("category", as_index=False)
            .agg(unitsSold=("quantity", "sum"), revenue=("totalAmount", "sum"), totalCost=("totalCost", "sum"), grossProfit=("profit", "sum"))
        )
        cat_summary["profitMarginPct"] = cat_summary.apply(
            lambda r: round((r["grossProfit"] / r["revenue"] * 100.0), 2) if r["revenue"] > 0 else 0.0, axis=1
        )
    else:
        cat_summary = pd.DataFrame(columns=["category", "unitsSold", "revenue", "totalCost", "grossProfit", "profitMarginPct"])

    if not active_items.empty:
        author_perf = (
            active_items.groupby("author", as_index=False)
            .agg(unitsSold=("quantity", "sum"), revenue=("totalAmount", "sum"), grossProfit=("profit", "sum"))
            .sort_values(by="revenue", ascending=False)
        )
    else:
        author_perf = pd.DataFrame(columns=["author", "unitsSold", "revenue", "grossProfit"])

    if not cust_df.empty:
        active_orders = orders_df[orders_df["orderStatus"] != "Cancelled"] if not orders_df.empty else pd.DataFrame()
        rev_by_email = active_orders.groupby("customerEmail")["totalAmount"].sum().to_dict() if not active_orders.empty else {}
        cust_df["totalRevenue"] = cust_df["email"].map(rev_by_email).fillna(0.0)
        demo_df = (
            cust_df.groupby(["city", "ageGroup", "gender"], as_index=False)
            .agg(customerCount=("customerId", "count"), totalRevenue=("totalRevenue", "sum"))
        )
    else:
        demo_df = pd.DataFrame(columns=["city", "ageGroup", "gender", "customerCount", "totalRevenue"])

    try:
        with pd.ExcelWriter(ORDERS_EXCEL_FILE, engine="openpyxl") as writer:
            orders_df.to_excel(writer, sheet_name="All_Orders", index=False)
            items_df.to_excel(writer, sheet_name="Order_Items", index=False)
            book_summary.to_excel(writer, sheet_name="Book_Profit_Summary", index=False)
            cat_summary.to_excel(writer, sheet_name="Category_Profit_Summary", index=False)
            author_perf.to_excel(writer, sheet_name="Author_Performance", index=False)
            demo_df.to_excel(writer, sheet_name="Customer_Demographics", index=False)
            reviews_df.to_excel(writer, sheet_name="Customer_Reviews", index=False)
            inv_df.to_excel(writer, sheet_name="Inventory_Summary", index=False)

        orders_df.to_csv(ORDERS_CSV_FILE, index=False)
        reviews_df.to_csv(REVIEWS_CSV_FILE, index=False)
        inv_df.to_csv(INVENTORY_CSV_FILE, index=False)
    except OSError:
        pass


# ==============================================================================
# 4. AUTH DECORATORS & CONTEXT PROCESSORS
# ==============================================================================
def customer_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("customer"):
            flash("Please sign in to continue.", "warning")
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("is_admin"):
            flash("Administrator authentication required.", "error")
            return redirect(url_for("login", tab="admin"))
        return f(*args, **kwargs)
    return decorated


@app.context_processor
def inject_globals():
    cart = session.get("cart", {})
    cart_count = sum(int(q) for q in cart.values()) if isinstance(cart, dict) else 0
    return {
        "app_name": cfg.APP_NAME,
        "app_tagline": cfg.APP_TAGLINE,
        "categories": cfg.CATEGORIES,
        "cart_count": cart_count,
        "current_customer": session.get("customer"),
        "is_admin": session.get("is_admin", False),
        "fallback_cover": cfg.FALLBACK_COVER,
        "firebase_active": firebase_enabled(),
    }


# ==============================================================================
# 5. CUSTOMER STOREFRONT ROUTES
# ==============================================================================
@app.route("/")
def home():
    store = get_store()
    active_books = [b for b in store["books"] if b.get("active", True)]
    featured_books = [b for b in active_books if b.get("featured")][:6]
    trending_books = sorted(active_books, key=lambda b: b.get("rating", 0), reverse=True)[:6]
    recent_books = sorted(active_books, key=lambda b: b.get("publication_year", 0), reverse=True)[:6]

    _, items_df = clean_orders_df(store.get("orders", []))
    if not items_df.empty:
        sold_map = items_df[items_df["orderStatus"] != "Cancelled"].groupby("bookId")["quantity"].sum().to_dict()
    else:
        sold_map = {}
    bestsellers = sorted(active_books, key=lambda b: sold_map.get(b["id"], 0), reverse=True)[:4]

    authors = []
    seen_authors = set()
    for b in active_books:
        if b["author"] not in seen_authors:
            seen_authors.add(b["author"])
            authors.append({"name": b["author"], "category": b["category"], "sample_book": b["title"]})
        if len(authors) >= 6:
            break

    testimonials = [r for r in store.get("reviews", []) if r.get("rating", 0) >= 4][:3]
    return render_template(
        "home.html",
        featured_books=featured_books,
        trending_books=trending_books,
        recent_books=recent_books,
        bestsellers=bestsellers,
        authors=authors,
        testimonials=testimonials,
    )


@app.route("/shop")
def shop():
    store = get_store()
    books = [b for b in store["books"] if b.get("active", True)]

    q = request.args.get("q", "").strip().lower()
    category = request.args.get("category", "").strip()
    publisher = request.args.get("publisher", "").strip()
    min_price = request.args.get("min_price", "").strip()
    max_price = request.args.get("max_price", "").strip()
    min_rating = request.args.get("min_rating", "").strip()
    featured_only = request.args.get("featured", "") == "1"
    sort_by = request.args.get("sort", "popularity")

    publishers = sorted({b["publisher"] for b in books})

    if q:
        books = [
            b for b in books
            if q in b["title"].lower() or q in b["author"].lower() or q in b["isbn"].lower()
        ]
    if category:
        books = [b for b in books if b["category"] == category]
    if publisher:
        books = [b for b in books if b["publisher"] == publisher]
    if min_price:
        try:
            books = [b for b in books if b["price"] >= float(min_price)]
        except ValueError:
            pass
    if max_price:
        try:
            books = [b for b in books if b["price"] <= float(max_price)]
        except ValueError:
            pass
    if min_rating:
        try:
            books = [b for b in books if float(b.get("rating", 0)) >= float(min_rating)]
        except ValueError:
            pass
    if featured_only:
        books = [b for b in books if b.get("featured")]

    if sort_by == "price_asc":
        books.sort(key=lambda b: b["price"])
    elif sort_by == "price_desc":
        books.sort(key=lambda b: b["price"], reverse=True)
    elif sort_by == "alpha":
        books.sort(key=lambda b: b["title"].lower())
    else:
        books.sort(key=lambda b: (b.get("rating", 0), b.get("featured", False)), reverse=True)

    return render_template(
        "shop.html",
        books=books,
        publishers=publishers,
        filters={
            "q": request.args.get("q", ""),
            "category": category,
            "publisher": publisher,
            "min_price": min_price,
            "max_price": max_price,
            "min_rating": min_rating,
            "featured": featured_only,
            "sort": sort_by,
        },
    )


@app.route("/book/<book_id>")
def book_details(book_id):
    store = get_store()
    book = next((b for b in store["books"] if b["id"] == book_id and b.get("active", True)), None)
    if not book:
        abort(404)
    book_reviews = [r for r in store.get("reviews", []) if r["bookId"] == book_id]
    related = [b for b in store["books"] if b["category"] == book["category"] and b["id"] != book_id and b.get("active", True)][:4]
    return render_template("book_details.html", book=book, reviews=book_reviews, related=related, review_tags=cfg.REVIEW_TAGS)


# ==============================================================================
# 6. SHOPPING CART ROUTES
# ==============================================================================
def _build_cart_view():
    store = get_store()
    book_map = {b["id"]: b for b in store["books"] if b.get("active", True)}
    cart = session.get("cart", {})
    items = []
    total = 0.0
    cleaned_cart = {}

    for bid, qty in cart.items():
        book = book_map.get(bid)
        if not book or book["stock_quantity"] <= 0:
            continue
        safe_qty = max(1, min(int(qty), int(book["stock_quantity"])))
        cleaned_cart[bid] = safe_qty
        subtotal = round(book["price"] * safe_qty, 2)
        total += subtotal
        items.append({"book": book, "quantity": safe_qty, "subtotal": subtotal})

    session["cart"] = cleaned_cart
    session.modified = True
    return items, round(total, 2)


@app.route("/cart")
def cart():
    items, total = _build_cart_view()
    return render_template("cart.html", items=items, total=total)


@app.route("/add-to-cart/<book_id>", methods=["POST"])
def add_to_cart(book_id):
    store = get_store()
    book = next((b for b in store["books"] if b["id"] == book_id and b.get("active", True)), None)
    if not book:
        abort(404)
    if book["stock_quantity"] <= 0:
        flash(f"'{book['title']}' is currently out of stock.", "error")
        return redirect(url_for("book_details", book_id=book_id))

    try:
        qty = max(1, int(request.form.get("quantity", 1)))
    except ValueError:
        qty = 1

    cart_data = session.get("cart", {})
    new_qty = cart_data.get(book_id, 0) + qty
    if new_qty > book["stock_quantity"]:
        new_qty = book["stock_quantity"]
        flash(f"Adjusted quantity to available stock ({book['stock_quantity']}).", "warning")
    else:
        flash(f"Added '{book['title']}' to your cart.", "success")

    cart_data[book_id] = new_qty
    session["cart"] = cart_data
    session.modified = True

    if request.form.get("buy_now") == "1":
        return redirect(url_for("checkout"))
    return redirect(request.referrer or url_for("cart"))


@app.route("/update-cart", methods=["POST"])
def update_cart():
    store = get_store()
    book_map = {b["id"]: b for b in store["books"]}
    cart_data = session.get("cart", {})

    for key, val in request.form.items():
        if key.startswith("qty_"):
            bid = key.replace("qty_", "", 1)
            if bid in cart_data and bid in book_map:
                try:
                    qty = int(val)
                except ValueError:
                    qty = 1
                if qty <= 0:
                    cart_data.pop(bid, None)
                else:
                    max_stock = book_map[bid]["stock_quantity"]
                    cart_data[bid] = min(qty, max_stock)

    session["cart"] = cart_data
    session.modified = True
    flash("Shopping cart updated.", "success")
    return redirect(url_for("cart"))


@app.route("/remove-from-cart/<book_id>", methods=["POST"])
def remove_from_cart(book_id):
    cart_data = session.get("cart", {})
    cart_data.pop(book_id, None)
    session["cart"] = cart_data
    session.modified = True
    flash("Item removed from cart.", "info")
    return redirect(url_for("cart"))


@app.route("/clear-cart", methods=["POST"])
def clear_cart():
    session["cart"] = {}
    session.modified = True
    flash("Your cart has been cleared.", "info")
    return redirect(url_for("cart"))


# ==============================================================================
# 7. AUTHENTICATION ROUTES (SEPARATE CUSTOMER & ADMIN SESSIONS)
# ==============================================================================
@app.route("/login")
def login():
    tab = request.args.get("tab", "login")
    return render_template(
        "login.html",
        tab=tab,
        age_groups=cfg.AGE_GROUPS,
        genders=cfg.GENDERS,
        form_data={},
    )


@app.route("/customer-signup", methods=["POST"])
def customer_signup():
    full_name = request.form.get("full_name", "").strip()
    email = request.form.get("email", "").strip().lower()
    phone = request.form.get("phone", "").strip()
    city = request.form.get("city", "").strip()
    gender = request.form.get("gender", "").strip()
    age_group = request.form.get("age_group", "").strip()
    password = request.form.get("password", "")
    confirm_password = request.form.get("confirm_password", "")

    preserved = {
        "full_name": full_name,
        "email": email,
        "phone": phone,
        "city": city,
        "gender": gender,
        "age_group": age_group,
    }

    def render_signup_error(msg):
        flash(msg, "error")
        return render_template(
            "login.html",
            tab="signup",
            age_groups=cfg.AGE_GROUPS,
            genders=cfg.GENDERS,
            form_data=preserved,
        )

    if not all([full_name, email, phone, city, gender, age_group, password]):
        return render_signup_error("All registration fields are required.")
    if not re.fullmatch(r"\d{10}", phone):
        return render_signup_error("Phone number must be exactly 10 digits.")
    if password != confirm_password:
        return render_signup_error("Passwords do not match.")
    if len(password) < 6:
        return render_signup_error("Password must be at least 6 characters.")

    with DATA_LOCK:
        customers = get_customers()
        if any(c["email"].lower() == email for c in customers):
            flash("An account with this email already exists. Please sign in.", "warning")
            return redirect(url_for("login", tab="login"))

        new_cust = {
            "customerId": f"CUST-{len(customers) + 1:04d}",
            "fullName": full_name,
            "email": email,
            "phone": phone,
            "city": city,
            "gender": gender,
            "ageGroup": age_group,
            "passwordHash": generate_password_hash(password),
            "createdAt": datetime.now().strftime("%Y-%m-%d"),
        }
        customers.append(new_cust)
        save_customers(customers)
        run_etl_pipeline()
        sync_customer(new_cust)

    session["customer"] = {k: v for k, v in new_cust.items() if k != "passwordHash"}
    flash(f"Welcome to BookVerse, {full_name}!", "success")
    return redirect(url_for("home"))


@app.route("/customer-login", methods=["POST"])
def customer_login():
    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")
    customers = get_customers()
    user = next((c for c in customers if c["email"].lower() == email), None)

    if not user or not check_password_hash(user["passwordHash"], password):
        flash("Invalid email or password.", "error")
        return redirect(url_for("login", tab="login"))

    session["customer"] = {k: v for k, v in user.items() if k != "passwordHash"}
    flash(f"Welcome back, {user['fullName']}!", "success")
    next_url = request.form.get("next") or url_for("home")
    return redirect(next_url)


@app.route("/admin-login", methods=["POST"])
def admin_login():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    if username == cfg.ADMIN_USERNAME and password == cfg.ADMIN_PASSWORD:
        session["is_admin"] = True
        flash("Signed in to BookVerse BI Administration.", "success")
        return redirect(url_for("admin_dashboard"))
    flash("Invalid administrator credentials.", "error")
    return redirect(url_for("login", tab="admin"))


@app.route("/logout", methods=["POST"])
def logout():
    role = request.form.get("role", "customer")
    if role == "admin":
        session.pop("is_admin", None)
        flash("Administrator signed out.", "info")
    else:
        session.pop("customer", None)
        flash("You have been signed out.", "info")
    return redirect(url_for("home"))


# ==============================================================================
# 8. CHECKOUT, SIMULATED PAYMENTS, RECEIPTS, PDF INVOICE & ORDER HISTORY
# ==============================================================================
@app.route("/checkout")
@customer_required
def checkout():
    items, total = _build_cart_view()
    if not items:
        flash("Your cart is empty. Add books before checking out.", "warning")
        return redirect(url_for("shop"))
    return render_template("checkout.html", items=items, total=total, payment_methods=cfg.PAYMENT_METHODS)


@app.route("/place-order", methods=["POST"])
@customer_required
def place_order():
    customer = session["customer"]
    full_name = request.form.get("customer_name", customer["fullName"]).strip()
    email = request.form.get("email", customer["email"]).strip().lower()
    phone = request.form.get("phone", customer["phone"]).strip()
    address = request.form.get("delivery_address", "").strip()
    city = request.form.get("city", customer["city"]).strip()
    state = request.form.get("state", "").strip()
    pin_code = request.form.get("pin_code", "").strip()
    payment_method = request.form.get("payment_method", "UPI").strip()

    if not all([full_name, email, phone, address, city, state, pin_code]):
        flash("Please complete all delivery details.", "error")
        return redirect(url_for("checkout"))
    if not re.fullmatch(r"\d{10}", phone):
        flash("Phone number must be exactly 10 digits.", "error")
        return redirect(url_for("checkout"))
    if not re.fullmatch(r"\d{6}", pin_code):
        flash("PIN Code must be exactly 6 digits.", "error")
        return redirect(url_for("checkout"))
    if payment_method not in cfg.PAYMENT_METHODS:
        flash("Invalid payment method selected.", "error")
        return redirect(url_for("checkout"))

    with DATA_LOCK:
        store = get_store()
        cart_data = session.get("cart", {})
        if not cart_data:
            flash("Your cart is empty.", "warning")
            return redirect(url_for("shop"))

        book_map = {b["id"]: b for b in store["books"]}
        order_items = []

        for bid, qty in cart_data.items():
            book = book_map.get(bid)
            qty = int(qty)
            if not book or not book.get("active", True):
                flash("One of the books in your cart is no longer available.", "error")
                return redirect(url_for("cart"))
            if qty <= 0 or qty > book["stock_quantity"]:
                flash(f"Insufficient stock for '{book['title']}'. Available: {book['stock_quantity']}.", "error")
                return redirect(url_for("cart"))

            rev = round(book["price"] * qty, 2)
            cost = round(book["cost_price"] * qty, 2)
            order_items.append(
                {
                    "bookId": book["id"],
                    "bookTitle": book["title"],
                    "author": book["author"],
                    "category": book["category"],
                    "quantity": qty,
                    "unitPrice": book["price"],
                    "costPrice": book["cost_price"],
                    "totalAmount": rev,
                    "totalCost": cost,
                    "profit": round(rev - cost, 2),
                }
            )

        for item in order_items:
            book_map[item["bookId"]]["stock_quantity"] -= item["quantity"]

        existing_nums = []
        for o in store["orders"]:
            m = re.search(r"ORD-BV-(\d+)", o.get("orderId", ""))
            if m:
                existing_nums.append(int(m.group(1)))
        next_num = (max(existing_nums) + 1) if existing_nums else 1
        order_id = f"ORD-BV-{next_num:04d}"

        payment_status = "Pending" if payment_method == "Cash on Delivery" else "Paid"

        new_order = {
            "orderId": order_id,
            "date": datetime.now().strftime("%Y-%m-%d"),
            "customerId": customer["customerId"],
            "customerName": full_name,
            "customerEmail": email,
            "customerPhone": phone,
            "city": city,
            "state": state,
            "pinCode": pin_code,
            "deliveryAddress": address,
            "paymentMethod": payment_method,
            "paymentStatus": payment_status,
            "orderStatus": "Confirmed",
            "stockRestored": False,
            "items": order_items,
        }

        store["orders"].append(new_order)
        save_store(store)
        run_etl_pipeline()
        sync_order(new_order)
        sync_inventory(store["books"])
        send_order_confirmation_email(new_order)

    session["cart"] = {}
    session.modified = True
    flash(f"Order {order_id} placed successfully!", "success")
    return redirect(url_for("receipt", order_id=order_id))


@app.route("/receipt/<order_id>")
def receipt(order_id):
    store = get_store()
    order = next((o for o in store["orders"] if o["orderId"] == order_id), None)
    if not order:
        abort(404)
    cust = session.get("customer")
    if not session.get("is_admin") and (not cust or cust["email"].lower() != order["customerEmail"].lower()):
        flash("Sign in with the account that placed this order to view the receipt.", "warning")
        return redirect(url_for("login"))

    total = sum(i["totalAmount"] for i in order["items"])
    return render_template("receipt.html", order=order, total=round(total, 2))


@app.route("/invoice-pdf/<order_id>")
def download_invoice_pdf(order_id):
    store = get_store()
    order = next((o for o in store["orders"] if o["orderId"] == order_id), None)
    if not order:
        abort(404)
    cust = session.get("customer")
    if not session.get("is_admin") and (not cust or cust["email"].lower() != order["customerEmail"].lower()):
        flash("Sign in with the account that placed this order to download the invoice.", "warning")
        return redirect(url_for("login"))

    total = round(sum(float(i["totalAmount"]) for i in order["items"]), 2)
    pdf = FPDF()
    pdf.add_page()

    pdf.set_fill_color(27, 42, 74)
    pdf.rect(0, 0, 210, 36, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_xy(14, 10)
    pdf.cell(120, 10, "BOOKVERSE", ln=False)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(62, 10, "TAX INVOICE", align="R", ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_x(14)
    pdf.cell(120, 6, "Discover Stories. Explore Knowledge. Understand Readers.", ln=False)
    pdf.cell(62, 6, f"Invoice #{order['orderId']}", align="R", ln=True)

    pdf.set_text_color(27, 42, 74)
    pdf.set_xy(14, 44)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(95, 6, "BILLED TO:", ln=False)
    pdf.cell(87, 6, "ORDER DETAILS:", align="R", ln=True)

    pdf.set_font("Helvetica", "", 10)
    pdf.set_x(14)
    pdf.cell(95, 6, order["customerName"], ln=False)
    pdf.cell(87, 6, f"Date: {order['date']}", align="R", ln=True)
    pdf.set_x(14)
    pdf.cell(95, 6, f"Email: {order['customerEmail']}", ln=False)
    pdf.cell(87, 6, f"Payment: {order['paymentMethod']} ({order['paymentStatus']})", align="R", ln=True)
    pdf.set_x(14)
    pdf.cell(95, 6, f"{order['city']}, {order['state']} {order.get('pinCode', '')}", ln=False)
    pdf.cell(87, 6, f"Status: {order['orderStatus']}", align="R", ln=True)

    pdf.ln(10)
    pdf.set_fill_color(250, 247, 242)
    pdf.set_draw_color(230, 223, 213)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_x(14)
    pdf.cell(12, 9, "#", border=1, fill=True, align="C")
    pdf.cell(96, 9, "Book Title & Author", border=1, fill=True)
    pdf.cell(18, 9, "Qty", border=1, fill=True, align="C")
    pdf.cell(26, 9, "Unit Price", border=1, fill=True, align="R")
    pdf.cell(30, 9, "Subtotal", border=1, fill=True, align="R", ln=True)

    pdf.set_font("Helvetica", "", 9)
    for idx, item in enumerate(order["items"], start=1):
        label = f"{item['bookTitle']} - {item['author']}"
        if len(label) > 50:
            label = label[:48] + ".."
        pdf.set_x(14)
        pdf.cell(12, 9, str(idx), border=1, align="C")
        pdf.cell(96, 9, label, border=1)
        pdf.cell(18, 9, str(item["quantity"]), border=1, align="C")
        pdf.cell(26, 9, f"Rs. {item['unitPrice']:,.2f}", border=1, align="R")
        pdf.cell(30, 9, f"Rs. {item['totalAmount']:,.2f}", border=1, align="R", ln=True)

    pdf.ln(4)
    pdf.set_x(126)
    pdf.set_fill_color(27, 42, 74)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(70, 11, f"Grand Total: Rs. {total:,.2f}", border=0, fill=True, align="R", ln=True)

    pdf_bytes = bytes(pdf.output())
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"BookVerse_Invoice_{order['orderId']}.pdf",
    )


@app.route("/my-orders")
@customer_required
def my_orders():
    store = get_store()
    email = session["customer"]["email"].lower()
    orders = [o for o in store["orders"] if o["customerEmail"].lower() == email]
    orders.sort(key=lambda o: o["orderId"], reverse=True)
    for o in orders:
        o["computedTotal"] = round(sum(i["totalAmount"] for i in o["items"]), 2)
        o["computedQty"] = sum(i["quantity"] for i in o["items"])
    return render_template("my_orders.html", orders=orders)


# ==============================================================================
# 9. CUSTOMER REVIEW SYSTEM & VERIFIED READER LOGIC
# ==============================================================================
def _is_verified_buyer(email: str, book_id: str, orders: list) -> bool:
    email = email.strip().lower()
    eligible_statuses = {"Confirmed", "Processing", "Shipped", "Delivered"}
    for o in orders:
        if o.get("customerEmail", "").lower() == email and o.get("orderStatus") in eligible_statuses:
            if any(i["bookId"] == book_id for i in o.get("items", [])):
                return True
    return False


@app.route("/reviews")
def reviews():
    store = get_store()
    all_reviews = list(reversed(store.get("reviews", [])))
    books = [b for b in store["books"] if b.get("active", True)]

    total_reviews = len(all_reviews)
    avg_rating = round(sum(r["rating"] for r in all_reviews) / total_reviews, 2) if total_reviews else 0.0
    star_counts = {s: sum(1 for r in all_reviews if int(r["rating"]) == s) for s in range(5, 0, -1)}

    return render_template(
        "reviews.html",
        reviews=all_reviews,
        books=books,
        review_tags=cfg.REVIEW_TAGS,
        total_reviews=total_reviews,
        avg_rating=avg_rating,
        star_counts=star_counts,
        selected_book=request.args.get("book_id", ""),
    )


@app.route("/submit-review", methods=["POST"])
def submit_review():
    book_id = request.form.get("book_id", "").strip()
    customer_name = request.form.get("customer_name", "").strip()
    customer_email = request.form.get("customer_email", "").strip().lower()
    review_tag = request.form.get("review_tag", "Highly Recommended").strip()
    comment = request.form.get("comment", "").strip()
    try:
        rating = max(1, min(5, int(request.form.get("rating", 5))))
    except ValueError:
        rating = 5

    if not all([book_id, customer_name, customer_email, comment]):
        flash("Please complete all review fields.", "error")
        return redirect(request.referrer or url_for("reviews"))

    with DATA_LOCK:
        store = get_store()
        book = next((b for b in store["books"] if b["id"] == book_id), None)
        if not book:
            flash("Selected book not found.", "error")
            return redirect(url_for("reviews"))

        verified = "Yes" if _is_verified_buyer(customer_email, book_id, store["orders"]) else "No"

        existing = next(
            (r for r in store["reviews"] if r["bookId"] == book_id and r["customerEmail"].lower() == customer_email),
            None,
        )
        if existing:
            existing.update(
                {
                    "customerName": customer_name,
                    "rating": rating,
                    "reviewTag": review_tag,
                    "comment": comment,
                    "reviewDate": datetime.now().strftime("%Y-%m-%d"),
                    "verifiedBuyer": verified,
                }
            )
            saved_rev = existing
            flash("Your existing review for this book has been updated.", "success")
        else:
            saved_rev = {
                "reviewId": f"REV-{len(store['reviews']) + 1:04d}",
                "bookId": book["id"],
                "bookTitle": book["title"],
                "customerName": customer_name,
                "customerEmail": customer_email,
                "rating": rating,
                "reviewTag": review_tag,
                "comment": comment,
                "reviewDate": datetime.now().strftime("%Y-%m-%d"),
                "verifiedBuyer": verified,
            }
            store["reviews"].append(saved_rev)
            flash("Thank you! Your review has been published.", "success")

        _recalculate_book_ratings(store)
        save_store(store)
        run_etl_pipeline()
        sync_review(saved_rev)

    return redirect(request.referrer or url_for("reviews"))


# ==============================================================================
# 10. ADMIN BI DASHBOARD (PUBLISHING ANALYTICS: TREEMAP, HBAR, SCATTER, STOCK)
# ==============================================================================
def _build_bi_chart(dimension: str, metric: str, chart_type: str, store: dict, customers: list) -> str:
    _, items_df = clean_orders_df(store.get("orders", []))
    inv_df = clean_inventory_df(store.get("books", []), items_df)
    cust_df = clean_customers_df(customers)

    book_pub_map = {b["id"]: b.get("publisher", "Independent") for b in store.get("books", [])}
    active_items = items_df[items_df["orderStatus"] != "Cancelled"].copy() if not items_df.empty else pd.DataFrame()
    if not active_items.empty:
        active_items["publisher"] = active_items["bookId"].map(book_pub_map).fillna("Independent")

    metric_col_map = {"units_sold": "quantity", "revenue": "totalAmount", "gross_profit": "profit"}
    metric_label_map = {"units_sold": "Copies Sold", "revenue": "Revenue (₹)", "gross_profit": "Gross Profit (₹)"}
    m_col = metric_col_map.get(metric, "totalAmount")
    m_label = metric_label_map.get(metric, "Revenue (₹)")

    palette = ["#1B2A4A", "#C47D3E", "#2A9D8F", "#D9A05B", "#3B6978", "#E76F51", "#5C5470", "#84A9AC"]
    fig = go.Figure()
    title = "BookVerse Publishing Intelligence"

    # Special Analyst View 1: Price vs. Demand Bubble Scatter
    if dimension == "price_elasticity" and not inv_df.empty:
        active_inv = inv_df[inv_df["active"] == True].copy()
        active_inv["gross_profit"] = (active_inv["price"] - active_inv["cost_price"]) * active_inv["units_sold"]
        fig.add_trace(
            go.Scatter(
                x=active_inv["price"],
                y=active_inv["units_sold"],
                mode="markers",
                text=active_inv["title"],
                marker=dict(
                    size=active_inv["gross_profit"].clip(lower=150) / 35 + 10,
                    color="#C47D3E",
                    opacity=0.85,
                    line=dict(width=1.5, color="#1B2A4A"),
                ),
                hovertemplate="<b>%{text}</b><br>Cover Price: ₹%{x}<br>Copies Sold: %{y}<extra></extra>",
            )
        )
        title = "Price Elasticity Matrix — Cover Price (₹) vs. Copies Sold"
        fig.update_xaxes(title_text="Cover Price (₹)")
        fig.update_yaxes(title_text="Copies Sold")

    # Special Analyst View 2: Stock Available vs. Copies Sold Grouped Comparison
    elif dimension == "stock_vs_sold" and not inv_df.empty:
        grp = inv_df.groupby("category", as_index=False).agg(
            available=("stock_quantity", "sum"),
            sold=("units_sold", "sum"),
        )
        fig.add_trace(go.Bar(name="Available Stock", x=grp["category"], y=grp["available"], marker_color="#1B2A4A"))
        fig.add_trace(go.Bar(name="Copies Sold", x=grp["category"], y=grp["sold"], marker_color="#C47D3E"))
        fig.update_layout(barmode="group")
        title = "Inventory Turnover — Available Stock vs. Copies Sold by Genre"

    # Special Analyst View 3: Reader Demographics (Age Group x Gender Stacked/Grouped)
    elif dimension == "demographics" and not active_items.empty and not cust_df.empty:
        merged = active_items.merge(cust_df[["email", "ageGroup", "gender"]], left_on="customerEmail", right_on="email", how="left")
        merged["ageGroup"] = merged["ageGroup"].fillna("26–40")
        merged["gender"] = merged["gender"].fillna("Other")
        piv = merged.pivot_table(index="ageGroup", columns="gender", values=m_col, aggfunc="sum", fill_value=0)
        for idx, g_col in enumerate(piv.columns):
            fig.add_trace(
                go.Bar(
                    name=str(g_col),
                    x=piv.index.tolist(),
                    y=piv[g_col].tolist(),
                    marker_color=palette[idx % len(palette)],
                )
            )
        fig.update_layout(barmode="group")
        title = f"Reader Cohort Breakdown (Age × Gender) — {m_label}"

    # Standard Dimensions rendered via Treemap / Horizontal Bar / Area / Donut
    else:
        labels, values = [], []
        if dimension == "category" and not active_items.empty:
            grp = active_items.groupby("category")[m_col].sum().sort_values(ascending=False)
            labels, values, title = grp.index.tolist(), grp.values.tolist(), f"Genre Performance — {m_label}"
        elif dimension == "publisher" and not active_items.empty:
            grp = active_items.groupby("publisher")[m_col].sum().sort_values(ascending=False).head(10)
            labels, values, title = grp.index.tolist(), grp.values.tolist(), f"Top Publishing Houses — {m_label}"
        elif dimension == "author" and not active_items.empty:
            grp = active_items.groupby("author")[m_col].sum().sort_values(ascending=False).head(10)
            labels, values, title = grp.index.tolist(), grp.values.tolist(), f"Top Authors — {m_label}"
        elif dimension == "book" and not active_items.empty:
            grp = active_items.groupby("bookTitle")[m_col].sum().sort_values(ascending=False).head(10)
            labels, values, title = grp.index.tolist(), grp.values.tolist(), f"Bestselling Titles — {m_label}"
        elif dimension == "city" and not active_items.empty:
            grp = active_items.groupby("city")[m_col].sum().sort_values(ascending=False)
            labels, values, title = grp.index.tolist(), grp.values.tolist(), f"City Reader Hubs — {m_label}"
        elif dimension == "timeline" and not active_items.empty:
            grp = active_items.groupby("date")[m_col].sum().sort_index()
            labels, values, title = grp.index.tolist(), grp.values.tolist(), f"Sales Velocity Timeline — {m_label}"
        elif dimension == "payment_method" and not active_items.empty:
            grp = active_items.groupby("paymentMethod")[m_col].sum()
            labels, values, title = grp.index.tolist(), grp.values.tolist(), f"Checkout Payment Split — {m_label}"

        if not labels:
            fig.add_annotation(text="No order data available for this filter.", showarrow=False, font=dict(size=14, color="#1B2A4A"))
        elif chart_type == "treemap":
            if dimension == "category" and not active_items.empty:
                ids, t_labels, parents, t_values = [], [], [], []
                cat_tot = active_items.groupby("category")[m_col].sum()
                bk_tot = active_items.groupby(["category", "bookTitle"])[m_col].sum()
                for c_name, c_val in cat_tot.items():
                    if c_val > 0:
                        ids.append(f"CAT::{c_name}")
                        t_labels.append(c_name)
                        parents.append("")
                        t_values.append(float(c_val))
                for (c_name, b_name), b_val in bk_tot.items():
                    if b_val > 0:
                        ids.append(f"BK::{c_name}::{b_name}")
                        t_labels.append(b_name[:22] + "…" if len(b_name) > 22 else b_name)
                        parents.append(f"CAT::{c_name}")
                        t_values.append(float(b_val))
                fig.add_trace(
                    go.Treemap(
                        ids=ids,
                        labels=t_labels,
                        parents=parents,
                        values=t_values,
                        branchvalues="total",
                        textinfo="label+value",
                        marker=dict(colorscale=[[0, "#FAF7F2"], [0.5, "#C47D3E"], [1, "#1B2A4A"]]),
                    )
                )
            else:
                fig.add_trace(
                    go.Treemap(
                        labels=labels,
                        parents=[""] * len(labels),
                        values=values,
                        textinfo="label+value+percent entry",
                        marker=dict(colorscale=[[0, "#FAF7F2"], [0.5, "#C47D3E"], [1, "#1B2A4A"]]),
                    )
                )
        elif chart_type == "donut":
            fig.add_trace(
                go.Pie(
                    labels=labels,
                    values=values,
                    hole=0.55,
                    marker=dict(colors=palette, line=dict(color="#FAF7F2", width=2)),
                    textinfo="percent",
                )
            )
        elif chart_type == "area":
            fig.add_trace(
                go.Scatter(
                    x=labels,
                    y=values,
                    mode="lines+markers",
                    fill="tozeroy",
                    line=dict(color="#1B2A4A", width=3, shape="spline"),
                    marker=dict(color="#C47D3E", size=8),
                    fillcolor="rgba(196, 125, 62, 0.20)",
                )
            )
        else:
            # Horizontal Ranking Bar (hbar) — distinct from vertical bars
            rev_labels = list(reversed(labels))
            rev_values = list(reversed(values))
            bar_colors = [palette[i % len(palette)] for i in range(len(rev_labels))]
            fig.add_trace(
                go.Bar(
                    x=rev_values,
                    y=rev_labels,
                    orientation="h",
                    marker=dict(color=bar_colors),
                    text=[f"₹{int(v):,}" if metric != "units_sold" else f"{int(v)}" for v in rev_values],
                    textposition="auto",
                )
            )

    fig.update_layout(
        title=dict(text=f"<b>{title}</b>", font=dict(size=15, color="#1B2A4A", family="Playfair Display, Georgia, serif")),
        height=410,
        margin=dict(l=20, r=20, t=52, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(250,247,242,0.45)",
        xaxis=dict(automargin=True, gridcolor="#E6DFD5"),
        yaxis=dict(automargin=True, gridcolor="#E6DFD5"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.24, xanchor="center", x=0.5),
    )
    return fig.to_html(full_html=False, include_plotlyjs="cdn", config={"responsive": True, "displayModeBar": False})


@app.route("/admin")
@admin_required
def admin_dashboard():
    store = get_store()
    customers = get_customers()
    orders_df, items_df = clean_orders_df(store.get("orders", []))
    reviews_df = clean_reviews_df(store.get("reviews", []))
    inv_df = clean_inventory_df(store.get("books", []), items_df)

    active_orders = orders_df[orders_df["orderStatus"] != "Cancelled"] if not orders_df.empty else pd.DataFrame()
    active_items = items_df[items_df["orderStatus"] != "Cancelled"] if not items_df.empty else pd.DataFrame()

    total_revenue = round(float(active_orders["totalAmount"].sum()), 2) if not active_orders.empty else 0.0
    gross_profit = round(float(active_orders["profit"].sum()), 2) if not active_orders.empty else 0.0
    margin_pct = round((gross_profit / total_revenue) * 100, 1) if total_revenue > 0 else 0.0
    total_orders = int(len(orders_df))
    active_order_count = int(len(active_orders))
    avg_order_value = round(total_revenue / active_order_count, 0) if active_order_count > 0 else 0
    total_books_sold = int(active_items["quantity"].sum()) if not active_items.empty else 0
    avg_basket_size = round(total_books_sold / active_order_count, 1) if active_order_count > 0 else 0.0
    avg_rating = round(float(reviews_df["rating"].mean()), 2) if not reviews_df.empty else 0.0

    remaining_stock = int(inv_df["stock_quantity"].sum()) if not inv_df.empty else 0
    total_stock_pool = total_books_sold + remaining_stock
    sell_through_pct = round((total_books_sold / total_stock_pool) * 100, 1) if total_stock_pool > 0 else 0.0
    stock_cost_value = round(float(inv_df["stock_value_cost"].sum()), 0) if not inv_df.empty else 0

    if not active_items.empty:
        bestseller = active_items.groupby("bookTitle")["quantity"].sum().idxmax()
        top_genre = active_items.groupby("category")["totalAmount"].sum().idxmax()
    else:
        bestseller = "N/A"
        top_genre = "N/A"

    kpis = {
        "total_revenue": total_revenue,
        "gross_profit": gross_profit,
        "margin_pct": margin_pct,
        "avg_order_value": avg_order_value,
        "total_orders": total_orders,
        "total_books_sold": total_books_sold,
        "avg_basket_size": avg_basket_size,
        "sell_through_pct": sell_through_pct,
        "stock_cost_value": stock_cost_value,
        "bestseller": bestseller,
        "top_genre": top_genre,
        "avg_rating": avg_rating,
        "total_customers": len(customers),
        "active_books": int(inv_df["active"].sum()) if not inv_df.empty else 0,
        "low_stock_books": int((inv_df["stock_quantity"] <= 5).sum()) if not inv_df.empty else 0,
        "pending_orders": int((orders_df["orderStatus"].isin(["Confirmed", "Processing"])).sum()) if not orders_df.empty else 0,
        "completed_orders": int((orders_df["orderStatus"] == "Delivered").sum()) if not orders_df.empty else 0,
    }

    dimension = request.args.get("dimension", "category")
    metric = request.args.get("metric", "revenue")
    chart_type = request.args.get("chart_type", "hbar")
    chart_html = _build_bi_chart(dimension, metric, chart_type, store, customers)

    q = request.args.get("order_q", "").strip().lower()
    f_city = request.args.get("city", "").strip()
    f_status = request.args.get("status", "").strip()
    f_payment = request.args.get("payment", "").strip()

    orders_list = list(reversed(store.get("orders", [])))
    if q:
        orders_list = [
            o for o in orders_list
            if q in o["orderId"].lower() or q in o["customerName"].lower() or q in o["customerEmail"].lower()
        ]
    if f_city:
        orders_list = [o for o in orders_list if o["city"] == f_city]
    if f_status:
        orders_list = [o for o in orders_list if o["orderStatus"] == f_status]
    if f_payment:
        orders_list = [o for o in orders_list if o["paymentMethod"] == f_payment]

    for o in orders_list:
        o["computedTotal"] = round(sum(i["totalAmount"] for i in o["items"]), 2)
        o["computedQty"] = sum(i["quantity"] for i in o["items"])

    cities = sorted({o["city"] for o in store.get("orders", [])})

    rev_q = request.args.get("rev_q", "").strip().lower()
    rev_rating = request.args.get("rev_rating", "").strip()
    reviews_list = list(reversed(store.get("reviews", [])))
    if rev_q:
        reviews_list = [r for r in reviews_list if rev_q in r["bookTitle"].lower() or rev_q in r["customerName"].lower() or rev_q in r["comment"].lower()]
    if rev_rating:
        reviews_list = [r for r in reviews_list if str(r["rating"]) == rev_rating]

    return render_template(
        "admin.html",
        kpis=kpis,
        chart_html=chart_html,
        dimension=dimension,
        metric=metric,
        chart_type=chart_type,
        orders=orders_list,
        reviews=reviews_list,
        cities=cities,
        order_statuses=cfg.ORDER_STATUSES,
        payment_statuses=cfg.PAYMENT_STATUSES,
        payment_methods=cfg.PAYMENT_METHODS,
    )


@app.route("/admin/orders")
@admin_required
def admin_orders():
    return redirect(url_for("admin_dashboard") + "#orders-section")


@app.route("/admin/reviews")
@admin_required
def admin_reviews():
    return redirect(url_for("admin_dashboard") + "#reviews-section")


@app.route("/admin/update-order-status", methods=["POST"])
@admin_required
def admin_update_order_status():
    order_id = request.form.get("order_id", "").strip()
    new_status = request.form.get("order_status", "").strip()
    new_pay_status = request.form.get("payment_status", "").strip()

    if new_status not in cfg.ORDER_STATUSES:
        flash("Invalid order status.", "error")
        return redirect(url_for("admin_dashboard"))

    with DATA_LOCK:
        store = get_store()
        order = next((o for o in store["orders"] if o["orderId"] == order_id), None)
        if not order:
            flash("Order not found.", "error")
            return redirect(url_for("admin_dashboard"))

        book_map = {b["id"]: b for b in store["books"]}
        if new_status == "Cancelled" and not order.get("stockRestored", False):
            for item in order.get("items", []):
                if item["bookId"] in book_map:
                    book_map[item["bookId"]]["stock_quantity"] += int(item["quantity"])
            order["stockRestored"] = True

        order["orderStatus"] = new_status
        if new_pay_status in cfg.PAYMENT_STATUSES:
            order["paymentStatus"] = new_pay_status

        save_store(store)
        run_etl_pipeline()
        sync_order(order)
        sync_inventory(store["books"])

    flash(f"Order {order_id} updated to {new_status}.", "success")
    return redirect(url_for("admin_dashboard") + "#orders-section")


@app.route("/admin/delete-order", methods=["POST"])
@admin_required
def admin_delete_order():
    order_id = request.form.get("order_id", "").strip()
    with DATA_LOCK:
        store = get_store()
        order = next((o for o in store["orders"] if o["orderId"] == order_id), None)
        if order:
            book_map = {b["id"]: b for b in store["books"]}
            if not order.get("stockRestored", False) and order.get("orderStatus") != "Cancelled":
                for item in order.get("items", []):
                    if item["bookId"] in book_map:
                        book_map[item["bookId"]]["stock_quantity"] += int(item["quantity"])
            store["orders"] = [o for o in store["orders"] if o["orderId"] != order_id]
            save_store(store)
            run_etl_pipeline()
            sync_inventory(store["books"])
            flash(f"Order {order_id} deleted and reports recalculated.", "info")
    return redirect(url_for("admin_dashboard") + "#orders-section")


@app.route("/admin/delete-review/<review_id>", methods=["POST"])
@admin_required
def admin_delete_review(review_id):
    with DATA_LOCK:
        store = get_store()
        store["reviews"] = [r for r in store["reviews"] if r["reviewId"] != review_id]
        _recalculate_book_ratings(store)
        save_store(store)
        run_etl_pipeline()
    flash(f"Review {review_id} removed.", "info")
    return redirect(url_for("admin_dashboard") + "#reviews-section")


# ==============================================================================
# 11. ADMIN INVENTORY MANAGEMENT
# ==============================================================================
@app.route("/admin/inventory")
@admin_required
def admin_inventory():
    store = get_store()
    _, items_df = clean_orders_df(store.get("orders", []))
    inv_df = clean_inventory_df(store.get("books", []), items_df)

    q = request.args.get("q", "").strip().lower()
    cat = request.args.get("category", "").strip()
    low_only = request.args.get("low_stock", "") == "1"

    records = inv_df.to_dict(orient="records") if not inv_df.empty else []
    if q:
        records = [r for r in records if q in r["title"].lower() or q in r["author"].lower() or q in r["id"].lower()]
    if cat:
        records = [r for r in records if r["category"] == cat]
    if low_only:
        records = [r for r in records if r["stock_quantity"] <= 5]

    total_valuation = round(sum(r["stock_value_cost"] for r in records), 2)
    retail_valuation = round(sum(r["stock_value_retail"] for r in records), 2)
    return render_template(
        "inventory.html",
        books=records,
        total_valuation=total_valuation,
        retail_valuation=retail_valuation,
        filters={"q": request.args.get("q", ""), "category": cat, "low_stock": low_only},
    )


@app.route("/admin/add-book", methods=["POST"])
@admin_required
def admin_add_book():
    with DATA_LOCK:
        store = get_store()
        next_num = len(store["books"]) + 1
        book_id = f"BK-{next_num:03d}"
        try:
            price = float(request.form.get("price", 0))
            cost_price = float(request.form.get("cost_price", 0))
            stock = max(0, int(request.form.get("stock_quantity", 10)))
            year = int(request.form.get("publication_year", 2025))
            pages = int(request.form.get("pages", 250))
        except ValueError:
            flash("Invalid numeric values for price, cost, or stock.", "error")
            return redirect(url_for("admin_inventory"))

        new_book = {
            "id": book_id,
            "isbn": request.form.get("isbn", "").strip() or f"978-000000{next_num:04d}",
            "title": request.form.get("title", "").strip(),
            "author": request.form.get("author", "").strip(),
            "publisher": request.form.get("publisher", "").strip() or "BookVerse Press",
            "category": request.form.get("category", "Fiction").strip(),
            "description": request.form.get("description", "").strip() or "Curated edition available at BookVerse.",
            "price": price,
            "cost_price": cost_price,
            "stock_quantity": stock,
            "initial_stock": stock,
            "image": request.form.get("image", "").strip() or cfg.FALLBACK_COVER,
            "publication_year": year,
            "language": request.form.get("language", "English").strip(),
            "pages": pages,
            "rating": 4.5,
            "featured": request.form.get("featured") == "on",
            "active": True,
        }
        if not new_book["title"] or not new_book["author"]:
            flash("Book title and author are required.", "error")
            return redirect(url_for("admin_inventory"))

        store["books"].append(new_book)
        save_store(store)
        run_etl_pipeline()
        sync_inventory(store["books"])

    flash(f"Added '{new_book['title']}' ({book_id}) to inventory.", "success")
    return redirect(url_for("admin_inventory"))


@app.route("/admin/edit-book/<book_id>", methods=["POST"])
@admin_required
def admin_edit_book(book_id):
    with DATA_LOCK:
        store = get_store()
        book = next((b for b in store["books"] if b["id"] == book_id), None)
        if not book:
            flash("Book not found.", "error")
            return redirect(url_for("admin_inventory"))
        try:
            book["price"] = max(1.0, float(request.form.get("price", book["price"])))
            book["cost_price"] = max(0.0, float(request.form.get("cost_price", book["cost_price"])))
            book["stock_quantity"] = max(0, int(request.form.get("stock_quantity", book["stock_quantity"])))
        except ValueError:
            flash("Invalid numeric input.", "error")
            return redirect(url_for("admin_inventory"))

        book["active"] = request.form.get("active") == "1"
        book["featured"] = request.form.get("featured") == "1"
        save_store(store)
        run_etl_pipeline()
        sync_inventory(store["books"])

    flash(f"Updated '{book['title']}'.", "success")
    return redirect(url_for("admin_inventory"))


@app.route("/admin/delete-book/<book_id>", methods=["POST"])
@admin_required
def admin_delete_book(book_id):
    with DATA_LOCK:
        store = get_store()
        store["books"] = [b for b in store["books"] if b["id"] != book_id]
        save_store(store)
        run_etl_pipeline()
        sync_inventory(store["books"])
    flash(f"Removed book {book_id} from catalog.", "info")
    return redirect(url_for("admin_inventory"))


@app.route("/admin/download-excel")
@admin_required
def admin_download_excel():
    run_etl_pipeline()
    if not os.path.exists(ORDERS_EXCEL_FILE):
        flash("Excel report could not be generated.", "error")
        return redirect(url_for("admin_dashboard"))
    return send_file(ORDERS_EXCEL_FILE, as_attachment=True, download_name="BookVerse_BI_Report.xlsx")


@app.route("/admin/reset-demo-data", methods=["POST"])
@admin_required
def admin_reset_demo_data():
    with DATA_LOCK:
        save_store(
            {
                "books": copy.deepcopy(cfg.INITIAL_BOOKS),
                "orders": cfg.get_demo_orders(),
                "reviews": cfg.get_demo_reviews(),
            }
        )
        save_customers(cfg.get_demo_customers())
        run_etl_pipeline()
    flash("Demo catalog, orders, customers, and reviews have been reset.", "success")
    return redirect(url_for("admin_dashboard"))


# ==============================================================================
# 12. ERROR HANDLERS
# ==============================================================================
@app.errorhandler(404)
@app.route("/404")
def not_found_error(error=None):
    return render_template("404.html"), 404


@app.errorhandler(500)
@app.route("/500")
def internal_error(error=None):
    return render_template("500.html"), 500


with app.app_context():
    if not os.path.exists(LOCAL_STORE_FILE):
        save_store(get_store())
    if not os.path.exists(CUSTOMERS_FILE):
        save_customers(get_customers())
    run_etl_pipeline()

if __name__ == "__main__":
    app.run(debug=os.getenv("FLASK_DEBUG", "False").lower() == "true", host="0.0.0.0", port=5000)