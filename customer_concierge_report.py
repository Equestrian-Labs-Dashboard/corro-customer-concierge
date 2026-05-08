"""
Corro Customer Concierge Report
Shopify Admin API -> Google Sheets -> GitHub Pages JSON dashboard

Public dashboard must NEVER receive Shopify or Google credentials.
This script runs server-side/local/GitHub Actions, writes to Google Sheets,
and exports data/report.json for index.html.
"""

from __future__ import annotations

import base64
import json
import os
import sys
import time
from collections import defaultdict
from dataclasses import asdict, dataclass
from datetime import date, datetime, timezone, timedelta
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

load_dotenv()

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Secrets/variables supported:
# GitHub Actions preferred: SHOPIFY_STORE_CORRO, SHOPIFY_TOKEN_CORRO, SHEET_ID_CORRO, GOOGLE_CREDENTIALS
# Local fallback: SHOPIFY_STORE, SHOPIFY_TOKEN, GOOGLE_SHEET_ID, GOOGLE_SERVICE_ACCOUNT_FILE
SHOPIFY_STORE = (
    os.getenv("SHOPIFY_STORE_CORRO")
    or os.getenv("SHOPIFY_STORE")
    or os.getenv("SHOPIFY_SHOP")
    or ""
).strip().replace("https://", "").replace("http://", "").rstrip("/")

SHOPIFY_TOKEN = (
    os.getenv("SHOPIFY_TOKEN_CORRO")
    or os.getenv("SHOPIFY_TOKEN")
    or os.getenv("SHOPIFY_ADMIN_TOKEN")
    or ""
).strip()

SHOPIFY_API_VERSION = (
    os.getenv("SHOPIFY_API_VERSION")
    or os.getenv("SHOPIFY_API_VERSION_CORRO")
    or "2026-04"
).strip()

SHEET_ID = (
    os.getenv("SHEET_ID_CORRO")
    or os.getenv("GOOGLE_SHEET_ID")
    or os.getenv("SHEET_ID")
    or ""
).strip()

GOOGLE_CREDENTIALS = os.getenv("GOOGLE_CREDENTIALS", "").strip()
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service-account.json").strip()

REPORT_START_DATE = (os.getenv("REPORT_START_DATE") or os.getenv("START_DATE") or "").strip()
REPORT_END_DATE = (os.getenv("REPORT_END_DATE") or os.getenv("END_DATE") or "").strip()

INCLUDE_CANCELLED_ORDERS = (os.getenv("INCLUDE_CANCELLED_ORDERS", "false").lower() == "true")


def parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def one_year_before(end: date) -> date:
    try:
        return end.replace(year=end.year - 1)
    except ValueError:
        # Feb 29 -> Feb 28 on non-leap year
        return end.replace(year=end.year - 1, day=28)


def get_period() -> tuple[date, date]:
    """
    Default range: rolling full year ending today UTC.
    Example: REPORT_END_DATE=2026-05-07 -> REPORT_START_DATE=2025-05-07
    End date is inclusive for humans. The Shopify query internally uses end + 1 day.
    """
    end = parse_date(REPORT_END_DATE) if REPORT_END_DATE else datetime.now(timezone.utc).date()
    start = parse_date(REPORT_START_DATE) if REPORT_START_DATE else one_year_before(end)
    if start > end:
        raise ValueError("REPORT_START_DATE must be before or equal to REPORT_END_DATE.")
    return start, end


def require_env() -> None:
    missing = []
    if not SHOPIFY_STORE:
        missing.append("SHOPIFY_STORE_CORRO or SHOPIFY_STORE")
    if not SHOPIFY_TOKEN:
        missing.append("SHOPIFY_TOKEN_CORRO or SHOPIFY_TOKEN")
    if not SHEET_ID:
        missing.append("SHEET_ID_CORRO or GOOGLE_SHEET_ID")
    if missing:
        raise RuntimeError("Missing required environment variables: " + ", ".join(missing))


def money(value: Any) -> float:
    try:
        if value in (None, ""):
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def safe_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def day_label(iso_datetime: str) -> str:
    return datetime.fromisoformat(iso_datetime.replace("Z", "+00:00")).strftime("%Y-%m-%d")


def month_label(iso_datetime: str) -> str:
    return datetime.fromisoformat(iso_datetime.replace("Z", "+00:00")).strftime("%Y-%m")


def amount_from_money_bag(node: dict[str, Any] | None, key: str) -> float:
    if not node:
        return 0.0
    bag = node.get(key) or {}
    shop_money = bag.get("shopMoney") or {}
    return money(shop_money.get("amount"))


def json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {k: json_safe(v) for k, v in value.items()}
    if isinstance(value, list):
        return [json_safe(v) for v in value]
    if isinstance(value, set):
        return sorted(list(value))
    return value


class ShopifyClient:
    def __init__(self, store: str, token: str, api_version: str):
        self.endpoint = f"https://{store}/admin/api/{api_version}/graphql.json"
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "application/json",
            "X-Shopify-Access-Token": token,
        })

    def graphql(self, query: str, variables: dict[str, Any]) -> dict[str, Any]:
        for attempt in range(1, 7):
            response = self.session.post(
                self.endpoint,
                json={"query": query, "variables": variables},
                timeout=120,
            )
            if response.status_code in (429, 500, 502, 503, 504):
                wait = min(2 ** attempt, 45)
                print(f"Shopify temporary status {response.status_code}; retrying in {wait}s...")
                time.sleep(wait)
                continue
            if response.status_code >= 400:
                print(response.text[:2000])
            response.raise_for_status()
            payload = response.json()
            if payload.get("errors"):
                raise RuntimeError(json.dumps(payload["errors"], indent=2))
            return payload["data"]
        raise RuntimeError("Shopify request failed after retries.")


ORDERS_QUERY = """
query OrdersForCorroCustomerConcierge($first: Int!, $after: String, $query: String!) {
  orders(first: $first, after: $after, query: $query, sortKey: CREATED_AT) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id
      name
      createdAt
      cancelledAt
      displayFinancialStatus
      currencyCode
      email
      phone
      customer {
        id
        displayName
        email
        phone
      }
      lineItems(first: 250) {
        nodes {
          id
          title
          quantity
          sku
          vendor
          originalTotalSet { shopMoney { amount currencyCode } }
          discountedTotalSet { shopMoney { amount currencyCode } }
          variant {
            id
            sku
            inventoryItem {
              unitCost { amount currencyCode }
            }
          }
        }
      }
    }
  }
}
"""


@dataclass
class OrderLine:
    order_id: str
    order_name: str
    created_at: str
    order_date: str
    month: str
    financial_status: str
    customer_id: str
    customer_name: str
    customer_email: str
    customer_phone: str
    sku: str
    vendor: str
    product_title: str
    units: int
    gross_sales: float
    discounts: float
    net_sales: float
    unit_cost: float
    cogs: float
    gross_profit: float
    cogs_status: str


def fetch_order_lines(start: date, end: date) -> list[OrderLine]:
    client = ShopifyClient(SHOPIFY_STORE, SHOPIFY_TOKEN, SHOPIFY_API_VERSION)
    # User-facing end date is inclusive. Shopify query end is exclusive.
    end_exclusive = end + timedelta(days=1)
    shopify_query = f"created_at:>={start.isoformat()} created_at:<{end_exclusive.isoformat()}"
    print(f"Fetching Shopify orders from {SHOPIFY_STORE}: {start.isoformat()} through {end.isoformat()} inclusive")

    after = None
    page = 0
    lines: list[OrderLine] = []

    while True:
        page += 1
        data = client.graphql(ORDERS_QUERY, {"first": 100, "after": after, "query": shopify_query})
        orders_conn = data["orders"]
        orders = orders_conn["nodes"]
        print(f"Page {page}: {len(orders)} orders")

        for order in orders:
            if order.get("cancelledAt") and not INCLUDE_CANCELLED_ORDERS:
                continue

            customer = order.get("customer") or {}
            customer_id = safe_text(customer.get("id")) or "guest"
            customer_name = safe_text(customer.get("displayName")) or "Guest Customer"
            customer_email = safe_text(customer.get("email")) or safe_text(order.get("email"))
            customer_phone = safe_text(customer.get("phone")) or safe_text(order.get("phone"))
            created_at = safe_text(order.get("createdAt"))

            for item in (order.get("lineItems") or {}).get("nodes", []):
                qty = int(item.get("quantity") or 0)
                if qty <= 0:
                    continue

                gross = amount_from_money_bag(item, "originalTotalSet")
                net = amount_from_money_bag(item, "discountedTotalSet")
                if net == 0 and gross > 0:
                    net = gross
                discounts = max(gross - net, 0.0)

                variant = item.get("variant") or {}
                inventory_item = variant.get("inventoryItem") or {}
                unit_cost = money((inventory_item.get("unitCost") or {}).get("amount"))
                cogs = unit_cost * qty
                gross_profit = net - cogs
                cogs_status = "missing" if unit_cost <= 0 and net > 0 else "loaded"

                lines.append(OrderLine(
                    order_id=safe_text(order.get("id")),
                    order_name=safe_text(order.get("name")),
                    created_at=created_at,
                    order_date=day_label(created_at),
                    month=month_label(created_at),
                    financial_status=safe_text(order.get("displayFinancialStatus")),
                    customer_id=customer_id,
                    customer_name=customer_name,
                    customer_email=customer_email,
                    customer_phone=customer_phone,
                    sku=safe_text(item.get("sku")) or safe_text(variant.get("sku")),
                    vendor=safe_text(item.get("vendor")),
                    product_title=safe_text(item.get("title")),
                    units=qty,
                    gross_sales=round(gross, 2),
                    discounts=round(discounts, 2),
                    net_sales=round(net, 2),
                    unit_cost=round(unit_cost, 2),
                    cogs=round(cogs, 2),
                    gross_profit=round(gross_profit, 2),
                    cogs_status=cogs_status,
                ))

        info = orders_conn["pageInfo"]
        if not info["hasNextPage"]:
            break
        after = info["endCursor"]
        time.sleep(0.25)

    print(f"Fetched {len(lines)} order lines")
    return lines


def customer_key(line: OrderLine) -> str:
    return (
        line.customer_email.lower()
        or line.customer_phone
        or line.customer_id
        or line.customer_name.lower()
        or "guest"
    )


def aggregate(lines: list[OrderLine], start: date, end: date) -> dict[str, list[dict[str, Any]]]:
    updated_at = datetime.now(timezone.utc).isoformat(timespec="seconds")

    by_customer: dict[str, dict[str, Any]] = {}
    by_customer_month: dict[tuple[str, str], dict[str, Any]] = {}
    by_month_total = defaultdict(float)

    for line in lines:
        key = customer_key(line)
        if key not in by_customer:
            by_customer[key] = {
                "customer_name": line.customer_name,
                "email": line.customer_email,
                "phone": line.customer_phone,
                "orders_set": set(),
                "first_order_date": line.order_date,
                "last_order_date": line.order_date,
                "units": 0,
                "gross_sales": 0.0,
                "discounts": 0.0,
                "net_sales": 0.0,
                "cogs": 0.0,
                "gross_profit": 0.0,
                "missing_cogs_sales": 0.0,
                "months": defaultdict(float),
            }

        c = by_customer[key]
        c["orders_set"].add(line.order_name)
        c["first_order_date"] = min(c["first_order_date"], line.order_date)
        c["last_order_date"] = max(c["last_order_date"], line.order_date)
        c["units"] += line.units
        c["gross_sales"] += line.gross_sales
        c["discounts"] += line.discounts
        c["net_sales"] += line.net_sales
        c["cogs"] += line.cogs
        c["gross_profit"] += line.gross_profit
        c["months"][line.month] += line.net_sales
        if line.cogs_status == "missing":
            c["missing_cogs_sales"] += line.net_sales

        mkey = (key, line.month)
        if mkey not in by_customer_month:
            by_customer_month[mkey] = {
                "customer_name": line.customer_name,
                "email": line.customer_email,
                "phone": line.customer_phone,
                "month": line.month,
                "orders_set": set(),
                "units": 0,
                "gross_sales": 0.0,
                "discounts": 0.0,
                "net_sales": 0.0,
                "cogs": 0.0,
                "gross_profit": 0.0,
                "missing_cogs_sales": 0.0,
            }
        cm = by_customer_month[mkey]
        cm["orders_set"].add(line.order_name)
        cm["units"] += line.units
        cm["gross_sales"] += line.gross_sales
        cm["discounts"] += line.discounts
        cm["net_sales"] += line.net_sales
        cm["cogs"] += line.cogs
        cm["gross_profit"] += line.gross_profit
        if line.cogs_status == "missing":
            cm["missing_cogs_sales"] += line.net_sales

        by_month_total[line.month] += line.net_sales

    customers = []
    for c in by_customer.values():
        orders = len(c["orders_set"])
        top_month, top_month_net_sales = ("", 0.0)
        if c["months"]:
            top_month, top_month_net_sales = max(c["months"].items(), key=lambda x: x[1])
        net = c["net_sales"]
        gp = c["gross_profit"]
        customers.append({
            "rank": 0,
            "customer_name": c["customer_name"],
            "email": c["email"],
            "phone": c["phone"],
            "orders": orders,
            "units": c["units"],
            "first_order_date": c["first_order_date"],
            "last_order_date": c["last_order_date"],
            "top_month": top_month,
            "top_month_net_sales": round(top_month_net_sales, 2),
            "gross_sales": round(c["gross_sales"], 2),
            "discounts": round(c["discounts"], 2),
            "net_sales": round(net, 2),
            "cogs": round(c["cogs"], 2),
            "gross_profit": round(gp, 2),
            "gross_margin": round(gp / net, 4) if net else 0,
            "aov_net": round(net / orders, 2) if orders else 0,
            "missing_cogs_sales": round(c["missing_cogs_sales"], 2),
            "period_start": start.isoformat(),
            "period_end": end.isoformat(),
            "updated_at": updated_at,
        })

    customers.sort(key=lambda r: r["net_sales"], reverse=True)
    for i, row in enumerate(customers, 1):
        row["rank"] = i

    months = []
    for cm in by_customer_month.values():
        net = cm["net_sales"]
        gp = cm["gross_profit"]
        months.append({
            "customer_name": cm["customer_name"],
            "email": cm["email"],
            "phone": cm["phone"],
            "month": cm["month"],
            "orders": len(cm["orders_set"]),
            "units": cm["units"],
            "gross_sales": round(cm["gross_sales"], 2),
            "discounts": round(cm["discounts"], 2),
            "net_sales": round(net, 2),
            "cogs": round(cm["cogs"], 2),
            "gross_profit": round(gp, 2),
            "gross_margin": round(gp / net, 4) if net else 0,
            "missing_cogs_sales": round(cm["missing_cogs_sales"], 2),
            "updated_at": updated_at,
        })
    months.sort(key=lambda r: (r["email"], r["month"]))

    raw = [asdict(line) for line in lines]

    total_customers = len(customers)
    total_orders = len({line.order_name for line in lines})
    total_units = sum(line.units for line in lines)
    gross_sales = round(sum(line.gross_sales for line in lines), 2)
    discounts = round(sum(line.discounts for line in lines), 2)
    net_sales = round(sum(line.net_sales for line in lines), 2)
    cogs = round(sum(line.cogs for line in lines), 2)
    gp = round(net_sales - cogs, 2)
    missing_cogs_sales = round(sum(line.net_sales for line in lines if line.cogs_status == "missing"), 2)

    top_month, top_month_sales = ("", 0.0)
    if by_month_total:
        top_month, top_month_sales = max(by_month_total.items(), key=lambda x: x[1])

    summary = [
        {"metric": "Period Start", "value": start.isoformat(), "updated_at": updated_at},
        {"metric": "Period End", "value": end.isoformat(), "updated_at": updated_at},
        {"metric": "Total Customers", "value": total_customers, "updated_at": updated_at},
        {"metric": "Orders", "value": total_orders, "updated_at": updated_at},
        {"metric": "Units", "value": total_units, "updated_at": updated_at},
        {"metric": "Gross Sales", "value": gross_sales, "updated_at": updated_at},
        {"metric": "Discounts", "value": discounts, "updated_at": updated_at},
        {"metric": "Net Sales", "value": net_sales, "updated_at": updated_at},
        {"metric": "COGS", "value": cogs, "updated_at": updated_at},
        {"metric": "Gross Profit", "value": gp, "updated_at": updated_at},
        {"metric": "Gross Margin", "value": round(gp / net_sales, 4) if net_sales else 0, "updated_at": updated_at},
        {"metric": "AOV Net", "value": round(net_sales / total_orders, 2) if total_orders else 0, "updated_at": updated_at},
        {"metric": "Top Month", "value": top_month, "updated_at": updated_at},
        {"metric": "Top Month Net Sales", "value": round(top_month_sales, 2), "updated_at": updated_at},
        {"metric": "Missing COGS Sales", "value": missing_cogs_sales, "updated_at": updated_at},
        {"metric": "Shopify Store", "value": SHOPIFY_STORE, "updated_at": updated_at},
    ]

    return {
        "summary_rolling_year": summary,
        "customers_rolling_year": customers,
        "customer_months": months,
        "raw_order_lines": raw,
    }


def get_google_credentials():
    if GOOGLE_CREDENTIALS:
        raw = GOOGLE_CREDENTIALS
        try:
            info = json.loads(raw)
        except json.JSONDecodeError:
            try:
                info = json.loads(base64.b64decode(raw).decode("utf-8"))
            except Exception as exc:
                raise RuntimeError("GOOGLE_CREDENTIALS must be raw JSON or base64 JSON.") from exc
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    path = ROOT / GOOGLE_SERVICE_ACCOUNT_FILE
    if not path.exists():
        raise RuntimeError(
            "Missing Google credentials. Set GOOGLE_CREDENTIALS in GitHub Secrets or put service-account.json locally."
        )
    return service_account.Credentials.from_service_account_file(str(path), scopes=SCOPES)


def sheets_service():
    creds = get_google_credentials()
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def ensure_tabs(service, spreadsheet_id: str, tab_names: list[str]) -> None:
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}
    requests_batch = []
    for tab in tab_names:
        if tab not in existing:
            requests_batch.append({
                "addSheet": {
                    "properties": {
                        "title": tab,
                        "gridProperties": {"rowCount": 1000, "columnCount": 30},
                    }
                }
            })
    if requests_batch:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests_batch},
        ).execute()


def rows_to_values(rows: list[dict[str, Any]], default_headers: list[str]) -> list[list[Any]]:
    headers = list(rows[0].keys()) if rows else default_headers
    values = [headers]
    for row in rows:
        values.append([json_safe(row.get(h, "")) for h in headers])
    return values


def clear_and_write_tab(service, spreadsheet_id: str, tab: str, rows: list[dict[str, Any]], headers: list[str]) -> None:
    values = rows_to_values(rows, headers)
    service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=f"'{tab}'").execute()
    chunk_size = 5000
    for start_idx in range(0, len(values), chunk_size):
        chunk = values[start_idx:start_idx + chunk_size]
        start_row = start_idx + 1
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab}'!A{start_row}",
            valueInputOption="USER_ENTERED",
            body={"values": chunk},
        ).execute()
    print(f"Wrote {len(values) - 1} rows to {tab}")


def write_google_sheets(data: dict[str, list[dict[str, Any]]]) -> None:
    default_headers = {
        "summary_rolling_year": ["metric", "value", "updated_at"],
        "customers_rolling_year": [
            "rank", "customer_name", "email", "phone", "orders", "units", "first_order_date",
            "last_order_date", "top_month", "top_month_net_sales", "gross_sales", "discounts",
            "net_sales", "cogs", "gross_profit", "gross_margin", "aov_net", "missing_cogs_sales",
            "period_start", "period_end", "updated_at",
        ],
        "customer_months": [
            "customer_name", "email", "phone", "month", "orders", "units", "gross_sales", "discounts",
            "net_sales", "cogs", "gross_profit", "gross_margin", "missing_cogs_sales", "updated_at",
        ],
        "raw_order_lines": [
            "order_id", "order_name", "created_at", "order_date", "month", "financial_status", "customer_id",
            "customer_name", "customer_email", "customer_phone", "sku", "vendor", "product_title", "units",
            "gross_sales", "discounts", "net_sales", "unit_cost", "cogs", "gross_profit", "cogs_status",
        ],
    }

    service = sheets_service()
    tabs = list(default_headers.keys())
    ensure_tabs(service, SHEET_ID, tabs)
    for tab in tabs:
        clear_and_write_tab(service, SHEET_ID, tab, data.get(tab, []), default_headers[tab])


def write_json(data: dict[str, list[dict[str, Any]]], start: date, end: date) -> None:
    summary = data.get("summary_rolling_year", [])
    updated_at = summary[0].get("updated_at") if summary else datetime.now(timezone.utc).isoformat(timespec="seconds")
    payload = {
        "meta": {
            "report_name": "Corro Customer Concierge Report",
            "period_start": start.isoformat(),
            "period_end": end.isoformat(),
            "updated_at": updated_at,
            "shopify_store": SHOPIFY_STORE,
            "source": "Shopify Admin API + Google Sheets",
        },
        "summary": data.get("summary_rolling_year", []),
        "customers": data.get("customers_rolling_year", []),
        "months": data.get("customer_months", []),
        "raw": data.get("raw_order_lines", []),
    }
    output = DATA_DIR / "report.json"
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {output.relative_to(ROOT)}")


def main() -> int:
    require_env()
    start, end = get_period()
    lines = fetch_order_lines(start, end)
    data = aggregate(lines, start, end)
    write_google_sheets(data)
    write_json(data, start, end)
    print("Done.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
