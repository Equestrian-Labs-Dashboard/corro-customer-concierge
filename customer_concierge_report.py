"""
Corro Customer Concierge Report — Range Based
Shopify Admin API -> Google Sheets -> data/report.json -> GitHub Pages dashboard

This version supports:
- All-years backfill from the first Shopify order.
- Annual/range-based storage.
- Incremental refresh: closed historical ranges are kept; the current/open range refreshes.
- Manual single-range refresh through GitHub Actions inputs.

Never expose Shopify or Google credentials in index.html.
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
from typing import Any, Iterable

import requests
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

load_dotenv()

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)
REPORT_JSON = DATA_DIR / "report.json"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# GitHub Secrets preferred:
# SHOPIFY_STORE_CORRO, SHOPIFY_TOKEN_CORRO, SHEET_ID_CORRO, GOOGLE_CREDENTIALS
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

REPORT_MODE = (os.getenv("REPORT_MODE") or "all_years").strip().lower()
REPORT_START_DATE = (os.getenv("REPORT_START_DATE") or os.getenv("START_DATE") or "").strip()
REPORT_END_DATE = (os.getenv("REPORT_END_DATE") or os.getenv("END_DATE") or "").strip()
HISTORICAL_START_DATE = (os.getenv("HISTORICAL_START_DATE") or "").strip()

INCLUDE_CANCELLED_ORDERS = (os.getenv("INCLUDE_CANCELLED_ORDERS", "false").lower() == "true")
FORCE_REFRESH_ALL_RANGES = (
    os.getenv("FORCE_REFRESH_ALL_RANGES", "false").lower() == "true"
    or REPORT_MODE == "force_all_years"
)
JSON_RAW_LIMIT = int(os.getenv("JSON_RAW_LIMIT", "5000"))


def parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def one_year_before(end: date) -> date:
    try:
        return end.replace(year=end.year - 1)
    except ValueError:
        return end.replace(year=end.year - 1, day=28)


def today_utc() -> date:
    return datetime.now(timezone.utc).date()


def require_env() -> None:
    missing = []
    if not SHOPIFY_STORE:
        missing.append("SHOPIFY_STORE_CORRO")
    if not SHOPIFY_TOKEN:
        missing.append("SHOPIFY_TOKEN_CORRO")
    if not SHEET_ID:
        missing.append("SHEET_ID_CORRO")
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


@dataclass(frozen=True)
class ReportRange:
    range_id: str
    range_label: str
    start: date
    end: date
    status: str  # open or closed


@dataclass
class OrderLine:
    range_id: str
    range_label: str
    range_start: str
    range_end: str
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


class ShopifyClient:
    def __init__(self, store: str, token: str, api_version: str):
        self.endpoint = f"https://{store}/admin/api/{api_version}/graphql.json"
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "application/json",
            "X-Shopify-Access-Token": token,
        })

    def graphql(self, query: str, variables: dict[str, Any] | None = None) -> dict[str, Any]:
        variables = variables or {}
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


EARLIEST_ORDER_QUERY = """
query EarliestOrderForCorroCustomerConcierge {
  orders(first: 1, sortKey: CREATED_AT) {
    nodes { createdAt }
  }
}
"""

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


def find_earliest_shopify_order_date(client: ShopifyClient) -> date:
    data = client.graphql(EARLIEST_ORDER_QUERY)
    nodes = ((data.get("orders") or {}).get("nodes") or [])
    if not nodes:
        return today_utc()
    return parse_date(day_label(nodes[0]["createdAt"]))


def range_id(start: date, end: date) -> str:
    return f"{start.isoformat()}_to_{end.isoformat()}"


def build_calendar_year_ranges(start: date, end: date) -> list[ReportRange]:
    if start > end:
        raise ValueError("Historical start date must be before or equal to end date.")
    ranges: list[ReportRange] = []
    current_start = start
    now = today_utc()
    while current_start <= end:
        current_end = min(date(current_start.year, 12, 31), end)
        status = "open" if current_end >= now or current_start.year == now.year else "closed"
        rid = range_id(current_start, current_end)
        ranges.append(ReportRange(
            range_id=rid,
            range_label=f"{current_start.isoformat()} → {current_end.isoformat()}",
            start=current_start,
            end=current_end,
            status=status,
        ))
        current_start = current_end + timedelta(days=1)
    return ranges


def build_ranges(client: ShopifyClient) -> list[ReportRange]:
    end = parse_date(REPORT_END_DATE) if REPORT_END_DATE else today_utc()

    if REPORT_MODE in {"single", "single_range", "rolling_year"}:
        start = parse_date(REPORT_START_DATE) if REPORT_START_DATE else one_year_before(end)
        rid = range_id(start, end)
        return [ReportRange(rid, f"{start.isoformat()} → {end.isoformat()}", start, end, "open")]

    # Default: all years/ranges from the first Shopify order to selected/today end date.
    if HISTORICAL_START_DATE:
        start = parse_date(HISTORICAL_START_DATE)
    elif REPORT_START_DATE:
        start = parse_date(REPORT_START_DATE)
    else:
        start = find_earliest_shopify_order_date(client)

    return build_calendar_year_ranges(start, end)


def load_existing_payload() -> dict[str, Any]:
    if not REPORT_JSON.exists():
        return {}
    try:
        return json.loads(REPORT_JSON.read_text(encoding="utf-8"))
    except Exception:
        return {}


def existing_range_ids(payload: dict[str, Any]) -> set[str]:
    return {safe_text(r.get("range_id")) for r in payload.get("ranges", []) if r.get("range_id")}


def fetch_order_lines_for_range(client: ShopifyClient, rr: ReportRange) -> list[OrderLine]:
    # Human end date is inclusive. Shopify query end is exclusive.
    end_exclusive = rr.end + timedelta(days=1)
    shopify_query = f"created_at:>={rr.start.isoformat()} created_at:<{end_exclusive.isoformat()}"
    print(f"Fetching Shopify orders: {rr.range_label} ({rr.range_id})")

    after = None
    page = 0
    lines: list[OrderLine] = []

    while True:
        page += 1
        data = client.graphql(ORDERS_QUERY, {"first": 100, "after": after, "query": shopify_query})
        orders_conn = data["orders"]
        orders = orders_conn["nodes"]
        print(f"  Page {page}: {len(orders)} orders")

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
                # Shopify may return zero if the line is fully discounted; keep zero to show discounts.
                discounts = max(gross - net, 0.0)

                variant = item.get("variant") or {}
                inventory_item = variant.get("inventoryItem") or {}
                unit_cost = money((inventory_item.get("unitCost") or {}).get("amount"))
                cogs = unit_cost * qty
                gross_profit = net - cogs
                cogs_status = "missing" if unit_cost <= 0 and net > 0 else "loaded"

                lines.append(OrderLine(
                    range_id=rr.range_id,
                    range_label=rr.range_label,
                    range_start=rr.start.isoformat(),
                    range_end=rr.end.isoformat(),
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

    print(f"  Fetched {len(lines)} line items for {rr.range_label}")
    return lines


def customer_key(line: OrderLine | dict[str, Any]) -> str:
    email = safe_text(line.customer_email if isinstance(line, OrderLine) else line.get("customer_email")).lower()
    phone = safe_text(line.customer_phone if isinstance(line, OrderLine) else line.get("customer_phone"))
    cid = safe_text(line.customer_id if isinstance(line, OrderLine) else line.get("customer_id"))
    name = safe_text(line.customer_name if isinstance(line, OrderLine) else line.get("customer_name")).lower()
    return email or phone or cid or name or "guest"


def aggregate_range(lines: list[OrderLine], rr: ReportRange) -> dict[str, list[dict[str, Any]]]:
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
            "range_id": rr.range_id,
            "range_label": rr.range_label,
            "range_start": rr.start.isoformat(),
            "range_end": rr.end.isoformat(),
            "range_status": rr.status,
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
            "range_id": rr.range_id,
            "range_label": rr.range_label,
            "range_start": rr.start.isoformat(),
            "range_end": rr.end.isoformat(),
            "range_status": rr.status,
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
    months.sort(key=lambda r: (r["month"], r["email"], r["customer_name"]))

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
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Period Start", "value": rr.start.isoformat(), "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Period End", "value": rr.end.isoformat(), "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Total Customers", "value": total_customers, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Orders", "value": total_orders, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Units", "value": total_units, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Gross Sales", "value": gross_sales, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Discounts", "value": discounts, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Net Sales", "value": net_sales, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "COGS", "value": cogs, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Gross Profit", "value": gp, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Gross Margin", "value": round(gp / net_sales, 4) if net_sales else 0, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "AOV Net", "value": round(net_sales / total_orders, 2) if total_orders else 0, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Top Month", "value": top_month, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Top Month Net Sales", "value": round(top_month_sales, 2), "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Missing COGS Sales", "value": missing_cogs_sales, "updated_at": updated_at},
        {"range_id": rr.range_id, "range_label": rr.range_label, "range_start": rr.start.isoformat(), "range_end": rr.end.isoformat(), "range_status": rr.status, "metric": "Shopify Store", "value": SHOPIFY_STORE, "updated_at": updated_at},
    ]

    range_row = {
        "range_id": rr.range_id,
        "range_label": rr.range_label,
        "range_start": rr.start.isoformat(),
        "range_end": rr.end.isoformat(),
        "range_status": rr.status,
        "customers": total_customers,
        "orders": total_orders,
        "units": total_units,
        "gross_sales": gross_sales,
        "discounts": discounts,
        "net_sales": net_sales,
        "cogs": cogs,
        "gross_profit": gp,
        "gross_margin": round(gp / net_sales, 4) if net_sales else 0,
        "missing_cogs_sales": missing_cogs_sales,
        "updated_at": updated_at,
    }

    return {
        "report_ranges": [range_row],
        "summary_by_range": summary,
        "customers_by_range": customers,
        "customer_months_by_range": months,
        "raw_order_lines_by_range": raw,
    }


def normalize_existing_rows(payload: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    # Supports both the new names and older names from the first version.
    return {
        "report_ranges": payload.get("ranges", []) or payload.get("report_ranges", []),
        "summary_by_range": payload.get("summary", []) or payload.get("summary_by_range", []),
        "customers_by_range": payload.get("customers", []) or payload.get("customers_by_range", []),
        "customer_months_by_range": payload.get("months", []) or payload.get("customer_months_by_range", []),
        "raw_order_lines_by_range": payload.get("raw", []) or payload.get("raw_order_lines_by_range", []),
    }


def remove_ranges(rows: Iterable[dict[str, Any]], range_ids: set[str]) -> list[dict[str, Any]]:
    return [r for r in rows if safe_text(r.get("range_id")) not in range_ids]


def merge_data(existing_payload: dict[str, Any], fetched: list[dict[str, list[dict[str, Any]]]], refreshed_ids: set[str]) -> dict[str, list[dict[str, Any]]]:
    existing = normalize_existing_rows(existing_payload)
    merged = {
        "report_ranges": remove_ranges(existing["report_ranges"], refreshed_ids),
        "summary_by_range": remove_ranges(existing["summary_by_range"], refreshed_ids),
        "customers_by_range": remove_ranges(existing["customers_by_range"], refreshed_ids),
        "customer_months_by_range": remove_ranges(existing["customer_months_by_range"], refreshed_ids),
        "raw_order_lines_by_range": remove_ranges(existing["raw_order_lines_by_range"], refreshed_ids),
    }
    for block in fetched:
        for key, rows in block.items():
            merged[key].extend(rows)

    merged["report_ranges"].sort(key=lambda r: safe_text(r.get("range_start")))
    merged["summary_by_range"].sort(key=lambda r: (safe_text(r.get("range_start")), safe_text(r.get("metric"))))
    merged["customers_by_range"].sort(key=lambda r: (safe_text(r.get("range_start")), -(money(r.get("net_sales")))))
    merged["customer_months_by_range"].sort(key=lambda r: (safe_text(r.get("range_start")), safe_text(r.get("month")), safe_text(r.get("email"))))
    merged["raw_order_lines_by_range"].sort(key=lambda r: (safe_text(r.get("range_start")), safe_text(r.get("created_at")), safe_text(r.get("order_name"))))
    return merged


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
        raise RuntimeError("Missing Google credentials. Set GOOGLE_CREDENTIALS in GitHub Secrets or use service-account.json locally.")
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
                        "gridProperties": {"rowCount": 1000, "columnCount": 40},
                    }
                }
            })
    if requests_batch:
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests_batch}).execute()


def get_sheet_properties(service, spreadsheet_id: str, tab: str) -> dict[str, Any]:
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == tab:
            return props
    raise RuntimeError(f"Google Sheet tab not found: {tab}")


def ensure_tab_size(service, spreadsheet_id: str, tab: str, needed_rows: int, needed_cols: int) -> None:
    """Resize a Google Sheets tab before writing large datasets.

    Google Sheets tabs often start with only 1,000 or 5,000 rows. If we try
    writing to A5001 while the tab has only 5,000 rows, the API throws:
    "Range exceeds grid limits". This expands the grid first.
    """
    props = get_sheet_properties(service, spreadsheet_id, tab)
    grid = props.get("gridProperties", {})
    current_rows = int(grid.get("rowCount", 0) or 0)
    current_cols = int(grid.get("columnCount", 0) or 0)

    # Add a buffer so the next refresh has room and does not resize on every run.
    target_rows = max(current_rows, needed_rows + 500, 1000)
    target_cols = max(current_cols, needed_cols + 5, 40)

    if target_rows == current_rows and target_cols == current_cols:
        return

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": props["sheetId"],
                            "gridProperties": {
                                "rowCount": target_rows,
                                "columnCount": target_cols,
                            },
                        },
                        "fields": "gridProperties(rowCount,columnCount)",
                    }
                }
            ]
        },
    ).execute()
    print(f"Resized {tab} to {target_rows} rows x {target_cols} columns")


def rows_to_values(rows: list[dict[str, Any]], default_headers: list[str]) -> list[list[Any]]:
    headers = list(rows[0].keys()) if rows else default_headers
    values = [headers]
    for row in rows:
        values.append([json_safe(row.get(h, "")) for h in headers])
    return values


def clear_and_write_tab(service, spreadsheet_id: str, tab: str, rows: list[dict[str, Any]], headers: list[str]) -> None:
    values = rows_to_values(rows, headers)
    needed_rows = max(len(values), 1)
    needed_cols = max((len(row) for row in values), default=len(headers) or 1)

    # Critical fix: expand the worksheet grid before writing chunks.
    # Without this, large tabs like customers_by_range fail at A5001.
    ensure_tab_size(service, spreadsheet_id, tab, needed_rows, needed_cols)

    service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=f"'{tab}'").execute()
    chunk_size = 4000
    for start_idx in range(0, len(values), chunk_size):
        chunk = values[start_idx:start_idx + chunk_size]
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{tab}'!A{start_idx + 1}",
            valueInputOption="USER_ENTERED",
            body={"values": chunk},
        ).execute()
    print(f"Wrote {len(values) - 1} rows to {tab}")


def write_google_sheets(data: dict[str, list[dict[str, Any]]]) -> None:
    default_headers = {
        "report_ranges": ["range_id", "range_label", "range_start", "range_end", "range_status", "customers", "orders", "units", "gross_sales", "discounts", "net_sales", "cogs", "gross_profit", "gross_margin", "missing_cogs_sales", "updated_at"],
        "summary_by_range": ["range_id", "range_label", "range_start", "range_end", "range_status", "metric", "value", "updated_at"],
        "customers_by_range": ["range_id", "range_label", "range_start", "range_end", "range_status", "rank", "customer_name", "email", "phone", "orders", "units", "first_order_date", "last_order_date", "top_month", "top_month_net_sales", "gross_sales", "discounts", "net_sales", "cogs", "gross_profit", "gross_margin", "aov_net", "missing_cogs_sales", "updated_at"],
        "customer_months_by_range": ["range_id", "range_label", "range_start", "range_end", "range_status", "customer_name", "email", "phone", "month", "orders", "units", "gross_sales", "discounts", "net_sales", "cogs", "gross_profit", "gross_margin", "missing_cogs_sales", "updated_at"],
        "raw_order_lines_by_range": ["range_id", "range_label", "range_start", "range_end", "order_id", "order_name", "created_at", "order_date", "month", "financial_status", "customer_id", "customer_name", "customer_email", "customer_phone", "sku", "vendor", "product_title", "units", "gross_sales", "discounts", "net_sales", "unit_cost", "cogs", "gross_profit", "cogs_status"],
    }
    service = sheets_service()
    tabs = list(default_headers.keys())
    ensure_tabs(service, SHEET_ID, tabs)
    for tab in tabs:
        clear_and_write_tab(service, SHEET_ID, tab, data.get(tab, []), default_headers[tab])


def write_json(data: dict[str, list[dict[str, Any]]]) -> None:
    ranges = data.get("report_ranges", [])
    default_range = ranges[-1]["range_id"] if ranges else ""
    updated_at = datetime.now(timezone.utc).isoformat(timespec="seconds")

    raw_full = data.get("raw_order_lines_by_range", [])
    raw_limited = raw_full[:JSON_RAW_LIMIT] if JSON_RAW_LIMIT > 0 else []

    payload = {
        "meta": {
            "report_name": "Corro Customer Concierge Report",
            "updated_at": updated_at,
            "shopify_store": SHOPIFY_STORE,
            "source": "Shopify Admin API + Google Sheets + GitHub Actions",
            "mode": REPORT_MODE,
            "default_range_id": default_range,
            "ranges_count": len(ranges),
            "raw_json_limit": JSON_RAW_LIMIT,
            "raw_total_rows_in_sheet": len(raw_full),
        },
        "ranges": ranges,
        "summary": data.get("summary_by_range", []),
        "customers": data.get("customers_by_range", []),
        "months": data.get("customer_months_by_range", []),
        "raw": raw_limited,
    }
    REPORT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {REPORT_JSON.relative_to(ROOT)} with {len(ranges)} ranges and {len(raw_limited)}/{len(raw_full)} raw rows in JSON")


def main() -> int:
    require_env()
    client = ShopifyClient(SHOPIFY_STORE, SHOPIFY_TOKEN, SHOPIFY_API_VERSION)
    wanted_ranges = build_ranges(client)
    existing_payload = load_existing_payload()
    already = existing_range_ids(existing_payload)

    to_fetch: list[ReportRange] = []
    for rr in wanted_ranges:
        if FORCE_REFRESH_ALL_RANGES:
            to_fetch.append(rr)
        elif REPORT_MODE in {"single", "single_range", "rolling_year"}:
            to_fetch.append(rr)
        elif rr.status == "open":
            # Always update current/open range so refresh adds new Shopify orders.
            to_fetch.append(rr)
        elif rr.range_id not in already:
            # Historical closed range missing from JSON, so fetch it once.
            to_fetch.append(rr)
        else:
            print(f"Skipping closed existing range: {rr.range_label}")

    print(f"Ranges wanted: {len(wanted_ranges)} | Ranges to fetch now: {len(to_fetch)}")
    fetched_blocks = []
    refreshed_ids: set[str] = set()
    for rr in to_fetch:
        lines = fetch_order_lines_for_range(client, rr)
        fetched_blocks.append(aggregate_range(lines, rr))
        refreshed_ids.add(rr.range_id)

    # If no range was fetched, keep existing data and just write it again.
    merged = merge_data(existing_payload, fetched_blocks, refreshed_ids)

    # Make sure every wanted range appears in the final JSON after first complete run.
    write_google_sheets(merged)
    write_json(merged)
    print("Done.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
