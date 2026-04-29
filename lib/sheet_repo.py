from __future__ import annotations
from datetime import datetime
from zoneinfo import ZoneInfo

TAIPEI_TZ = ZoneInfo("Asia/Taipei")

def now_str() -> str:
    return datetime.now(TAIPEI_TZ).strftime("%Y/%m/%d %H:%M")

def today_mmdd() -> str:
    return datetime.now(TAIPEI_TZ).strftime("%m%d")
from typing import Dict, List, Optional, Set, Tuple
import time
import os
import json

import gspread
from google.oauth2.service_account import Credentials

from collections import defaultdict
from collections import defaultdict

PROJECT_ROOT = os.path.dirname(os.path.dirname(__file__))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_google_credentials():
    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
    return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)

SERVICE_ACCOUNT_FILE = "credentials/google-service-account.json"
SPREADSHEET_ID = "1jtPGTV1dCT6QhI9gehKrMu_YwSvNaCrSViKU9S9Rxp8"
WORKSHEET_NAME = "2026Summer_Taipei"

HEADERS = [
    "訂單日期與時間",  # A
    "訂單ID",        # B
    "訂單狀態",      # C
    "名字",          # D
    "樓層",          # E
    "排數",          # F
    "座位",          # G
    "票價",          # H
    "訂單備註",      # I
    "是否開放取票",  # J
    "是否已取票",    # K
    "付款狀態",      # L
    "是否已調票", #M
]

# ===== caches =====
_ws_cache = None
_sold_cache = None
_sold_cache_time = 0
_SOLD_CACHE_TTL = 2  # 秒


def get_worksheet():
    global _ws_cache

    if _ws_cache is not None:
        return _ws_cache

    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    _ws_cache = spreadsheet.worksheet(WORKSHEET_NAME)
    return _ws_cache

def get_config_worksheet(name: str):
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    return spreadsheet.worksheet(name)
    
def clear_caches():
    global _sold_cache, _sold_cache_time
    _sold_cache = None
    _sold_cache_time = 0


def ensure_headers() -> None:
    ws = get_worksheet()
    current = ws.row_values(1)
    if current != HEADERS:
        ws.update("A1:M1", [HEADERS])


def now_str() -> str:
    return datetime.now(TAIPEI_TZ).strftime("%Y/%m/%d %H:%M")


def today_mmdd() -> str:
    return datetime.now(TAIPEI_TZ).strftime("%m%d")


def normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_int(value) -> Optional[int]:
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def normalize_bool(value) -> bool:
    text = normalize_text(value).lower()
    return text in {"true", "1", "yes", "y", "是"}


def sanitize_name_for_order_id(name: str) -> str:
    """
    讓 order_id 比較穩定，去掉空白與底線
    """
    text = normalize_text(name)
    text = text.replace(" ", "")
    text = text.replace("_", "")
    return text or "未命名"


def get_all_records() -> List[dict]:
    ws = get_worksheet()
    return ws.get_all_records()

import random

def generate_order_id(name: str) -> str:
    now = datetime.now(TAIPEI_TZ)

    return (
        "TP"
        + now.strftime("%m%d-%H%M-")
        + f"{now.second:02d}"
        + str(random.randint(0,9))
    )


def get_active_records() -> List[dict]:
    rows = get_all_records()
    return [
        row for row in rows
        if normalize_text(row.get("訂單狀態")).lower() in {"active", "locked"}
    ]


def append_order_rows(name: str, seat_rows: List[Dict]) -> str:
    ws = get_worksheet()
    order_id = generate_order_id(name)
    dt = now_str()

    values = []
    for seat in seat_rows:
        values.append([
            dt,                         # A 訂單日期與時間
            order_id,                   # B 訂單ID
            "active",                   # C 訂單狀態
            name,                       # D 名字
            seat["floor"],              # E 樓層
            str(seat["row_label"]),     # F 排數
            int(seat["seat_number"]),   # G 座位
            int(seat["price"]),         # H 票價
            "",                         # I 訂單備註
            False,                      # J 是否開放取票
            False,                      # K 是否已取票
            False,                      # L 付款狀態
            False,                      # M 調票
        ])

    if values:
        ws.append_rows(values, value_input_option="USER_ENTERED")
        clear_caches()

    return order_id


def build_active_sold_seat_keys() -> Set[Tuple[str, str, int]]:
    """
    回傳 active 訂單佔用的座位：
    (樓層, 排數, 座位)
    例如 ('1樓', '3', 12)
    """
    global _sold_cache, _sold_cache_time

    now = time.time()
    if _sold_cache is not None and (now - _sold_cache_time) < _SOLD_CACHE_TTL:
        return _sold_cache

    sold = set()
    for row in get_active_records():
        floor = normalize_text(row.get("樓層"))
        row_label = normalize_text(row.get("排數"))
        seat_number = normalize_int(row.get("座位"))

        if floor and row_label and seat_number is not None:
            sold.add((floor, row_label, seat_number))

    _sold_cache = sold
    _sold_cache_time = now
    return sold


def get_orders_by_name(name: str) -> List[dict]:

    target = normalize_text(name)
    if not target:
        return []

    rows = get_all_records()
    grouped = {}

    for row in rows:
        status = normalize_text(row.get("訂單狀態")).lower()
        row_name = normalize_text(row.get("名字"))

        if status not in {"active", "locked"}:
            continue
        if row_name != target:
            continue

        order_id = normalize_text(row.get("訂單ID"))
        dt = normalize_text(row.get("訂單日期與時間"))
        floor = normalize_text(row.get("樓層"))
        row_label = normalize_text(row.get("排數"))
        seat_number = normalize_int(row.get("座位"))
        price = normalize_int(row.get("票價")) or 0
        note = normalize_text(row.get("訂單備註"))
        pickup_open = normalize_bool(row.get("是否開放取票"))
        picked_up = normalize_bool(row.get("是否已取票"))
        payment_done = normalize_bool(row.get("付款狀態"))
        order_status = normalize_text(row.get("訂單狀態")).lower()

        key = (order_id, dt, floor, row_label)

        if key not in grouped:
            grouped[key] = {
                "order_id": order_id,
                "datetime": dt,
                "name": row_name,
                "floor": floor,
                "row_label": row_label,
                "seats": [],
                "price": 0,
                "note": note,
                "pickup_open": pickup_open,
                "picked_up": picked_up,
                "payment_done": payment_done,
                "order_status": order_status,
            }

        if payment_done:
            grouped[key]["payment_done"] = True

        if pickup_open:
            grouped[key]["pickup_open"] = True

        if picked_up:
            grouped[key]["picked_up"] = True

        if order_status:
            grouped[key]["order_status"] = order_status

        if seat_number is not None:
            grouped[key]["seats"].append(seat_number)

        grouped[key]["price"] += price

        if note:
            grouped[key]["note"] = note
        if pickup_open:
            grouped[key]["pickup_open"] = True
        if picked_up:
            grouped[key]["picked_up"] = True

    results = list(grouped.values())

    for item in results:
        item["seats"] = sorted(item["seats"])

    results.sort(
        key=lambda x: (x["datetime"], x["floor"], x["row_label"]),
        reverse=True
    )

    return results


def update_order_note(order_id: str, note: str) -> bool:
    ws = get_worksheet()
    records = ws.get_all_records()

    target_order_id = normalize_text(order_id)
    updated_any = False

    for idx, row in enumerate(records, start=2):  # Google Sheet 第 2 列開始是資料
        current_order_id = normalize_text(row.get("訂單ID"))
        status = normalize_text(row.get("訂單狀態")).lower()

        if current_order_id == target_order_id and status in {"active", "locked"}:
            ws.update_cell(idx, 9, note)  # I 欄 = 訂單備註
            updated_any = True

    return updated_any


def mark_order_deleted(order_id: str) -> bool:
    """
    同一訂單ID的所有列，把訂單狀態改成 deleted
    如果有任一列為 locked，則不刪除
    """
    ws = get_worksheet()
    all_values = ws.get_all_values()

    # 先檢查是否 locked
    for row_idx in range(2, len(all_values) + 1):
        row = all_values[row_idx - 1]
        current_order_id = normalize_text(row[1] if len(row) > 1 else "")  # B
        current_status = normalize_text(row[2] if len(row) > 2 else "").lower()  # C

        if current_order_id == order_id and current_status == "locked":
            return False

    updated_any = False
    for row_idx in range(2, len(all_values) + 1):
        row = all_values[row_idx - 1]
        current_order_id = normalize_text(row[1] if len(row) > 1 else "")  # B

        if current_order_id == order_id:
            ws.update_cell(row_idx, 3, "deleted")  # C
            updated_any = True

    if updated_any:
        clear_caches()

    return updated_any


def update_order_pickup_status(order_id: str, pickup_open: bool = None, picked_up: bool = None) -> bool:
    ws = get_worksheet()
    records = ws.get_all_records()

    target_order_id = normalize_text(order_id)
    updated_any = False

    for idx, row in enumerate(records, start=2):
        current_order_id = normalize_text(row.get("訂單ID"))
        status = normalize_text(row.get("訂單狀態")).lower()

        if current_order_id == target_order_id and status in {"active", "locked"}:
            if pickup_open is not None:
                ws.update_cell(idx, 10, bool(pickup_open))  # J
            if picked_up is not None:
                ws.update_cell(idx, 11, bool(picked_up))    # K
            updated_any = True

    if updated_any:
        clear_caches()

    return updated_any

def admin_search_orders(keyword: str) -> List[dict]:

    target = normalize_text(keyword)
    rows = get_all_records()
    grouped = {}

    for row in rows:
        status = normalize_text(row.get("訂單狀態")).lower()
        if status not in {"active", "locked"}:
            continue

        row_name = normalize_text(row.get("名字"))
        order_id = normalize_text(row.get("訂單ID"))

        if target and target not in row_name and target not in order_id:
            continue

        dt = normalize_text(row.get("訂單日期與時間"))
        floor = normalize_text(row.get("樓層"))
        row_label = normalize_text(row.get("排數"))
        seat_number = normalize_int(row.get("座位"))
        price = normalize_int(row.get("票價")) or 0
        note = normalize_text(row.get("訂單備註"))
        pickup_open = normalize_bool(row.get("是否開放取票"))
        picked_up = normalize_bool(row.get("是否已取票"))
        payment_done = normalize_bool(row.get("付款狀態"))
        ticket_adjusted = normalize_bool(row.get("是否已調票"))

        key = (order_id, dt, floor, row_label)

        if key not in grouped:
            grouped[key] = {
                "order_id": order_id,
                "datetime": dt,
                "name": row_name,
                "floor": floor,
                "row_label": row_label,
                "seats": [],
                "price": 0,
                "note": note,
                "pickup_open": pickup_open,
                "picked_up": picked_up,
                "payment_done": payment_done,
                "order_status": status,
                "ticket_adjusted": ticket_adjusted,
            }

        if seat_number is not None:
            grouped[key]["seats"].append(seat_number)

        grouped[key]["price"] += price

        if note:
            grouped[key]["note"] = note
        if pickup_open:
            grouped[key]["pickup_open"] = True
        if picked_up:
            grouped[key]["picked_up"] = True
        if payment_done:
            grouped[key]["payment_done"] = True
        if ticket_adjusted:
            grouped[key]["ticket_adjusted"] = True

    results = list(grouped.values())

    for item in results:
        item["seats"] = sorted(item["seats"])

    results.sort(
        key=lambda x: (x["datetime"], x["floor"], x["row_label"]),
        reverse=True
    )

    return results

def admin_toggle_lock_status(order_id: str):
    ws = get_worksheet()
    all_values = ws.get_all_values()

    target_rows = []
    current_status = None

    for row_idx in range(2, len(all_values) + 1):
        row = all_values[row_idx - 1]
        current_order_id = normalize_text(row[1] if len(row) > 1 else "")
        status = normalize_text(row[2] if len(row) > 2 else "").lower()

        if normalize_text(current_order_id) == normalize_text(order_id):
            target_rows.append(row_idx)
            current_status = status

    if not target_rows:
        return False, "找不到訂單"

    new_status = "active" if current_status == "locked" else "locked"

    for row_idx in target_rows:
        ws.update_cell(row_idx, 3, new_status)

    clear_caches()
    return True, f"訂單狀態已改為 {new_status}"

def admin_toggle_payment_status(order_id: str):
    ws = get_worksheet()
    all_values = ws.get_all_values()

    target_rows = []
    current_payment = False

    for row_idx in range(2, len(all_values) + 1):
        row = all_values[row_idx - 1]
        current_order_id = normalize_text(row[1] if len(row) > 1 else "")
        payment_done = normalize_bool(row[11] if len(row) > 11 else "")

        if normalize_text(current_order_id) == normalize_text(order_id):
            target_rows.append(row_idx)
            current_payment = payment_done

    if not target_rows:
        return False, "找不到訂單"

    new_value = not current_payment

    for row_idx in target_rows:
        ws.update_cell(row_idx, 12, bool(new_value))

    return True, "付款狀態已更新"
    
def admin_toggle_ticket_adjusted_status(order_id: str):
    ws = get_worksheet()
    all_values = ws.get_all_values()

    target_rows = []
    current_adjusted = False

    for row_idx in range(2, len(all_values) + 1):
        row = all_values[row_idx - 1]
        current_order_id = normalize_text(row[1] if len(row) > 1 else "")
        adjusted = normalize_bool(row[12] if len(row) > 12 else "")  # M

        if normalize_text(current_order_id) == normalize_text(order_id):
            target_rows.append(row_idx)
            current_adjusted = adjusted

    if not target_rows:
        return False, "找不到訂單"

    new_value = not current_adjusted

    for row_idx in target_rows:
        ws.update_cell(row_idx, 13, bool(new_value))  # M

    return True, "調票狀態已更新"
    
def admin_advance_pickup_status(order_id: str):
    ws = get_worksheet()
    all_values = ws.get_all_values()

    target_rows = []
    pickup_open = False
    picked_up = False

    for row_idx in range(2, len(all_values) + 1):
        row = all_values[row_idx - 1]
        current_order_id = normalize_text(row[1] if len(row) > 1 else "")
        if normalize_text(current_order_id) == normalize_text(order_id):
            target_rows.append(row_idx)
            pickup_open = normalize_bool(row[9] if len(row) > 9 else "")   # J
            picked_up = normalize_bool(row[10] if len(row) > 10 else "")   # K

    if not target_rows:
        return False, "找不到訂單"

    if not pickup_open and not picked_up:
        new_open, new_picked = True, False
    elif pickup_open and not picked_up:
        new_open, new_picked = True, True
    else:
        new_open, new_picked = True, True

    for row_idx in target_rows:
        ws.update_cell(row_idx, 10, bool(new_open))   # J
        ws.update_cell(row_idx, 11, bool(new_picked)) # K

    return True, "取票狀態已更新"

def admin_delete_order(order_id: str):
    ws = get_worksheet()
    all_values = ws.get_all_values()

    target_rows = []
    current_status = None

    for row_idx in range(2, len(all_values) + 1):
        row = all_values[row_idx - 1]
        current_order_id = normalize_text(row[1] if len(row) > 1 else "")
        status = normalize_text(row[2] if len(row) > 2 else "").lower()

        if normalize_text(current_order_id) == normalize_text(order_id):
            target_rows.append(row_idx)
            current_status = status

    if not target_rows:
        return False, "找不到訂單"

    if current_status == "locked":
        return False, "已鎖定，無法刪除"

    for row_idx in target_rows:
        ws.update_cell(row_idx, 3, "deleted")

    clear_caches()
    return True, "訂單已刪除"

def load_section_members():
    member_to_section = {}

    try:
        ws = get_config_worksheet("section_members")
        rows = ws.get_all_records()
    except Exception:
        return member_to_section

    for row in rows:
        name = normalize_text(row.get("姓名"))
        section = normalize_text(row.get("聲部"))

        if name and section:
            member_to_section[name] = section

    return member_to_section

def price_to_reward_zone(price: int) -> str:
    if price in {500, 400}:
        return "500"
    if price in {300, 240}:
        return "300"
    if price in {200, 160}:
        return "200"
    return str(price)
    
def build_stats_summary():
    rows = get_all_records()
    member_to_section = load_section_members()
    stats_config = load_stats_config()

    valid_rows = [
        row for row in rows
        if normalize_text(row.get("訂單狀態")).lower() in {"active", "locked"}
    ]

    total_tickets = 0
    paid_tickets = 0
    unpaid_tickets = 0
    paid_amount = 0
    total_amount = 0
    picked_tickets = 0
    unpicked_tickets = 0

    person_ticket_count = defaultdict(int)
    person_zone_counts = defaultdict(lambda: defaultdict(int))
    section_ticket_count = defaultdict(int)
    section_members = defaultdict(lambda: defaultdict(int))

    conductor_count = 0
    fanpage_count = 0
    other_source_count = 0

    for row in valid_rows:
        name = normalize_text(row.get("名字"))
        seat = normalize_int(row.get("座位"))
        price = normalize_int(row.get("票價")) or 0
        payment_done = normalize_bool(row.get("付款狀態"))
        picked_up = normalize_bool(row.get("是否已取票"))

        if seat is None:
            continue

        total_tickets += 1
        total_amount += price

        if payment_done:
            paid_tickets += 1
            paid_amount += price
        else:
            unpaid_tickets += 1

        if picked_up:
            picked_tickets += 1
        else:
            unpicked_tickets += 1

        person_ticket_count[name] += 1
        reward_zone = price_to_reward_zone(price)
        person_zone_counts[name][reward_zone] += 1

        section = member_to_section.get(name, "未分類")
        section_ticket_count[section] += 1
        section_members[section][name] += 1

        if section == "指揮組":
            conductor_count += 1
        elif name == "粉專購票":
            fanpage_count += 1
        elif section == "未分類":
            other_source_count += 1

    ranking = sorted(
        [
            {
                "name": name,
                "section": member_to_section.get(name, "未分類"),
                "tickets": count
            }
            for name, count in person_ticket_count.items()
        ],
        key=lambda x: x["tickets"],
        reverse=True
    )

    section_summary = []
    for section in ["吹管", "彈撥", "拉弦", "低音", "打擊", "特殊來源"]:
        members = section_members.get(section, {})
        member_list = sorted(
            [{"name": n, "tickets": c} for n, c in members.items()],
            key=lambda x: x["tickets"],
            reverse=True
        )
        section_summary.append({
            "section": section,
            "subtotal": section_ticket_count.get(section, 0),
            "members": member_list
        })

    reward_summary = []
    for rule in stats_config["rewards"]:
        qualified = []

        for item in ranking:
            name = item["name"]
            total_person_tickets = person_ticket_count[name]
            zone_counts = person_zone_counts[name]

            ok = True
            for key, need in rule["conditions"].items():
                if key == "TOTAL":
                    if total_person_tickets < need:
                        ok = False
                        break
                else:
                    if zone_counts.get(key, 0) < need:
                        ok = False
                        break

            if ok:
                qualified.append(name)

        reward_summary.append({
            "reward": rule["reward"],
            "requirement": format_reward_conditions(rule["conditions"]),
            "count": len(qualified),
            "names": qualified
        })

    return {
        "overview": {
            "total_tickets": total_tickets,
            "target_tickets": stats_config["target_tickets"],
            "paid_tickets": paid_tickets,
            "unpaid_tickets": unpaid_tickets,
            "paid_amount": paid_amount,
            "total_amount": total_amount,
            "picked_tickets": picked_tickets,
            "unpicked_tickets": unpicked_tickets,
        },
        "special": {
            "conductor_count": conductor_count,
            "fanpage_count": fanpage_count,
            "other_source_count": other_source_count,
        },
        "ranking": ranking,
        "sections": section_summary,
        "rewards": reward_summary,
        "section_chart": [
            {
                "section": item["section"],
                "tickets": item["subtotal"]
            }
            for item in section_summary
        ],
    }

def load_stats_config():
    config = {
        "target_tickets": 0,
        "rewards": []
    }

    try:
        ws = get_config_worksheet("stats_config")
        rows = ws.get_all_records()
    except Exception:
        return config

    for row in rows:
        row_type = normalize_text(row.get("類型")).lower()
        name = normalize_text(row.get("名稱"))
        condition_text = normalize_text(row.get("條件"))

        if row_type == "target":
            try:
                config["target_tickets"] = int(condition_text)
            except Exception:
                config["target_tickets"] = 0
            continue

        if row_type == "reward":
            conditions = {}

            for item in condition_text.split(","):
                item = item.strip()
                if not item or ":" not in item:
                    continue

                key, value = [x.strip() for x in item.split(":", 1)]

                try:
                    conditions[key.upper()] = int(value)
                except Exception:
                    continue

            if name and conditions:
                config["rewards"].append({
                    "reward": name,
                    "conditions": conditions
                })

    return config

def format_reward_conditions(conditions: dict) -> str:
    parts = []

    for key, value in conditions.items():
        if key == "TOTAL":
            parts.append(f" ≥ {value}")
        else:
            parts.append(f"{key}元區 × {value}")

    return " + ".join(parts)

def get_section_members_rows():
    ws = get_config_worksheet("section_members")
    rows = ws.get_all_records()

    return [
        {
            "name": normalize_text(row.get("姓名")),
            "section": normalize_text(row.get("聲部")),
        }
        for row in rows
        if normalize_text(row.get("姓名")) or normalize_text(row.get("聲部"))
    ]


def get_stats_config_rows():
    ws = get_config_worksheet("stats_config")
    rows = ws.get_all_records()

    return [
        {
            "type": normalize_text(row.get("類型")),
            "name": normalize_text(row.get("名稱")),
            "condition": normalize_text(row.get("條件")),
        }
        for row in rows
        if normalize_text(row.get("類型")) or normalize_text(row.get("名稱")) or normalize_text(row.get("條件"))
    ]


def save_section_members_rows(rows):
    ws = get_config_worksheet("section_members")

    values = [["姓名", "聲部"]]

    for row in rows:
        name = normalize_text(row.get("name"))
        section = normalize_text(row.get("section"))

        if not name and not section:
            continue

        values.append([name, section])

    ws.clear()
    ws.update("A1", values)


def save_stats_config_rows(rows):
    ws = get_config_worksheet("stats_config")

    values = [["類型", "名稱", "條件"]]

    for row in rows:
        row_type = normalize_text(row.get("type"))
        name = normalize_text(row.get("name"))
        condition = normalize_text(row.get("condition"))

        if not row_type and not name and not condition:
            continue

        values.append([row_type, name, condition])

    ws.clear()
    ws.update("A1", values)
