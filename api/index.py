import os
import sys

CURRENT_DIR = os.path.dirname(__file__)
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
LIB_DIR = os.path.join(PROJECT_ROOT, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import json
import traceback
from urllib.parse import parse_qs
from threading import Lock

from seat_parser import parse_seat_map
from sheet_repo import (
    append_order_rows,
    build_active_sold_seat_keys,
    mark_order_deleted,
    get_orders_by_name,
    update_order_note,
    update_order_pickup_status,
    admin_search_orders,
    admin_toggle_lock_status,
    admin_toggle_payment_status,
    admin_advance_pickup_status,
    admin_delete_order,
    build_stats_summary,
    get_all_records,
    normalize_text,
)
import time

SEAT_FILE = "data/seat_map.xlsx"
SEAT_CACHE = {
    "seats": None,
    "row_labels": None,
    "loaded_at": 0,
}
SEAT_CACHE_TTL = 60
SECOND_FLOOR_START_ROW = 33
confirm_lock = Lock()


def json_response(start_response, data, status="200 OK"):
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    headers = [
        ("Content-Type", "application/json; charset=utf-8"),
        ("Content-Length", str(len(body))),
        ("Cache-Control", "no-store"),
    ]
    start_response(status, headers)
    return [body]


def get_floor_label_from_excel_row(excel_row: int) -> str:
    return "2樓" if excel_row >= SECOND_FLOOR_START_ROW else "1樓"


def get_cached_seat_map():
    now = time.time()
    if (
        SEAT_CACHE["seats"] is not None
        and SEAT_CACHE["row_labels"] is not None
        and (now - SEAT_CACHE["loaded_at"]) < SEAT_CACHE_TTL
    ):
        return SEAT_CACHE["seats"], SEAT_CACHE["row_labels"]

    seats, row_labels, _ = parse_seat_map(SEAT_FILE)
    SEAT_CACHE["seats"] = seats
    SEAT_CACHE["row_labels"] = row_labels
    SEAT_CACHE["loaded_at"] = now
    return seats, row_labels


def get_request_json(environ):
    try:
        length = int(environ.get("CONTENT_LENGTH") or 0)
    except ValueError:
        length = 0
    raw = environ["wsgi.input"].read(length) if length > 0 else b""
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def route_api(environ, start_response):
    method = environ.get("REQUEST_METHOD", "GET").upper()
    path = environ.get("PATH_INFO", "")
    query = parse_qs(environ.get("QUERY_STRING", ""), keep_blank_values=True)

    if path == "/api/seats" and method == "GET":
        seats, row_labels = get_cached_seat_map()
        active_sold_keys = build_active_sold_seat_keys()

        result_seats = []
        for seat in seats:
            seat_copy = seat.copy()
            seat_id = f"{seat_copy['excel_row']}-{seat_copy['excel_col']}"
            floor = get_floor_label_from_excel_row(seat_copy["excel_row"])
            seat_key = (floor, str(seat_copy["row_label"]), int(seat_copy["seat_number"]))
            seat_copy["seat_id"] = seat_id
            seat_copy["floor"] = floor
            seat_copy["sold"] = seat_key in active_sold_keys
            result_seats.append(seat_copy)

        return json_response(start_response, {
            "seats": result_seats,
            "row_labels": row_labels,
        })

    if path == "/api/stats" and method == "GET":
        return json_response(start_response, {
            "success": True,
            "data": build_stats_summary(),
        })

    if path == "/api/confirm" and method == "POST":
        with confirm_lock:
            data = get_request_json(environ) or {}
            name = (data.get("name") or "").strip()
            selected_seat_ids = data.get("seats", [])

            if not name:
                return json_response(start_response, {"success": False, "message": "請輸入姓名"}, "400 Bad Request")
            if not selected_seat_ids:
                return json_response(start_response, {"success": False, "message": "請選擇座位"}, "400 Bad Request")

            seats, _ = get_cached_seat_map()
            seat_map = {f"{seat['excel_row']}-{seat['excel_col']}": seat for seat in seats}
            active_sold_keys = build_active_sold_seat_keys()
            seat_rows_to_save = []

            for seat_id in selected_seat_ids:
                seat = seat_map.get(seat_id)
                if not seat:
                    return json_response(start_response, {"success": False, "message": f"找不到座位 {seat_id}"}, "400 Bad Request")

                floor = get_floor_label_from_excel_row(seat["excel_row"])
                seat_key = (floor, str(seat["row_label"]), int(seat["seat_number"]))
                if seat_key in active_sold_keys:
                    return json_response(start_response, {"success": False, "message": f"{floor}{seat['row_label']}排{seat['seat_number']}號 已被選走"}, "400 Bad Request")
                if not seat["available"]:
                    return json_response(start_response, {"success": False, "message": f"{floor}{seat['row_label']}排{seat['seat_number']}號 不開放購買"}, "400 Bad Request")

                seat_rows_to_save.append({
                    "floor": floor,
                    "row_label": str(seat["row_label"]),
                    "seat_number": int(seat["seat_number"]),
                    "price": int(seat["price"]),
                })

            order_id = append_order_rows(name=name, seat_rows=seat_rows_to_save)
            return json_response(start_response, {
                "success": True,
                "message": f"訂位成功！訂單編號：{order_id}",
                "order_id": order_id,
            })

    if path == "/api/orders" and method == "GET":
        name = (query.get("name", [""])[0] or "").strip()
        if not name:
            return json_response(start_response, {"success": False, "message": "請輸入姓名", "orders": []}, "400 Bad Request")
        return json_response(start_response, {"success": True, "orders": get_orders_by_name(name)})

    if path == "/api/admin/orders" and method == "GET":
        keyword = (query.get("keyword", [""])[0] or "").strip()
        if not keyword:
            return json_response(start_response, {"success": False, "message": "請輸入姓名或訂單ID", "orders": []}, "400 Bad Request")
        return json_response(start_response, {"success": True, "orders": admin_search_orders(keyword)})

    if path.startswith("/api/orders/"):
        parts = [p for p in path.split("/") if p]
        # api orders {order_id}
        if len(parts) >= 3:
            order_id = parts[2]
            if len(parts) == 4 and parts[3] == "note" and method == "PATCH":
                data = get_request_json(environ) or {}
                note = (data.get("note") or "").strip()
                ok = update_order_note(order_id, note)
                if not ok:
                    return json_response(start_response, {"success": False, "message": "找不到訂單"}, "404 Not Found")
                return json_response(start_response, {"success": True, "message": "備註已更新"})

            if len(parts) == 4 and parts[3] == "pickup" and method == "PATCH":
                data = get_request_json(environ) or {}
                ok = update_order_pickup_status(order_id, pickup_open=data.get("pickup_open"), picked_up=data.get("picked_up"))
                if not ok:
                    return json_response(start_response, {"success": False, "message": "找不到訂單"}, "404 Not Found")
                return json_response(start_response, {"success": True, "message": "取票狀態已更新"})

            if len(parts) == 3 and method == "DELETE":
                rows = get_all_records()
                locked = any(
                    normalize_text(row.get("訂單ID")) == order_id and normalize_text(row.get("訂單狀態")).lower() == "locked"
                    for row in rows
                )
                if locked:
                    return json_response(start_response, {"success": False, "message": "已鎖定，無法刪除"}, "403 Forbidden")
                ok = mark_order_deleted(order_id)
                if not ok:
                    return json_response(start_response, {"success": False, "message": "找不到訂單"}, "404 Not Found")
                return json_response(start_response, {"success": True, "message": "訂單已刪除，座位已重新釋出"})

    if path.startswith("/api/admin/orders/"):
        parts = [p for p in path.split("/") if p]
        if len(parts) >= 4:
            order_id = parts[3]
            if len(parts) == 5 and parts[4] == "lock" and method == "PATCH":
                ok, message = admin_toggle_lock_status(order_id)
                return json_response(start_response, {"success": ok, "message": message}, "200 OK" if ok else "404 Not Found")
            if len(parts) == 5 and parts[4] == "payment" and method == "PATCH":
                ok, message = admin_toggle_payment_status(order_id)
                return json_response(start_response, {"success": ok, "message": message}, "200 OK" if ok else "404 Not Found")
            if len(parts) == 6 and parts[4] == "pickup" and parts[5] == "advance" and method == "PATCH":
                ok, message = admin_advance_pickup_status(order_id)
                return json_response(start_response, {"success": ok, "message": message}, "200 OK" if ok else "404 Not Found")
            if len(parts) == 4 and method == "DELETE":
                ok, message = admin_delete_order(order_id)
                return json_response(start_response, {"success": ok, "message": message}, "200 OK" if ok else "403 Forbidden")

    return json_response(start_response, {"success": False, "message": f"找不到 API：{path}"}, "404 Not Found")


def app(environ, start_response):
    try:
        return route_api(environ, start_response)
    except Exception as e:
        traceback.print_exc()
        return json_response(start_response, {
            "success": False,
            "message": "伺服器發生錯誤",
            "detail": str(e),
        }, "500 Internal Server Error")
