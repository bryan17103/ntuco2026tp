import os
import time
from functools import wraps
from threading import Lock

from dotenv import load_dotenv
from flask import Flask, jsonify, request, session, send_from_directory
from werkzeug.security import check_password_hash, generate_password_hash

from lib.seat_parser import parse_seat_map
from lib.sheet_repo import (
    now_str,
    append_order_rows,
    append_consignment_rows,
    append_consignment_user_row,
    admin_advance_pickup_status,
    admin_delete_order,
    admin_search_orders,
    admin_toggle_lock_status,
    admin_toggle_payment_status,
    admin_toggle_ticket_adjusted_status,
    build_active_sold_seat_keys,
    build_stats_summary,
    build_stats_summary_all,
    delete_consignment_record,
    get_all_consignment_front_records,
    get_all_records,
    get_consignment_records_by_owner_id,
    get_consignment_users_rows,
    get_next_consignment_batch_id,
    get_next_consignment_ids,
    get_next_consignment_owner_id,
    get_order_open,
    get_orders_by_name,
    get_section_members_rows,
    get_stats_config_rows,
    mark_consignment_paid_and_picked_up,
    mark_consignment_sent_to_front,
    mark_order_deleted,
    normalize_name,
    normalize_text,
    save_section_members_rows,
    save_stats_config_rows,
    search_consignment_front_records,
    search_consignment_records_by_audience,
    update_order_note,
    update_order_pickup_status,
    reset_consignment_owner_password,
    get_next_vip_consignment_ids,
)

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")

PROJECT_ROOT = os.path.dirname(os.path.dirname(__file__))

SEAT_FILES = {
    "tp": os.path.join(PROJECT_ROOT, "data", "seat_map_tp.xlsx"),
    "kh": os.path.join(PROJECT_ROOT, "data", "seat_map_kh.xlsx"),
}

SEAT_CACHE = {}
SEAT_CACHE_TTL = 86400

_query_cache = {}
_query_cache_time = {}
_QUERY_CACHE_TTL = 10

SECOND_FLOOR_START_ROW = 33
confirm_lock = Lock()

VALID_CONSIGNMENT_PAYMENT_STATUS = {"paid", "unpaid", "free"}
VALID_CONSIGNMENT_PICKUP_STATUS = {"pending", "sent", "picked_up"}


# ============================================================
# Common Helpers
# ============================================================

def normalize_mode(mode):
    mode = str(mode or "all").strip().lower()

    if mode in ("tp", "taipei"):
        return "tp"

    if mode in ("kh", "kaohsiung", "kaoshiung"):
        return "kh"

    return "all"


def normalize_concert_code(value):
    value = str(value or "").strip().lower()

    if value in ("tp", "taipei"):
        return "tp"

    if value in ("kh", "kaohsiung", "kaoshiung"):
        return "kh"

    return None


def normalize_consignment_payment_status(value):
    value = str(value or "").strip().lower()

    if value in VALID_CONSIGNMENT_PAYMENT_STATUS:
        return value

    return "unpaid"


def check_front_password_from_request(data=None):
    data = data or {}

    # 1. 如果 Session 已經記錄登入成功，直接放行
    if session.get("front_ok"):
        return True, ""

    expected_password = os.getenv("FRONT_PASSWORD", "").strip()
    if not expected_password:
        return False, "尚未設定 FRONT_PASSWORD"

    # 2. 同時嘗試從 Header (不分大小寫) 與 Request Body 中撈取密碼
    input_password = (
        request.headers.get("X-Front-Password")
        or request.headers.get("x-front-password")  # 💡 新增：相容部分瀏覽器自動轉小寫 Header 的問題
        or data.get("front_password")
        or ""
    ).strip()

    # 3. 比對密碼
    if input_password != expected_password:
        return False, "前台密碼錯誤"

    # 4. 比對成功，寫入 Session 紀錄
    session["front_ok"] = True
    return True, ""

def require_admin(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("admin_ok"):
            return jsonify({
                "success": False,
                "message": "你不是票務！"
            }), 401

        return fn(*args, **kwargs)

    return wrapper

# ============================================================
# Page Routes
# ============================================================

@app.route("/")
def serve_index():
    return send_from_directory(PROJECT_ROOT, "index.html")


@app.route("/consignment")
def consignment_page():
    return send_from_directory(PROJECT_ROOT, "consignment.html")


@app.route("/consignment-front")
def consignment_front_page():
    return send_from_directory(PROJECT_ROOT, "consignment-front.html")


# ============================================================
# Seat Map Helpers
# ============================================================

def get_floor_label_from_excel_row(excel_row: int) -> str:
    return "2樓" if excel_row >= SECOND_FLOOR_START_ROW else "1樓"


def get_cached_seat_map(concert_code="tp"):
    now = time.time()

    if concert_code not in SEAT_FILES:
        raise ValueError(f"未知場次：{concert_code}")

    if concert_code not in SEAT_CACHE:
        SEAT_CACHE[concert_code] = {
            "seats": None,
            "row_labels": None,
            "loaded_at": 0,
        }

    cache = SEAT_CACHE[concert_code]

    if (
        cache["seats"] is not None
        and cache["row_labels"] is not None
        and (now - cache["loaded_at"]) < SEAT_CACHE_TTL
    ):
        return cache["seats"], cache["row_labels"]

    seats, row_labels, _ = parse_seat_map(
        SEAT_FILES[concert_code],
        concert_code=concert_code,
    )

    cache["seats"] = seats
    cache["row_labels"] = row_labels
    cache["loaded_at"] = now

    return seats, row_labels


def build_result_seats(concert_code):
    seats, row_labels = get_cached_seat_map(concert_code)
    active_sold_keys = build_active_sold_seat_keys(concert_code)

    result_seats = []

    for seat in seats:
        seat_copy = seat.copy()
        seat_id = f"{seat_copy['excel_row']}-{seat_copy['excel_col']}"

        floor = seat_copy.get("floor")
        if not floor:
            floor = get_floor_label_from_excel_row(seat_copy["excel_row"])

        seat_key = (
            floor,
            str(seat_copy["row_label"]),
            int(seat_copy["seat_number"]),
        )

        seat_copy["seat_id"] = seat_id
        seat_copy["floor"] = floor
        seat_copy["sold"] = seat_key in active_sold_keys

        result_seats.append(seat_copy)

    return result_seats, row_labels


# ============================================================
# Seat Map APIs
# ============================================================

@app.route("/api/tp/seats", methods=["GET"])
def api_tp_seats():
    result_seats, row_labels = build_result_seats("tp")

    return jsonify({
        "success": True,
        "seats": result_seats,
        "row_labels": row_labels,
        "order_open": get_order_open("tp"),
    })


@app.route("/api/kh/seats", methods=["GET"])
def api_kh_seats():
    show_third = request.args.get("show_third", "false") == "true"

    result_seats, row_labels = build_result_seats("kh")

    if not show_third:
        result_seats = [
            seat for seat in result_seats
            if seat.get("floor") != "3樓"
        ]

    return jsonify({
        "success": True,
        "seats": result_seats,
        "row_labels": row_labels,
        "order_open": get_order_open("kh"),
    })


@app.route("/api/debug/kh-seat-count", methods=["GET"])
def debug_kh_seat_count():
    seats, row_labels = get_cached_seat_map("kh")

    zone_counts = {}
    color_counts = {}

    for seat in seats:
        zone_counts[seat["zone"]] = zone_counts.get(seat["zone"], 0) + 1
        color_counts[seat["color"]] = color_counts.get(seat["color"], 0) + 1

    unknown_sample = [
        seat for seat in seats
        if seat["zone"] == "unknown"
    ][:30]

    return jsonify({
        "success": True,
        "seat_count": len(seats),
        "row_label_count": len(row_labels),
        "zone_counts": zone_counts,
        "color_counts": color_counts,
        "unknown_sample": unknown_sample,
    })


# ============================================================
# Order Creation APIs
# ============================================================

def handle_confirm(concert_code):
    with confirm_lock:
        if not get_order_open(concert_code):
            return jsonify({
                "success": False,
                "message": "目前團內購票已截止，無法新增訂單。",
            }), 403

        data = request.get_json(silent=True) or {}

        name = str(data.get("name", "")).strip()
        note = str(data.get("note", "")).strip()
        selected_seat_ids = data.get("seats", [])

        if not name:
            return jsonify({
                "success": False,
                "message": "請輸入姓名",
            }), 400

        if not selected_seat_ids:
            return jsonify({
                "success": False,
                "message": "請選擇座位",
            }), 400

        seats, _ = get_cached_seat_map(concert_code)

        seat_map = {
            f"{seat['excel_row']}-{seat['excel_col']}": seat
            for seat in seats
        }

        active_sold_keys = build_active_sold_seat_keys(concert_code)
        seat_rows_to_save = []

        for seat_id in selected_seat_ids:
            seat = seat_map.get(seat_id)

            if not seat:
                return jsonify({
                    "success": False,
                    "message": f"找不到座位 {seat_id}",
                }), 400

            floor = seat.get("floor")
            if not floor:
                floor = get_floor_label_from_excel_row(seat["excel_row"])

            seat_key = (
                floor,
                str(seat["row_label"]),
                int(seat["seat_number"]),
            )

            if seat_key in active_sold_keys:
                return jsonify({
                    "success": False,
                    "message": (
                        f"{floor}{seat['row_label']}排"
                        f"{seat['seat_number']}號 已被選走"
                    ),
                }), 400

            if not seat["available"]:
                return jsonify({
                    "success": False,
                    "message": (
                        f"{floor}{seat['row_label']}排"
                        f"{seat['seat_number']}號 不開放購買"
                    ),
                }), 400

            seat_rows_to_save.append({
                "floor": floor,
                "row_label": str(seat["row_label"]),
                "seat_number": int(seat["seat_number"]),
                "price": int(seat["price"]),
            })

        order_id = append_order_rows(
            name=name,
            seat_rows=seat_rows_to_save,
            note=note,
            concert_code=concert_code,
        )

        return jsonify({
            "success": True,
            "message": f"訂位成功！訂單編號：{order_id}",
            "order_id": order_id,
        })


@app.route("/api/tp/confirm", methods=["POST"])
def api_tp_confirm():
    return handle_confirm("tp")


@app.route("/api/kh/confirm", methods=["POST"])
def api_kh_confirm():
    return handle_confirm("kh")


# ============================================================
# Public Order Query / Edit APIs
# ============================================================

@app.route("/api/orders", methods=["GET"])
def api_orders():
    name = request.args.get("name", "").strip()
    mode = normalize_mode(request.args.get("mode", "all"))

    if not name:
        return jsonify({
            "success": False,
            "message": "請輸入姓名",
            "orders": [],
        }), 400

    cache_key = f"{name}_{mode}"
    now = time.time()

    if (
        cache_key in _query_cache
        and now - _query_cache_time.get(cache_key, 0) < _QUERY_CACHE_TTL
    ):
        return jsonify(_query_cache[cache_key])

    result = get_orders_by_name(name, concert_code=mode)

    response_data = {
        "success": True,
        "orders": result.get("orders", []),
        "manual_points": result.get("manual_points", 0),
        "total_points": result.get("total_points", 0),
        "identity_code": result.get("identity_code", "5"),
        "identity": result.get("identity", "暫時未分類，請耐心等待"),
        "all_total_points": result.get("all_total_points", 0),
        "discount_amount": result.get("discount_amount", 0),
    }

    _query_cache[cache_key] = response_data
    _query_cache_time[cache_key] = now

    return jsonify(response_data)


@app.route("/api/orders/<order_id>/note", methods=["PATCH"])
def api_update_order_note(order_id):
    data = request.get_json(silent=True) or {}

    note = str(data.get("note", "")).strip()
    floor = request.args.get("floor", "").strip()
    row_label = request.args.get("row_label", "").strip()
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    ok = update_order_note(
        order_id,
        note,
        floor=floor,
        row_label=row_label,
        concert_code=mode,
    )

    if not ok:
        return jsonify({
            "success": False,
            "message": "找不到訂單",
        }), 404

    return jsonify({
        "success": True,
        "message": "備註已更新",
    })


@app.route("/api/orders/<order_id>", methods=["DELETE"])
def api_delete_order(order_id):
    floor = request.args.get("floor", "").strip()
    row_label = request.args.get("row_label", "").strip()
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    rows = get_all_records(mode)

    locked = any(
        normalize_text(row.get("訂單ID")) == order_id
        and normalize_text(row.get("訂單狀態")).lower() == "locked"
        for row in rows
    )

    if locked:
        return jsonify({
            "success": False,
            "message": "已鎖定，無法刪除",
        }), 403

    ok = mark_order_deleted(
        order_id,
        floor=floor,
        row_label=row_label,
        concert_code=mode,
    )

    if not ok:
        return jsonify({
            "success": False,
            "message": "找不到訂單",
        }), 404

    return jsonify({
        "success": True,
        "message": "訂單已刪除，座位已重新釋出",
    })


@app.route("/api/orders/<order_id>/pickup", methods=["PATCH"])
def api_update_order_pickup(order_id):
    data = request.get_json(silent=True) or {}
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    ok = update_order_pickup_status(
        order_id,
        pickup_open=data.get("pickup_open"),
        picked_up=data.get("picked_up"),
        concert_code=mode,
    )

    if not ok:
        return jsonify({
            "success": False,
            "message": "找不到訂單",
        }), 404

    return jsonify({
        "success": True,
        "message": "取票狀態已更新",
    })


# ============================================================
# Consignment Helpers
# ============================================================

def find_consignment_owner(owner_name):
    target = normalize_text(owner_name)

    if not target:
        return None

    rows = get_consignment_users_rows()

    for row in rows:
        if normalize_text(row.get("owner_name")) == target:
            return {
                "created_at": normalize_text(row.get("created_at")),
                "owner_id": normalize_text(row.get("owner_id")),
                "owner_name": normalize_text(row.get("owner_name")),
                "password_hash": normalize_text(row.get("password_hash")),
            }

    return None


def create_consignment_owner(owner_name, password):
    created_at = now_str()
    owner_id = get_next_consignment_owner_id()
    password_hash = generate_password_hash(password)

    row = {
        "created_at": created_at,
        "owner_id": owner_id,
        "owner_name": owner_name,
        "password_hash": password_hash,
    }

    append_consignment_user_row(row)

    return {
        "created_at": created_at,
        "owner_id": owner_id,
        "owner_name": owner_name,
    }


def get_or_create_consignment_owner(owner_name, password, is_new_owner):
    owner_name = normalize_text(owner_name)
    password = str(password or "").strip()

    if not owner_name:
        return None, "請輸入寄票人姓名"

    if not password:
        return None, "請輸入寄票密碼"

    existing_owner = find_consignment_owner(owner_name)

    if is_new_owner:
        if existing_owner:
            return None, "這個寄票人姓名已經建立過密碼，請改選「我已建立密碼」並輸入原密碼"

        owner = create_consignment_owner(owner_name, password)
        return owner, None

    if not existing_owner:
        return None, "找不到這個寄票人。若是第一次寄票，請選擇「第一次寄票」"

    if not check_password_hash(existing_owner["password_hash"], password):
        return None, "寄票密碼錯誤"

    return existing_owner, None


def format_consignment_record(row):
    price = int(row.get("price") or 0)
    quantity = int(row.get("quantity") or 0)

    return {
        "timestamp": row.get("timestamp", ""),
        "consignment_id": row.get("consignment_id", ""),
        "batch_id": row.get("batch_id", ""),
        "owner_name": row.get("owner_name", ""),
        "audience_name": row.get("audience_name", ""),
        "price": price,
        "quantity": quantity,
        "total_amount": price,
        "payment_status": row.get("payment_status", ""),
        "pickup_status": row.get("pickup_status", ""),
        "note": row.get("note", ""),
        "concert_code": row.get("concert_code", ""),
    }

@app.route("/api/consignment/reset-password", methods=["POST"])
def api_reset_consignment_password():
    data = request.get_json(silent=True) or {}

    owner_name = data.get("owner_name", "")
    consignment_id = data.get("consignment_id", "")
    new_password = data.get("new_password", "")

    success, message = reset_consignment_owner_password(
        owner_name=owner_name,
        consignment_id=consignment_id,
        new_password=new_password,
    )

    status_code = 200 if success else 400

    return jsonify({
        "success": success,
        "message": message,
    }), status_code

# ============================================================
# Consignment Public APIs
# ============================================================

@app.route("/api/consignment/submit", methods=["POST"])
def api_consignment_submit():
    try:
        data = request.get_json(silent=True) or {}

        is_new_owner = bool(data.get("is_new_owner"))
        owner_name = normalize_text(data.get("owner_name"))
        password = str(data.get("password") or "").strip()
        confirm_password = str(data.get("confirm_password") or "").strip()
        concert_code = normalize_concert_code(data.get("concert_code"))
        items = data.get("items") or []

        if not concert_code:
            return jsonify({
                "success": False,
                "message": "請選擇場次",
            }), 400

        if is_new_owner and password != confirm_password:
            return jsonify({
                "success": False,
                "message": "兩次輸入的密碼不一致",
            }), 400

        if not isinstance(items, list) or len(items) == 0:
            return jsonify({
                "success": False,
                "message": "請至少新增一筆取票人資料",
            }), 400

        owner, owner_error = get_or_create_consignment_owner(
            owner_name=owner_name,
            password=password,
            is_new_owner=is_new_owner,
        )

        if owner_error:
            return jsonify({
                "success": False,
                "message": owner_error,
            }), 400

        cleaned_items = []

        for item in items:
            audience_name = normalize_text(item.get("audience_name"))
            note = normalize_text(item.get("note"))
            raw_price = str(item.get("price") or "").strip()

            if raw_price == "":
                price = 0
            else:
                try:
                    price = int(float(raw_price))
                except (TypeError, ValueError):
                    price = -1

            payment_status = "free" if price == 0 else "unpaid"

            try:
                quantity = int(float(item.get("quantity")))
            except (TypeError, ValueError):
                quantity = 0

            if not audience_name:
                return jsonify({
                    "success": False,
                    "message": "每一筆都需要填寫取票人姓名！",
                }), 400

            if price < 0:
                return jsonify({
                    "success": False,
                    "message": f"{audience_name} 的金額不正確 :(",
                }), 400

            if quantity <= 0:
                return jsonify({
                    "success": False,
                    "message": f"{audience_name} 的張數不正確 :(",
                }), 400

            cleaned_items.append({
                "audience_name": audience_name,
                "price": price,
                "quantity": quantity,
                "payment_status": payment_status,
                "pickup_status": "pending",
                "note": note,
            })

        timestamp = now_str()
        batch_id = get_next_consignment_batch_id(concert_code)

        is_vip_owner = normalize_name(owner_name) == "貴賓票"

        if is_vip_owner:
            consignment_ids = get_next_vip_consignment_ids(
                concert_code=concert_code,
                count=len(cleaned_items),
            )
        else:
            consignment_ids = get_next_consignment_ids(
                concert_code=concert_code,
                count=len(cleaned_items),
            )

        rows_to_append = []

        for consignment_id, item in zip(consignment_ids, cleaned_items):
            rows_to_append.append({
                "timestamp": timestamp,
                "consignment_id": consignment_id,
                "batch_id": batch_id,
                "owner_id": owner["owner_id"],
                "owner_name": owner["owner_name"],
                "audience_name": item["audience_name"],
                "price": item["price"],
                "quantity": item["quantity"],
                "payment_status": item["payment_status"],
                "pickup_status": item["pickup_status"],
                "note": item["note"],
            })

        append_consignment_rows(concert_code, rows_to_append)

        return jsonify({
            "success": True,
            "message": "寄票資料已送出！",
            "concert_code": concert_code,
            "owner_id": owner["owner_id"],
            "owner_name": owner["owner_name"],
            "batch_id": batch_id,
            "items": [
                {
                    "consignment_id": row["consignment_id"],
                    "audience_name": row["audience_name"],
                    "price": row["price"],
                    "quantity": row["quantity"],
                    "payment_status": row["payment_status"],
                    "pickup_status": row["pickup_status"],
                    "note": row["note"],
                }
                for row in rows_to_append
            ],
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


@app.route("/api/consignment/my-records", methods=["POST"])
def api_consignment_my_records():
    try:
        data = request.get_json(silent=True) or {}

        owner_name = normalize_text(data.get("owner_name"))
        password = str(data.get("password") or "").strip()

        if not owner_name:
            return jsonify({
                "success": False,
                "message": "請輸入寄票人姓名",
            }), 400

        if not password:
            return jsonify({
                "success": False,
                "message": "請輸入寄票密碼",
            }), 400

        owner = find_consignment_owner(owner_name)

        if not owner:
            return jsonify({
                "success": False,
                "message": "找不到這個寄票人",
            }), 404

        if not check_password_hash(owner["password_hash"], password):
            return jsonify({
                "success": False,
                "message": "寄票密碼錯誤",
            }), 401

        records_by_concert = get_consignment_records_by_owner_id(owner["owner_id"])

        tp_records = [
            format_consignment_record(row)
            for row in records_by_concert.get("tp", [])
        ]

        kh_records = [
            format_consignment_record(row)
            for row in records_by_concert.get("kh", [])
        ]

        return jsonify({
            "success": True,
            "owner_name": owner["owner_name"],
            "owner_id": owner["owner_id"],
            "records": {
                "tp": tp_records,
                "kh": kh_records,
            },
            "summary": {
                "tp_count": len(tp_records),
                "kh_count": len(kh_records),
                "total_count": len(tp_records) + len(kh_records),
            },
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


@app.route("/api/consignment/delete", methods=["POST"])
def api_consignment_delete():
    try:
        data = request.get_json(silent=True) or {}

        owner_name = normalize_text(data.get("owner_name"))
        password = str(data.get("password") or "").strip()
        concert_code = normalize_concert_code(data.get("concert_code"))
        consignment_id = normalize_text(data.get("consignment_id"))

        if not owner_name:
            return jsonify({
                "success": False,
                "message": "請輸入寄票人姓名",
            }), 400

        if not password:
            return jsonify({
                "success": False,
                "message": "請輸入寄票密碼",
            }), 400

        if not concert_code:
            return jsonify({
                "success": False,
                "message": "缺少場次資訊",
            }), 400

        if not consignment_id:
            return jsonify({
                "success": False,
                "message": "缺少取票編號",
            }), 400

        owner = find_consignment_owner(owner_name)

        if not owner:
            return jsonify({
                "success": False,
                "message": "找不到這個寄票人",
            }), 404

        if not check_password_hash(owner["password_hash"], password):
            return jsonify({
                "success": False,
                "message": "寄票密碼錯誤",
            }), 401

        success, message = delete_consignment_record(
            concert_code=concert_code,
            consignment_id=consignment_id,
            owner_id=owner["owner_id"],
        )

        if not success:
            return jsonify({
                "success": False,
                "message": message,
            }), 400

        return jsonify({
            "success": True,
            "message": message,
            "consignment_id": consignment_id,
            "concert_code": concert_code,
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


@app.route("/api/consignment/audience-lookup", methods=["POST"])
def api_consignment_audience_lookup():
    try:
        data = request.get_json(silent=True) or {}

        concert_code = normalize_concert_code(data.get("concert_code"))
        audience_name = normalize_text(data.get("audience_name"))

        if not concert_code:
            return jsonify({
                "success": False,
                "message": "請選擇場次",
            }), 400

        if not audience_name:
            return jsonify({
                "success": False,
                "message": "請輸入取票人姓名",
            }), 400

        records = search_consignment_records_by_audience(
            concert_code=concert_code,
            audience_name=audience_name,
        )

        formatted_records = [
            format_consignment_record(row)
            for row in records
        ]

        return jsonify({
            "success": True,
            "concert_code": concert_code,
            "audience_name": audience_name,
            "records": formatted_records,
            "count": len(formatted_records),
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


# ============================================================
# Consignment Front Desk APIs
# ============================================================

@app.route("/api/consignment-front/login", methods=["POST"])
def api_consignment_front_login():
    data = request.get_json(silent=True) or {}

    ok, message = check_front_password_from_request(data)

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 401

    session["front_ok"] = True

    return jsonify({
        "success": True,
        "message": "登入成功",
    })


@app.route("/api/consignment-front/search", methods=["POST"])
def api_consignment_front_search():
    data = request.get_json(silent=True) or {}

    ok, message = check_front_password_from_request(data)

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
            "records": [],
        }), 401

    concert_code = normalize_concert_code(data.get("concert_code"))
    keyword = normalize_text(data.get("keyword"))

    if not concert_code:
        return jsonify({
            "success": False,
            "message": "請選擇場次",
            "records": [],
        }), 400

    if not keyword:
        return jsonify({
            "success": False,
            "message": "請輸入取票編號、取票人姓名或寄票人姓名",
            "records": [],
        }), 400

    records = search_consignment_front_records(concert_code, keyword)

    return jsonify({
        "success": True,
        "records": records,
        "count": len(records),
    })


@app.route("/api/consignment-front/all", methods=["POST"])
def api_consignment_front_all():
    data = request.get_json(silent=True) or {}

    ok, message = check_front_password_from_request(data)

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
            "records": [],
        }), 401

    concert_code = normalize_concert_code(data.get("concert_code"))

    if not concert_code:
        return jsonify({
            "success": False,
            "message": "請選擇場次",
            "records": [],
        }), 400

    records = get_all_consignment_front_records(concert_code)
    keyword = normalize_text(data.get("keyword"))

    if keyword:
        target = normalize_name(keyword)

        records = [
            record for record in records
            if target in normalize_name(record.get("owner_name"))
        ]

    return jsonify({
        "success": True,
        "records": records,
        "count": len(records),
    })


@app.route("/api/consignment-front/paid-picked-up", methods=["PATCH"])
def api_consignment_front_paid_picked_up():
    data = request.get_json(silent=True) or {}

    ok, message = check_front_password_from_request(data)

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 401

    concert_code = normalize_concert_code(data.get("concert_code"))
    consignment_id = normalize_text(data.get("consignment_id"))
    front_note = data.get("front_note")

    if not concert_code:
        return jsonify({
            "success": False,
            "message": "請選擇場次",
        }), 400

    if not consignment_id:
        return jsonify({
            "success": False,
            "message": "缺少取票編號",
        }), 400

    success, message = mark_consignment_paid_and_picked_up(
        concert_code=concert_code,
        consignment_id=consignment_id,
        front_note=front_note
    )

    return jsonify({
        "success": success,
        "message": message,
    }), 200 if success else 400

@app.route("/api/consignment-front/update-note", methods=["PATCH"])
def api_consignment_front_update_note():
    data = request.get_json(silent=True) or {}

    # 驗證前台密碼權限
    ok, message = check_front_password_from_request(data)
    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 401

    concert_code = normalize_concert_code(data.get("concert_code"))
    consignment_id = normalize_text(data.get("consignment_id"))
    front_note = str(data.get("front_note", "")).strip()

    if not concert_code:
        return jsonify({
            "success": False,
            "message": "請選擇場次",
        }), 400

    if not consignment_id:
        return jsonify({
            "success": False,
            "message": "缺少取票編號",
        }), 400

    # 呼叫先前在 sheet_repo.py 寫好的 update_consignment_front_note 函式
    from lib.sheet_repo import update_consignment_front_note
    success, message = update_consignment_front_note(
        concert_code=concert_code,
        consignment_id=consignment_id,
        front_note=front_note
    )

    return jsonify({
        "success": success,
        "message": message,
    }), 200 if success else 400

@app.route("/api/consignment-front/sent", methods=["PATCH"])
def api_consignment_front_sent():
    data = request.get_json(silent=True) or {}

    ok, message = check_front_password_from_request(data)

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 401

    concert_code = normalize_concert_code(data.get("concert_code"))
    consignment_id = normalize_text(data.get("consignment_id"))

    if not concert_code:
        return jsonify({
            "success": False,
            "message": "請選擇場次",
        }), 400

    if not consignment_id:
        return jsonify({
            "success": False,
            "message": "缺少取票編號",
        }), 400

    success, message = mark_consignment_sent_to_front(
        concert_code=concert_code,
        consignment_id=consignment_id,
    )

    return jsonify({
        "success": success,
        "message": message,
    }), 200 if success else 400


# ============================================================
# Admin Auth APIs
# ============================================================

@app.route("/api/admin/login", methods=["POST"])
def api_admin_login():
    try:
        data = request.get_json(silent=True) or {}
        password = str(data.get("password", ""))

        admin_password = os.environ.get("ADMIN_PASSWORD")

        if not admin_password:
            return jsonify({
                "success": False,
                "message": "後台密碼尚未設定",
            }), 500

        if password == admin_password:
            session["admin_ok"] = True
            return jsonify({"success": True})

        return jsonify({
            "success": False,
            "message": "你不是票務 :(",
        }), 401

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


# ============================================================
# Admin Order Management APIs
# ============================================================

@app.route("/api/admin/toggle-order-open", methods=["POST"])
@require_admin
def api_admin_toggle_order_open():
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    rows = get_stats_config_rows()
    target_name = f"order_open_{mode}"
    target_index = None

    for i, row in enumerate(rows):
        row_type = str(row.get("type", "")).strip()
        row_name = str(row.get("name", "")).strip()

        if row_type == "open" and row_name == target_name:
            target_index = i
            break

    if target_index is None:
        rows.append({
            "type": "open",
            "name": target_name,
            "condition": "false",
        })
        new_value = False
    else:
        current = str(
            rows[target_index].get("condition", "true")
        ).strip().lower() == "true"
        new_value = not current
        rows[target_index]["condition"] = "true" if new_value else "false"

    save_stats_config_rows(rows)

    return jsonify({
        "success": True,
        "order_open": new_value,
        "mode": mode,
    })


@app.route("/api/admin/orders", methods=["GET"])
@require_admin
def api_admin_orders():
    keyword = request.args.get("keyword", "").strip()
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    orders = admin_search_orders(keyword, concert_code=mode)

    return jsonify({
        "success": True,
        "mode": mode,
        "orders": orders,
    })


@app.route("/api/admin/orders/<order_id>/ticket-adjusted", methods=["PATCH"])
@require_admin
def api_admin_ticket_adjusted(order_id):
    floor = request.args.get("floor", "").strip()
    row_label = request.args.get("row_label", "").strip()
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    ok, message = admin_toggle_ticket_adjusted_status(
        order_id,
        floor=floor,
        row_label=row_label,
        concert_code=mode,
    )

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 404

    return jsonify({
        "success": True,
        "message": message,
    })


@app.route("/api/admin/orders/<order_id>/pickup/advance", methods=["PATCH"])
@require_admin
def api_admin_pickup_advance(order_id):
    floor = request.args.get("floor", "").strip()
    row_label = request.args.get("row_label", "").strip()
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    ok, message = admin_advance_pickup_status(
        order_id,
        floor=floor,
        row_label=row_label,
        concert_code=mode,
    )

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 404

    return jsonify({
        "success": True,
        "message": message,
    })


@app.route("/api/admin/orders/<order_id>/lock", methods=["PATCH"])
@require_admin
def api_admin_lock(order_id):
    floor = request.args.get("floor", "").strip()
    row_label = request.args.get("row_label", "").strip()
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    ok, message = admin_toggle_lock_status(
        order_id,
        floor=floor,
        row_label=row_label,
        concert_code=mode,
    )

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 404

    return jsonify({
        "success": True,
        "message": message,
    })


@app.route("/api/admin/orders/<order_id>/payment", methods=["PATCH"])
@require_admin
def api_admin_payment(order_id):
    floor = request.args.get("floor", "").strip()
    row_label = request.args.get("row_label", "").strip()
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    ok, message = admin_toggle_payment_status(
        order_id,
        floor=floor,
        row_label=row_label,
        concert_code=mode,
    )

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 404

    return jsonify({
        "success": True,
        "message": message,
    })


@app.route("/api/admin/orders/<order_id>", methods=["DELETE"])
@require_admin
def api_admin_delete(order_id):
    floor = request.args.get("floor", "").strip()
    row_label = request.args.get("row_label", "").strip()
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        mode = "tp"

    ok, message = admin_delete_order(
        order_id,
        floor=floor,
        row_label=row_label,
        concert_code=mode,
    )

    if not ok:
        return jsonify({
            "success": False,
            "message": message,
        }), 403

    return jsonify({
        "success": True,
        "message": message,
    })


# ============================================================
# Admin Edit APIs
# ============================================================

@app.route("/api/edit/config", methods=["GET"])
@require_admin
def api_edit_get_config():
    try:
        return jsonify({
            "success": True,
            "section_members": get_section_members_rows(),
            "stats_config": get_stats_config_rows(),
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


@app.route("/api/edit/section-members", methods=["PUT"])
@require_admin
def api_edit_section_members():
    try:
        data = request.get_json(silent=True) or {}
        rows = data.get("rows", [])

        save_section_members_rows(rows)

        return jsonify({
            "success": True,
            "message": "聲部名單已更新",
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


@app.route("/api/edit/stats-config", methods=["PUT"])
@require_admin
def api_edit_stats_config():
    try:
        data = request.get_json(silent=True) or {}
        rows = data.get("rows", [])

        save_stats_config_rows(rows)

        return jsonify({
            "success": True,
            "message": "統計設定已更新",
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e),
        }), 500


# ============================================================
# Stats APIs
# ============================================================

@app.route("/api/stats", methods=["GET"])
def api_stats():
    mode = normalize_mode(request.args.get("mode", "tp"))

    if mode == "all":
        return jsonify({
            "success": True,
            "data": build_stats_summary_all(),
        })

    return jsonify({
        "success": True,
        "data": build_stats_summary(concert_code=mode),
    })


# ============================================================
# Static File Fallback
# Keep this at the very bottom.
# ============================================================

@app.route("/<path:filename>")
def serve_static_file(filename):
    return send_from_directory(PROJECT_ROOT, filename)
