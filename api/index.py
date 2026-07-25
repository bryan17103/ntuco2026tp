import os
import time
import io
import csv
import re
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
    get_spreadsheet,
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

    # 💡 ⚡ 新增：預先在後端算出該使用者達成的獎勵清單，避免前端二次 fetch 造成 delay
    unlocked_rewards = []
    try:
        from lib.sheet_repo import build_stats_summary_all
        all_stats = build_stats_summary_all()
        all_rewards = all_stats.get("rewards", [])
        target_name = normalize_name(name)

        for rule in all_rewards:
            reward_name = rule.get("reward", "")
            names = rule.get("names", [])
            # 檢查名字是否在合規名單中
            if target_name in [normalize_name(n) for n in names]:
                unlocked_rewards.append({
                    "name": reward_name,
                    "requirement": rule.get("requirement", ""),
                })
    except Exception as e:
        print(f"計算解鎖獎勵失敗: {e}")

    response_data = {
        "success": True,
        "orders": result.get("orders", []),
        "manual_points": result.get("manual_points", 0),
        "total_points": result.get("total_points", 0),
        "identity_code": result.get("identity_code", "5"),
        "identity": result.get("identity", "暫時未分類，請耐心等待"),
        "all_total_points": result.get("all_total_points", 0),
        "discount_amount": result.get("discount_amount", 0),
        "unlocked_rewards": unlocked_rewards  # 💡 直接回傳已解鎖的獎勵
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
# 1. 派發調票網頁路由
# ============================================================
@app.route("/booking")
def booking_page():
    if not session.get("admin_ok"):
        return "<script>alert('你不是票務！'); window.location.href='/';</script>", 401
    return send_from_directory(PROJECT_ROOT, "booking.html")


# ============================================================
# 2. 處理調票 CSV 萃取、跨表比對、回填與寫入 booking 工作表 API
# ============================================================
@app.route("/api/admin/booking/import", methods=["POST"])
@require_admin
def api_admin_import_booking():
    try:
        batch_count = request.form.get("batch_count", "").strip()
        if not batch_count:
            return jsonify({"success": False, "message": "請先填寫這是第幾次調票！"}), 400
            
        if "file" not in request.files:
            return jsonify({"success": False, "message": "未偵測到上傳檔案"}), 400
            
        file = request.files["file"]
        if file.filename == "":
            return jsonify({"success": False, "message": "未選取檔案"}), 400

        # 解析 CSV 檔案資料
        file_bytes = file.read()
        try:
            csv_text = file_bytes.decode("utf-8-sig") 
        except UnicodeDecodeError:
            csv_text = file_bytes.decode("cp950", errors="ignore")
            
        csv_file = io.StringIO(csv_text)
        reader = csv.DictReader(csv_file)
        
        reader.fieldnames = [f.strip().replace("\ufeff", "") if f else "" for f in reader.fieldnames]
        
        # OPENTIX 欄位名稱對照字典
        field_map = {
            "order_no": next((f for f in reader.fieldnames if "訂單號碼" in f), "訂單號碼"),
            "seat_raw": next((f for f in reader.fieldnames if "座位" in f), "座位"),
            "order_date": next((f for f in reader.fieldnames if "訂單日期" in f), "訂單日期"),
            "location": next((f for f in reader.fieldnames if "地點" in f), "地點"),
            "price": next((f for f in reader.fieldnames if "售價" in f), "售價"),
        }
        
        # 初始化 Google Sheet 連線
        spreadsheet = get_spreadsheet()
        tp_ws = spreadsheet.worksheet("2026Summer_Taipei")
        kh_ws = spreadsheet.worksheet("2026Summer_Kaohsiung")
        booking_ws = spreadsheet.worksheet("booking")
        
        tp_rows = tp_ws.get_all_values()
        kh_rows = kh_ws.get_all_values()
        
        #排數欄位改為 ([0-9A-Za-z]+) 完美兼容 A2排, B14排 等英文編號排數
        seat_pattern = re.compile(r"^(.*?)\s*-\s*([0-9A-Za-z]+)排\s*-\s*([0-9]+)號")

        def build_seat_map(sheet_values):
            mapping = {}
            if len(sheet_values) < 2:
                return mapping
            headers = [h.strip() for h in sheet_values[0]]
            
            try:
                idx_floor = headers.index("樓層")
                idx_row = headers.index("排數")
                idx_seat = headers.index("座位")
                idx_name = headers.index("名字")
            except ValueError:
                idx_floor, idx_row, idx_seat, idx_name = 4, 5, 6, 3 
                
            for r_idx, row in enumerate(sheet_values[1:], start=2):
                if len(row) <= max(idx_floor, idx_row, idx_seat, idx_name):
                    continue
                f_val = str(row[idx_floor]).strip()
                r_val = str(row[idx_row]).strip()
                s_val = str(row[idx_seat]).strip()
                n_val = str(row[idx_name]).strip()
                
                mapping[(f_val, r_val, s_val)] = {
                    "row_num": r_idx,
                    "buyer_name": n_val
                }
            return mapping

        tp_seat_map = build_seat_map(tp_rows)
        kh_seat_map = build_seat_map(kh_rows)

        booking_entries = []
        tp_updates = []
        kh_updates = []

        for row in reader:
            order_no = row.get(field_map["order_no"], "").strip()
            seat_raw = row.get(field_map["seat_raw"], "").strip()
            order_date = row.get(field_map["order_date"], "").strip()
            location = row.get(field_map["location"], "").strip()
            price = row.get(field_map["price"], "").strip()
            
            if not order_no or not seat_raw:
                continue
                
            buyer_name = ""  
            target_map = None
            update_list = None
            
            if "衛武營" in location:
                target_map = kh_seat_map
                update_list = kh_updates
            elif "中山堂" in location:
                target_map = tp_seat_map
                update_list = tp_updates
                
            # 解析座位字串
            match = seat_pattern.search(seat_raw)
            if match and target_map is not None:
                raw_floor = match.group(1).strip() # 例如 "2樓7號門"
                parsed_row = str(match.group(2)).strip() # 💡 直接取字串 "A2"，不再轉 int，完美保留英文字母
                parsed_seat = str(int(match.group(3))) # 座位號碼轉 int 去除前導 0 轉回字串 "17"
                
                # 模糊比對樓層、排數與座位
                matched_seat_info = None
                for (f_k, r_k, s_k), info in target_map.items():
                    # 比對排數 (r_k == parsed_row) 與 座位 (s_k == parsed_seat)
                    # 並且透過包含關係比對樓層（例如 "2樓" 是否包含在 "2樓7號門" 中）
                    if r_k == parsed_row and s_k == parsed_seat and (f_k in raw_floor or raw_floor in f_k):
                        matched_seat_info = info
                        break
                
                if matched_seat_info:
                    buyer_name = matched_seat_info["buyer_name"]
                    target_row_num = matched_seat_info["row_num"]
                    
                    update_list.append({
                        'range': f'O{target_row_num}',
                        'values': [[order_no]]
                    })
            
            booking_entries.append([
                order_no,                  
                seat_raw,                  
                order_date,                
                location,                  
                int(price) if price.isdigit() else price,  
                buyer_name,                
                "已調",                    
                f"第{batch_count}次調票"    
            ])

        if booking_entries:
            booking_ws.append_rows(booking_entries, value_input_option="USER_ENTERED")
        if tp_updates:
            tp_ws.batch_update(tp_updates, value_input_option="USER_ENTERED")
        if kh_updates:
            kh_ws.batch_update(kh_updates, value_input_option="USER_ENTERED")

        return jsonify({
            "success": True, 
            "message": f"成功匯入 {len(booking_entries)} 筆調票資料！已同步串聯名字比對與回填 O 欄。",
            "data": booking_entries
        })

    except Exception as e:
        return jsonify({"success": False, "message": f"匯入失敗，錯誤原因：{str(e)}"}), 500

# ============================================================
# 3. 終極升級版：批次同步更新 booking 工作表中的「情況」與「購買者姓名」
# ============================================================
@app.route("/api/admin/booking/batch-update-status", methods=["PATCH"])
@require_admin
def api_admin_booking_batch_update_status():
    global _booking_cache, _booking_cache_time
    try:
        data = request.get_json(silent=True) or {}
        updates = data.get("updates", []) # [{order_no, seat_raw, status, buyer_name}, ...]
        
        if not updates:
            return jsonify({"success": True, "message": "沒有偵測到任何變更項目"}), 200
            
        spreadsheet = get_spreadsheet()
        booking_ws = spreadsheet.worksheet("booking")
        booking_rows = booking_ws.get_all_values()
        
        if len(booking_rows) < 2:
            return jsonify({"success": False, "message": "工作表無資料，無法更新"}), 404
            
        row_map = {}
        for idx, row in enumerate(booking_rows[1:], start=2):
            if len(row) >= 2:
                row_map[(row[0].strip(), row[1].strip())] = idx
                
        gspread_updates = []
        success_count = 0
        
        for item in updates:
            order_no = item.get("order_no", "").strip()
            seat_raw = item.get("seat_raw", "").strip()
            new_status = item.get("status", "").strip()
            new_buyer_name = item.get("buyer_name", "").strip() 
            
            key = (order_no, seat_raw)
            if key in row_map:
                row_idx = row_map[key]
        
                gspread_updates.append({
                    'range': f'F{row_idx}:G{row_idx}',
                    'values': [[new_buyer_name, new_status]]
                })
                success_count += 1
                
        if gspread_updates:
            booking_ws.batch_update(gspread_updates, value_input_option="USER_ENTERED")
            
        _booking_cache = None
        _booking_cache_time = 0
            
        return jsonify({
            "success": True, 
            "message": f"成功將 {success_count} 筆資料儲存至 Google Sheet 中！"
        })
        
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

# ============================================================
# 4. 取得歷史調票紀錄 API
# ============================================================
@app.route("/api/admin/booking/records", methods=["GET"])
def api_admin_get_booking_records():
    try:
        spreadsheet = get_spreadsheet()
        booking_ws = spreadsheet.worksheet("booking")
        booking_rows = booking_ws.get_all_values()
        
        if len(booking_rows) < 2:
            return jsonify({"success": True, "data": []})
            
        # 排除第一行的表頭，並將其餘資料倒序排列（讓最新調票的紀錄顯示在最上面）
        history_data = booking_rows[1:]
        history_data.reverse() 
        
        return jsonify({
            "success": True,
            "data": history_data
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"無法載入歷史紀錄：{str(e)}"}), 500
# ============================================================
# 5. 【完全體修復版】補上 defaultdict 匯入，完美支援單筆訂單內有多張座位的更新
# ============================================================
@app.route("/api/admin/orders/batch-update-workflow", methods=["PATCH"])
@require_admin
def api_admin_orders_batch_update_workflow():
    try:
        mode = normalize_mode(request.args.get("mode", "tp"))
        if mode == "all":
            mode = "tp"
            
        data = request.get_json(silent=True) or {}
        updates = data.get("updates", []) # [{order_id, floor, row_label, pickup_status, payment_status, order_status}, ...]
        
        if not updates:
            return jsonify({"success": True, "message": "沒有偵測到任何變更項目"}), 200
            
        # 💡 修正點：匯入 collections 模組中的 defaultdict，徹底解決未定義錯誤
        from collections import defaultdict
        from lib.sheet_repo import get_worksheet, get_header_col_map
        ws = get_worksheet(mode)
        
        all_values = ws.get_all_values()
        col_map = get_header_col_map(ws)
        
        # 收集所有符合的 row_num
        row_matching_map = defaultdict(list)
        for r_idx in range(2, len(all_values) + 1):
            row = all_values[r_idx - 1]
            c_order_id = normalize_text(row[1] if len(row) > 1 else "")
            c_floor = normalize_text(row[4] if len(row) > 4 else "")
            c_row_label = normalize_text(row[5] if len(row) > 5 else "")
            
            key = (c_order_id, c_floor, c_row_label)
            row_matching_map[key].append(r_idx)

        gspread_updates = []
        dt_now = now_str()

        for item in updates:
            order_id = item.get("order_id", "").strip()
            floor = item.get("floor", "").strip()
            row_label = item.get("row_label", "").strip()
            p_status = item.get("pickup_status", "").strip()    # 未調票 / 已調票 / 開放取票 / 已取票
            m_status = item.get("payment_status", "").strip()   # 未付 / 已付 [現金] / 已付 [轉帳]
            o_status = item.get("order_status")                 # locked / active
            
            key = (order_id, floor, row_label)
            
            if key in row_matching_map:
                for row_num in row_matching_map[key]:
                    orig_row = all_values[row_num - 1]
                    orig_status = normalize_text(orig_row[2] if len(orig_row) > 2 else "").lower()
                    
                    val_open = "FALSE"
                    val_picked = "FALSE"
                    val_payment = "FALSE"
                    val_adjusted = "FALSE"
                    
                    val_p_time = ""
                    val_q_mode = ""
                    val_r_time = ""

                    # 鎖定狀態
                    val_order_status = o_status if o_status is not None else orig_status

                    # 付款大底
                    if m_status != "未付":
                        val_payment = "TRUE"
                        val_open = "TRUE"
                        val_adjusted = "TRUE" 
                        
                        if m_status == "已付 [現金]":
                            val_p_time = dt_now
                            val_q_mode = "cash"
                        elif m_status == "已付 [轉帳]":
                            val_p_time = dt_now
                            val_q_mode = "bank"

                    # 取票狀態疊加
                    if p_status == "已調票":
                        val_adjusted = "TRUE"
                    elif p_status == "開放取票":
                        val_open = "TRUE"
                    elif p_status == "已取票":
                        val_open = "TRUE"
                        val_picked = "TRUE"
                        val_r_time = dt_now

                    # 為該列追加更新任務
                    gspread_updates.append({'range': f'C{row_num}', 'values': [[val_order_status]]})
                    gspread_updates.append({'range': f'J{row_num}', 'values': [[val_open]]})
                    gspread_updates.append({'range': f'K{row_num}', 'values': [[val_picked]]})
                    gspread_updates.append({'range': f'L{row_num}', 'values': [[val_payment]]})
                    gspread_updates.append({'range': f'M{row_num}', 'values': [[val_adjusted]]})
                    gspread_updates.append({'range': f'P{row_num}:R{row_num}', 'values': [[val_p_time, val_q_mode, val_r_time]]})

        if gspread_updates:
            ws.batch_update(gspread_updates, value_input_option="USER_ENTERED")
            
        from lib.sheet_repo import clear_caches
        clear_caches(mode)
        
        return jsonify({"success": True, "message": "取票與付款狀態（含多座位關聯與P,Q,R時間）已整批成功儲存！"})
        
    except Exception as e:
        return jsonify({"success": False, "message": f"批次更新失敗：{str(e)}"}), 500

# ============================================================
# 6. 修正版：每日帳務對帳明細 API（改用 get_worksheet）
# ============================================================
@app.route("/api/admin/orders/daily-finance", methods=["GET"])
@require_admin
def api_admin_daily_finance():
    try:
        mode = normalize_mode(request.args.get("mode", "tp"))
        if mode == "all":
            mode = "tp"
            
        target_date = request.args.get("date", "").strip() # 格式例如 "2026/07/20"
        
        # 💡 修正點：改為呼叫 sheet_repo 內建的 get_worksheet，拔除未定義的 WORKSHEET_NAMES
        from lib.sheet_repo import get_worksheet, get_header_col_map
        ws = get_worksheet(mode)
        
        all_values = ws.get_all_values()
        col_map = get_header_col_map(ws)
        
        # 抓取必要欄位索引
        idx_order_id = 1
        idx_name = 3
        idx_floor = 4
        idx_row = 5
        idx_seat = 6
        idx_price = 7
        idx_note = 8
        
        idx_p_time = 15 # Column P (payment_time)
        idx_q_mode = 16 # Column Q (payment_mode)
        
        # 1. 收集所有有記帳的「不重複日期」供前端選單使用
        available_dates = set()
        for row in all_values[1:]:
            if len(row) > idx_p_time and row[idx_p_time]:
                date_part = row[idx_p_time].split(" ")[0]
                available_dates.add(date_part)
                
        sorted_dates = sorted(list(available_dates), reverse=True) # 最新日期排最前
        
        if not target_date and sorted_dates:
            target_date = sorted_dates[0]
            
        # 2. 開始撈取目標日期的對帳明細並計算總額
        finance_details = []
        cash_total = 0
        bank_total = 0
        
        if target_date:
            for row in all_values[1:]:
                if len(row) > idx_q_mode and row[idx_p_time]:
                    p_time_val = row[idx_p_time]
                    if p_time_val.startswith(target_date):
                        q_mode_val = row[idx_q_mode].strip().lower() # cash / bank
                        price_val = int(float(row[idx_price])) if row[idx_price] else 0
                        
                        if q_mode_val == "cash":
                            cash_total += price_val
                            mode_text = "現金"
                        elif q_mode_val == "bank":
                            bank_total += price_val
                            mode_text = "轉帳"
                        else:
                            mode_text = "未知"
                            
                        finance_details.append({
                            "time": p_time_val.split(" ")[1] if " " in p_time_val else p_time_val,
                            "order_id": row[idx_order_id],
                            "name": row[idx_name],
                            "seat": f"{row[idx_floor]}{row[idx_row]}排{row[idx_seat]}號",
                            "price": price_val,
                            "mode": mode_text,
                            "note": row[idx_note] if len(row) > idx_note else ""
                        })
                        
        finance_details.sort(key=lambda x: x["time"], reverse=True)
        
        return jsonify({
            "success": True,
            "target_date": target_date,
            "available_dates": sorted_dates,
            "summary": {
                "cash_total": cash_total,
                "bank_total": bank_total,
                "grand_total": cash_total + bank_total,
                "count": len(finance_details)
            },
            "details": finance_details
        })
        
    except Exception as e:
        return jsonify({"success": False, "message": f"讀取每日帳務失敗：{str(e)}"}), 500
    
# ============================================================
# 7. 【升級版】匯出折抵金額清單 (含下載時間與精細排序)
# ============================================================
@app.route("/api/admin/export-discount-list", methods=["GET"])
@require_admin
def api_admin_export_discount_list():
    try:
        import io
        import csv
        from datetime import datetime
        from collections import defaultdict
        from flask import make_response
        from lib.sheet_repo import (
            get_config_worksheet,
            load_section_members,
            get_all_records,
            group_order_rows,
            calc_points_from_orders,
            calc_discount_amount,
            normalize_name,
            normalize_identity,
            parse_manual_points_string,
            now_str,
            TAIPEI_TZ,
        )

        # 1. 取得當前時間與格式化時間字串
        dt_now = datetime.now(TAIPEI_TZ)
        time_str_display = dt_now.strftime("%Y/%m/%d %H:%M:%S")
        time_str_file = dt_now.strftime("%Y%m%d_%H%M")

        # 2. 取得所有的聲部成員
        member_map = load_section_members("tp")
        
        # 3. 抓取台北場與高雄場的所有有效訂單
        tp_records = get_all_records("tp")
        kh_records = get_all_records("kh")

        # 4. 按姓名歸接所有人的訂單
        tp_orders_by_name = defaultdict(list)
        for order in group_order_rows(tp_records):
            tp_orders_by_name[normalize_name(order["name"])].append(order)

        kh_orders_by_name = defaultdict(list)
        for order in group_order_rows(kh_records):
            kh_orders_by_name[normalize_name(order["name"])].append(order)

        # 收集所有人名
        all_names = set(member_map.keys()) | set(tp_orders_by_name.keys()) | set(kh_orders_by_name.keys())

        ws_members = get_config_worksheet("section_members")
        member_rows = ws_members.get_all_records(expected_headers=["姓名", "聲部", "手動加分_TP", "手動加分_KH", "身份"])

        discount_list = []

        for name in all_names:
            if not name:
                continue

            info = member_map.get(name, {
                "section": "未分類",
                "identity_code": "5",
                "identity": "暫時未分類",
                "manual_tickets": 0,
                "manual_points": 0.0,
            })

            section = info.get("section") or "未分類"
            identity_code = str(info.get("identity_code") or "5").strip()
            identity_text = normalize_identity(identity_code)

            # 計算台北與高雄訂單基礎劃位積分
            tp_orders = tp_orders_by_name.get(name, [])
            kh_orders = kh_orders_by_name.get(name, [])
            
            base_pts_tp = calc_points_from_orders(tp_orders)
            base_pts_kh = calc_points_from_orders(kh_orders)

            # 讀取手動加分
            manual_pts_tp = 0.0
            manual_pts_kh = 0.0
            
            for m_row in member_rows:
                if normalize_name(m_row.get("姓名")) == name:
                    _, manual_pts_tp = parse_manual_points_string(m_row.get("手動加分_TP"))
                    _, manual_pts_kh = parse_manual_points_string(m_row.get("手動加分_KH"))
                    break

            total_pts_tp = base_pts_tp + manual_pts_tp
            total_pts_kh = base_pts_kh + manual_pts_kh
            grand_total_pts = total_pts_tp + total_pts_kh

            # 計算應有的折抵金額
            discount_amount = calc_discount_amount(grand_total_pts, identity_code)

            discount_list.append({
                "section": section,
                "name": name,
                "identity_code": identity_code,
                "identity": identity_text,
                "tp_order_pts": base_pts_tp,
                "tp_manual_pts": manual_pts_tp,
                "tp_total_pts": total_pts_tp,
                "kh_order_pts": base_pts_kh,
                "kh_manual_pts": manual_pts_kh,
                "kh_total_pts": total_pts_kh,
                "grand_total_pts": grand_total_pts,
                "discount_amount": discount_amount
            })

        # 💡 指定多重排序優先順序：
        # 1. 聲部順序：吹管 -> 彈撥 -> 拉弦 -> 低音 -> 打擊 -> 指揮 -> 未分類
        section_order = {"吹管": 1, "彈撥": 2, "拉弦": 3, "低音": 4, "打擊": 5, "指揮": 6, "未分類": 99, "特殊來源": 999}
        
        # 2. 身份順序：學生協奏(3) -> 團長群(4) -> 團員(1) -> 協演/工人(2) -> 未分類(5)
        identity_order = {"3": 1, "4": 2, "1": 3, "2": 4, "5": 99}

        discount_list.sort(key=lambda x: (
            section_order.get(x["section"], 90),
            identity_order.get(x["identity_code"], 90),
            x["name"]
        ))

        # 生成 CSV 內容
        output = io.StringIO()
        writer = csv.writer(output)
        
        # 💡 寫入匯出時間資訊列
        writer.writerow([f"《嶋聲》團內購票折抵金額清單 - 匯出時間：{time_str_display}"])
        writer.writerow([]) # 空白行隔開

        # 寫入正式欄位表頭
        writer.writerow([
            "聲部", "姓名", "身份類別", 
            "台北場劃位積分", "台北場手動加分", "台北場總積分",
            "高雄場劃位積分", "高雄場手動加分", "高雄場總積分",
            "兩場累積總積分", "可折抵金額(元)"
        ])

        for item in discount_list:
            writer.writerow([
                item["section"],
                item["name"],
                item["identity"],
                item["tp_order_pts"],
                item["tp_manual_pts"],
                item["tp_total_pts"],
                item["kh_order_pts"],
                item["kh_manual_pts"],
                item["kh_total_pts"],
                item["grand_total_pts"],
                item["discount_amount"]
            ])

        # 加上 UTF-8-SIG (BOM) 防止 Excel 打開中文變成亂碼
        csv_data = "\ufeff" + output.getvalue()
        
        response = make_response(csv_data)
        response.headers["Content-Disposition"] = f"attachment; filename=discount_summary_{time_str_file}.csv"
        response.headers["Content-Type"] = "text/csv; charset=utf-8-sig"
        return response

    except Exception as e:
        return jsonify({"success": False, "message": f"匯出折抵清單失敗：{str(e)}"}), 500
    
# ============================================================
# Static File Fallback
# Keep this at the very bottom.
# ============================================================

@app.route("/<path:filename>")
def serve_static_file(filename):
    return send_from_directory(PROJECT_ROOT, filename)

