import os
import time
from threading import Lock

from flask import Flask, jsonify, request, session
from functools import wraps

from lib.seat_parser import parse_seat_map
from lib.sheet_repo import (
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
    admin_toggle_ticket_adjusted_status,
    admin_delete_order,
    build_stats_summary,
    get_all_records,
    normalize_text,
    get_section_members_rows,
    get_stats_config_rows,
    save_section_members_rows,
    save_stats_config_rows,
)

app = Flask(__name__)

app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")
PROJECT_ROOT = os.path.dirname(os.path.dirname(__file__))
SEAT_FILE = os.path.join(PROJECT_ROOT, "data", "seat_map.xlsx")

SEAT_CACHE = {
    "seats": None,
    "row_labels": None,
    "loaded_at": 0,
}
SEAT_CACHE_TTL = 60
SECOND_FLOOR_START_ROW = 33
confirm_lock = Lock()

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


@app.route("/api/seats", methods=["GET"])
def api_seats():
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

    return jsonify({
        "seats": result_seats,
        "row_labels": row_labels
    })


@app.route("/api/confirm", methods=["POST"])
def api_confirm():
    with confirm_lock:
        data = request.get_json(silent=True) or {}
        name = str(data.get("name", "")).strip()
        selected_seat_ids = data.get("seats", [])

        if not name:
            return jsonify({"success": False, "message": "請輸入姓名"}), 400

        if not selected_seat_ids:
            return jsonify({"success": False, "message": "請選擇座位"}), 400

        seats, _ = get_cached_seat_map()
        seat_map = {
            f"{seat['excel_row']}-{seat['excel_col']}": seat
            for seat in seats
        }

        active_sold_keys = build_active_sold_seat_keys()
        seat_rows_to_save = []

        for seat_id in selected_seat_ids:
            seat = seat_map.get(seat_id)

            if not seat:
                return jsonify({"success": False, "message": f"找不到座位 {seat_id}"}), 400

            floor = get_floor_label_from_excel_row(seat["excel_row"])
            seat_key = (floor, str(seat["row_label"]), int(seat["seat_number"]))

            if seat_key in active_sold_keys:
                return jsonify({
                    "success": False,
                    "message": f"{floor}{seat['row_label']}排{seat['seat_number']}號 已被選走"
                }), 400

            if not seat["available"]:
                return jsonify({
                    "success": False,
                    "message": f"{floor}{seat['row_label']}排{seat['seat_number']}號 不開放購買"
                }), 400

            seat_rows_to_save.append({
                "floor": floor,
                "row_label": str(seat["row_label"]),
                "seat_number": int(seat["seat_number"]),
                "price": int(seat["price"]),
            })

        order_id = append_order_rows(name=name, seat_rows=seat_rows_to_save)

        return jsonify({
            "success": True,
            "message": f"訂位成功！訂單編號：{order_id}",
            "order_id": order_id,
        })


@app.route("/api/orders", methods=["GET"])
def api_orders():
    name = request.args.get("name", "").strip()

    if not name:
        return jsonify({
            "success": False,
            "message": "請輸入姓名",
            "orders": []
        }), 400

    return jsonify({
        "success": True,
        "orders": get_orders_by_name(name)
    })

@app.route("/api/orders/<order_id>/note", methods=["PATCH"])
def api_update_order_note(order_id):
    data = request.get_json(silent=True) or {}
    note = str(data.get("note", "")).strip()

    from lib.sheet_repo import get_all_records, normalize_text

    rows = get_all_records()
    debug_ids = [normalize_text(row.get("訂單ID")) for row in rows[:20]]

    print("DEBUG note order_id =", repr(order_id))
    print("DEBUG first 20 sheet order_ids =", debug_ids)

    ok = update_order_note(order_id, note)
    if not ok:
        return jsonify({
            "success": False,
            "message": "找不到訂單",
            "debug_order_id": order_id,
            "debug_first_ids": debug_ids
        }), 404

    return jsonify({"success": True, "message": "備註已更新"})

@app.route("/api/orders/<order_id>", methods=["DELETE"])
def api_delete_order(order_id):
    rows = get_all_records()
    locked = any(
        normalize_text(row.get("訂單ID")) == order_id
        and normalize_text(row.get("訂單狀態")).lower() == "locked"
        for row in rows
    )

    if locked:
        return jsonify({"success": False, "message": "已鎖定，無法刪除"}), 403

    ok = mark_order_deleted(order_id)
    if not ok:
        return jsonify({"success": False, "message": "找不到訂單"}), 404

    return jsonify({"success": True, "message": "訂單已刪除，座位已重新釋出"})


@app.route("/api/orders/<order_id>/pickup", methods=["PATCH"])
def api_update_order_pickup(order_id):
    data = request.get_json(silent=True) or {}

    ok = update_order_pickup_status(
        order_id,
        pickup_open=data.get("pickup_open"),
        picked_up=data.get("picked_up")
    )

    if not ok:
        return jsonify({"success": False, "message": "找不到訂單"}), 404

    return jsonify({"success": True, "message": "取票狀態已更新"})

@app.route("/api/admin/login", methods=["POST"])
def api_admin_login():
    try:
        data = request.get_json(silent=True) or {}
        password = str(data.get("password", ""))

        admin_password = os.environ.get("ADMIN_PASSWORD")

        if not admin_password:
            return jsonify({
                "success": False,
                "message": "後台密碼尚未設定"
            }), 500

        if password == admin_password:
            session["admin_ok"] = True
            return jsonify({"success": True})

        return jsonify({
            "success": False,
            "message": "你不是票務 :("
        }), 401

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e)
        }), 500

def require_admin(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("admin_ok"):
            return jsonify({"success": False, "message": "你不是票務！"}), 401
        return fn(*args, **kwargs)
    return wrapper

@app.route("/api/admin/orders", methods=["GET"])
@require_admin
def api_admin_orders():
    keyword = request.args.get("keyword", "").strip()

    if keyword:
        orders = admin_search_orders(keyword)
    else:
        orders = admin_search_orders("")  # 空字串代表全部訂單

    return jsonify({
        "success": True,
        "orders": orders
    })

@app.route("/api/admin/orders/<order_id>/ticket-adjusted", methods=["PATCH"])
@require_admin
def api_admin_ticket_adjusted(order_id):
    ok, message = admin_toggle_ticket_adjusted_status(order_id)
    if not ok:
        return jsonify({"success": False, "message": message}), 404
    return jsonify({"success": True, "message": message})
    
@app.route("/api/admin/orders/<order_id>/lock", methods=["PATCH"])
@require_admin
def api_admin_lock(order_id):
    ok, message = admin_toggle_lock_status(order_id)
    if not ok:
        return jsonify({"success": False, "message": message}), 404
    return jsonify({"success": True, "message": message})


@app.route("/api/admin/orders/<order_id>/payment", methods=["PATCH"])
@require_admin
def api_admin_payment(order_id):
    ok, message = admin_toggle_payment_status(order_id)
    if not ok:
        return jsonify({"success": False, "message": message}), 404
    return jsonify({"success": True, "message": message})


@app.route("/api/admin/orders/<order_id>/pickup/advance", methods=["PATCH"])
@require_admin
def api_admin_pickup_advance(order_id):
    ok, message = admin_advance_pickup_status(order_id)
    if not ok:
        return jsonify({"success": False, "message": message}), 404
    return jsonify({"success": True, "message": message})


@app.route("/api/admin/orders/<order_id>", methods=["DELETE"])
@require_admin
def api_admin_delete(order_id):
    ok, message = admin_delete_order(order_id)
    if not ok:
        return jsonify({"success": False, "message": message}), 403
    return jsonify({"success": True, "message": message})

@app.route("/api/edit/config", methods=["GET"])
@require_admin
def api_edit_get_config():
    try:
        section_members = worksheet_to_csv_text("section_members")
        stats_config = worksheet_to_csv_text("stats_config")

        return jsonify({
            "success": True,
            "section_members": section_members,
            "stats_config": stats_config
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e)
        }), 500


@app.route("/api/edit/section-members", methods=["PUT"])
@require_admin
def api_edit_section_members():
    try:
        data = request.get_json(silent=True) or {}
        text = str(data.get("text", ""))

        replace_worksheet_from_csv_text(
            "section_members",
            ["姓名", "聲部"],
            text
        )

        return jsonify({
            "success": True,
            "message": "聲部名單已更新"
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e)
        }), 500


@app.route("/api/edit/stats-config", methods=["PUT"])
@require_admin
def api_edit_stats_config():
    try:
        data = request.get_json(silent=True) or {}
        text = str(data.get("text", ""))

        replace_worksheet_from_csv_text(
            "stats_config",
            ["類型", "名稱", "條件"],
            text
        )

        return jsonify({
            "success": True,
            "message": "統計設定已更新"
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "message": str(e)
        }), 500

@app.route("/api/stats", methods=["GET"])
def api_stats():
    return jsonify({
        "success": True,
        "data": build_stats_summary()
    })


