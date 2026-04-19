import json
import os
import sys
from http.server import BaseHTTPRequestHandler

sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from lib.seat_parser import parse_seat_map
from lib.sheet_repo import build_active_sold_seat_keys

SEAT_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "seat_map.xlsx")
SECOND_FLOOR_START_ROW = 33


def get_floor_label_from_excel_row(excel_row: int) -> str:
    return "2樓" if excel_row >= SECOND_FLOOR_START_ROW else "1樓"


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            seats, row_labels, _ = parse_seat_map(SEAT_FILE)
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

            body = json.dumps({
                "seats": result_seats,
                "row_labels": row_labels
            }, ensure_ascii=False).encode("utf-8")

            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        except Exception as e:
            body = json.dumps({
                "success": False,
                "message": str(e)
            }, ensure_ascii=False).encode("utf-8")

            self.send_response(500)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)