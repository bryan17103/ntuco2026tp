# NTUCO ticket system - Vercel serverless version

## 結構
- `index.html`, `orders.html`, `stats.html`, `admin.html`: 靜態頁面
- `app.js`, `style.css`: 前端資源
- `api/index.py`: Vercel Python WSGI serverless API 路由入口
- `lib/seat_parser.py`, `lib/sheet_repo.py`: 後端邏輯
- `data/`: 需自行放入 `seat_map.xlsx`, `section_members.txt`, `stats_config.txt`

## Vercel 環境變數
必填：
- `GOOGLE_CREDENTIALS_JSON`

## 部署前確認
1. `data/seat_map.xlsx`
2. `data/section_members.txt`
3. `data/stats_config.txt`
4. Vercel project settings 已設定 `GOOGLE_CREDENTIALS_JSON`
