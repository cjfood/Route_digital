import os
import io
import sys
import json
import sqlite3
import subprocess

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Flask, request, jsonify, send_file, render_template

app = Flask(__name__, template_folder='templates', static_folder='static')
DB_PATH = os.path.abspath('master_dms.db')
DATA_ORDERS_DIR = os.path.abspath('data_don_hang')
os.makedirs(DATA_ORDERS_DIR, exist_ok=True)

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_orders_db():
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id TEXT UNIQUE,
                created_at TEXT,
                date_key TEXT,
                salesman TEXT,
                distributor TEXT,
                customer_type TEXT,
                customer_code TEXT,
                customer_name TEXT,
                customer_address TEXT,
                customer_phone TEXT,
                items_json TEXT,
                items_summary TEXT,
                subtotal REAL,
                discount REAL,
                final_amount REAL,
                payment_method TEXT,
                delivery_date TEXT,
                notes TEXT,
                gps_lat REAL,
                gps_lon REAL,
                gps_distance_meters REAL,
                status TEXT DEFAULT 'COMPLETED'
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Lỗi khởi tạo bảng orders: {e}")

init_orders_db()

import socket

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

@app.route('/')
def index():
    lan_ip = get_local_ip()
    return render_template('index.html', lan_ip=lan_ip)

@app.route('/api/filters')
def get_filters():
    region = request.args.get('region', '').strip()
    area = request.args.get('area', '').strip()
    sup = request.args.get('sup', '').strip()
    distributor = request.args.get('distributor', '').strip()
    salesman = request.args.get('salesman', '').strip()

    conn = get_db()
    cur = conn.cursor()

    # Base WHERE for cascading
    where_clauses = ["1=1"]
    params = []

    # 1. Regions (Luôn lấy toàn bộ)
    cur.execute("SELECT DISTINCT region_name FROM outlets_cj WHERE region_name != '' ORDER BY region_name")
    regions = [r['region_name'] for r in cur.fetchall()]

    # 2. Areas (phụ thuộc Region)
    area_where = ["1=1"]
    area_params = []
    if region:
        area_where.append("region_name = ?")
        area_params.append(region)
    cur.execute(f"SELECT DISTINCT area_name FROM outlets_cj WHERE {' AND '.join(area_where)} AND area_name != '' ORDER BY area_name", area_params)
    areas = [r['area_name'] for r in cur.fetchall()]

    # 3. Supervisors (phụ thuộc Region, Area)
    sup_where = ["1=1"]
    sup_params = []
    if region:
        sup_where.append("region_name = ?")
        sup_params.append(region)
    if area:
        sup_where.append("area_name = ?")
        sup_params.append(area)
    cur.execute(f"SELECT DISTINCT sup_name FROM outlets_cj WHERE {' AND '.join(sup_where)} AND sup_name != '' ORDER BY sup_name", sup_params)
    sups = [r['sup_name'] for r in cur.fetchall()]

    # 4. Distributors (phụ thuộc Region, Area, Sup)
    dist_where = ["1=1"]
    dist_params = []
    if region:
        dist_where.append("region_name = ?")
        dist_params.append(region)
    if area:
        dist_where.append("area_name = ?")
        dist_params.append(area)
    if sup:
        dist_where.append("sup_name = ?")
        dist_params.append(sup)
    cur.execute(f"SELECT DISTINCT distributor_name FROM outlets_cj WHERE {' AND '.join(dist_where)} AND distributor_name != '' ORDER BY distributor_name", dist_params)
    distributors = [r['distributor_name'] for r in cur.fetchall()]

    # 5. Salesmen (phụ thuộc Region, Area, Sup, Dist)
    sm_where = ["1=1"]
    sm_params = []
    if region:
        sm_where.append("region_name = ?")
        sm_params.append(region)
    if area:
        sm_where.append("area_name = ?")
        sm_params.append(area)
    if sup:
        sm_where.append("sup_name = ?")
        sm_params.append(sup)
    if distributor:
        sm_where.append("distributor_name = ?")
        sm_params.append(distributor)
    cur.execute(f"SELECT DISTINCT salesman_name FROM outlets_cj WHERE {' AND '.join(sm_where)} AND salesman_name != '' ORDER BY salesman_name", sm_params)
    salesmen = [r['salesman_name'] for r in cur.fetchall()]

    # 6. Routes (phụ thuộc trên)
    rt_where = list(sm_where)
    rt_params = list(sm_params)
    if salesman:
        rt_where.append("salesman_name = ?")
        rt_params.append(salesman)
    cur.execute(f"SELECT DISTINCT route_code, route_name FROM outlets_cj WHERE {' AND '.join(rt_where)} AND route_code != '' ORDER BY route_code", rt_params)
    routes = [{'code': r['route_code'], 'name': r['route_name']} for r in cur.fetchall()]

    # Frequencies & Weeks
    cur.execute("SELECT DISTINCT frequency FROM outlets_cj WHERE frequency != '' ORDER BY frequency")
    frequencies = [r['frequency'] for r in cur.fetchall()]

    cur.execute("SELECT DISTINCT visit_week FROM outlets_cj WHERE visit_week != '' ORDER BY visit_week")
    visit_weeks = [r['visit_week'] for r in cur.fetchall()]

    conn.close()

    return jsonify({
        'regions': regions,
        'areas': areas,
        'sups': sups,
        'distributors': distributors,
        'salesmen': salesmen,
        'routes': routes,
        'frequencies': frequencies,
        'visit_weeks': visit_weeks
    })

def build_cj_query(args):
    where = ["1=1"]
    params = []

    if args.get('region'):
        where.append("region_name = ?")
        params.append(args['region'])
    if args.get('area'):
        where.append("area_name = ?")
        params.append(args['area'])
    if args.get('sup'):
        where.append("sup_name = ?")
        params.append(args['sup'])
    if args.get('distributor'):
        where.append("distributor_name = ?")
        params.append(args['distributor'])
    if args.get('salesman'):
        where.append("salesman_name = ?")
        params.append(args['salesman'])
    if args.get('route_code'):
        where.append("route_code = ?")
        params.append(args['route_code'])
    if args.get('frequency'):
        where.append("frequency = ?")
        params.append(args['frequency'])
    if args.get('visit_week'):
        where.append("visit_week = ?")
        params.append(args['visit_week'])

    # Day filter: t2, t3, t4, t5, t6, t7, cn
    day = args.get('day', '').lower().strip()
    if day in ['t2', 't3', 't4', 't5', 't6', 't7', 'cn']:
        where.append(f"{day} = 1")

    search = args.get('search', '').strip()
    if search:
        s_term = f"%{search}%"
        where.append("(customer_code LIKE ? OR customer_name LIKE ? OR address LIKE ? OR phone LIKE ?)")
        params.extend([s_term, s_term, s_term, s_term])

    return " AND ".join(where), params

def build_pot_query(args):
    where = ["1=1"]
    params = []

    # BẮT BUỘC CHỈ LẤY KH LOẠI A (Cột X) theo đúng yêu cầu người dùng
    where.append("shop_type = 'A'")

    # Lọc theo Nhà phân phối hoặc NVBH nếu được chọn
    distributor = args.get('distributor', '').strip()
    salesman = args.get('salesman', '').strip()

    if salesman:
        # Match NVBH trong Potential (thường là SNACKEMPNM)
        where.append("(salesman_name LIKE ? OR salesman_name = ?)")
        params.extend([f"%{salesman}%", salesman])
    elif distributor:
        where.append("(distributor_name LIKE ? OR distributor_name = ?)")
        params.extend([f"%{distributor}%", distributor])
    elif args.get('area'):
        # Nếu chỉ chọn Vùng (vd: Hà Nội, Cần Thơ, HCM), lọc theo Tỉnh mới tương ứng
        area = args.get('area', '').strip()
        where.append("province_new LIKE ?")
        params.append(f"%{area}%")

    search = args.get('search', '').strip()
    if search:
        s_term = f"%{search}%"
        where.append("(customer_code LIKE ? OR customer_name LIKE ? OR address LIKE ?)")
        params.extend([s_term, s_term, s_term])

    return " AND ".join(where), params

@app.route('/api/outlets')
def get_outlets():
    args = request.args
    conn = get_db()
    cur = conn.cursor()

    cj_where, cj_params = build_cj_query(args)

    # 1. Thống kê KPI CJ
    cur.execute(f"SELECT COUNT(*), SUM(has_gps) FROM outlets_cj WHERE {cj_where}", cj_params)
    cj_total, cj_gps = cur.fetchone()
    cj_total = cj_total or 0
    cj_gps = cj_gps or 0

    # Thống kê ngày viếng thăm
    cur.execute(f"""
        SELECT 
            SUM(t2) as t2_cnt,
            SUM(t3) as t3_cnt,
            SUM(t4) as t4_cnt,
            SUM(t5) as t5_cnt,
            SUM(t6) as t6_cnt,
            SUM(t7) as t7_cnt,
            SUM(cn) as cn_cnt
        FROM outlets_cj WHERE {cj_where}
    """, cj_params)
    day_row = cur.fetchone()
    day_stats = {
        't2': day_row['t2_cnt'] or 0,
        't3': day_row['t3_cnt'] or 0,
        't4': day_row['t4_cnt'] or 0,
        't5': day_row['t5_cnt'] or 0,
        't6': day_row['t6_cnt'] or 0,
        't7': day_row['t7_cnt'] or 0,
        'cn': day_row['cn_cnt'] or 0,
    }

    # 2. Thống kê KPI Potential Loại A
    include_potential = args.get('include_potential', '1') == '1'
    pot_total = 0
    pot_gps = 0
    pot_points = []
    pot_table = []

    if include_potential:
        pot_where, pot_params = build_pot_query(args)
        cur.execute(f"SELECT COUNT(*), SUM(has_gps) FROM outlets_potential WHERE {pot_where}", pot_params)
        p_row = cur.fetchone()
        pot_total = p_row[0] or 0
        pot_gps = p_row[1] or 0

        # Lấy điểm tiềm năng có GPS cho Map (tối đa 2500 điểm để mượt mà)
        cur.execute(f"""
            SELECT id, customer_code, customer_name, address, latitude, longitude,
                   distributor_name, salesman_name, sellout_month, shop_type
            FROM outlets_potential
            WHERE {pot_where} AND has_gps = 1
            LIMIT 2500
        """, pot_params)
        for r in cur.fetchall():
            pot_points.append({
                'id': r['id'],
                'type': 'pot',
                'code': r['customer_code'],
                'name': r['customer_name'],
                'lat': r['latitude'],
                'lon': r['longitude'],
                'addr': r['address'],
                'dist': r['distributor_name'],
                'sm': r['salesman_name'],
                'shop_type': r['shop_type'],
                'sellout': r['sellout_month']
            })

    # 3. Lấy dữ liệu CJ cho Map (các điểm có GPS)
    cur.execute(f"""
        SELECT id, customer_code, customer_name, address, phone, latitude, longitude,
               frequency, visit_week, visit_days, salesman_name, distributor_name,
               route_code, route_name, channel_name, business_type_name,
               t2, t3, t4, t5, t6, t7, cn
        FROM outlets_cj
        WHERE {cj_where} AND has_gps = 1
        LIMIT 5000
    """, cj_params)
    cj_points = []
    for r in cur.fetchall():
        cj_points.append({
            'id': r['id'],
            'type': 'cj',
            'code': r['customer_code'],
            'name': r['customer_name'],
            'lat': r['latitude'],
            'lon': r['longitude'],
            'addr': r['address'],
            'phone': r['phone'],
            'freq': r['frequency'],
            'week': r['visit_week'],
            'days': r['visit_days'],
            'sm': r['salesman_name'],
            'dist': r['distributor_name'],
            'route': r['route_name'] or r['route_code'],
            'channel': r['channel_name'],
            'btype': r['business_type_name'],
            't2': r['t2'], 't3': r['t3'], 't4': r['t4'],
            't5': r['t5'], 't6': r['t6'], 't7': r['t7'], 'cn': r['cn']
        })

    # 4. Lấy dữ liệu hiển thị Bảng (Table Data - giới hạn 1000 dòng đầu tiên)
    cur.execute(f"""
        SELECT id, 'CJ' as source, customer_code, customer_name, address, phone,
               frequency, visit_week, visit_days, salesman_name, distributor_name,
               route_name, channel_name, latitude, longitude, has_gps
        FROM outlets_cj
        WHERE {cj_where}
        ORDER BY salesman_name, route_name, customer_code
        LIMIT 1000
    """, cj_params)
    table_rows = [dict(r) for r in cur.fetchall()]

    # Thêm mẫu KH tiềm năng Loại A vào bảng nếu có lọc cụ thể
    if include_potential and (args.get('salesman') or args.get('distributor') or args.get('area')):
        cur.execute(f"""
            SELECT id, 'POTENTIAL_A' as source, customer_code, customer_name, address, '' as phone,
                   'Shop Loại A' as frequency, '' as visit_week, '' as visit_days,
                   salesman_name, distributor_name, '' as route_name,
                   'Tiềm năng (Sellout: ' || ROUND(sellout_month, 0) || 'đ)' as channel_name,
                   latitude, longitude, has_gps
            FROM outlets_potential
            WHERE {pot_where}
            LIMIT 200
        """, pot_params)
        table_rows.extend([dict(r) for r in cur.fetchall()])

    conn.close()

    return jsonify({
        'kpis': {
            'total_all': cj_total + pot_total,
            'cj_total': cj_total,
            'cj_gps': cj_gps,
            'cj_gps_pct': round((cj_gps / cj_total * 100), 1) if cj_total > 0 else 0,
            'pot_total': pot_total,
            'pot_gps': pot_gps,
            'pot_gps_pct': round((pot_gps / pot_total * 100), 1) if pot_total > 0 else 0,
            'day_stats': day_stats
        },
        'map_points': cj_points + pot_points,
        'table_data': table_rows
    })

@app.route('/api/export')
def export_excel():
    args = request.args
    conn = get_db()
    cur = conn.cursor()

    cj_where, cj_params = build_cj_query(args)

    cur.execute(f"""
        SELECT 
            region_name, area_name, sup_name, distributor_name,
            salesman_code, salesman_name, route_code, route_name,
            customer_code, customer_name, address, phone,
            channel_name, business_type_name,
            frequency, visit_week, visit_days,
            latitude, longitude,
            CASE WHEN has_gps=1 THEN 'Có GPS' ELSE 'Chưa có GPS' END as gps_status
        FROM outlets_cj
        WHERE {cj_where}
        ORDER BY salesman_name, route_code, customer_code
    """, cj_params)
    rows = cur.fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DMS_CJ_Routes"

    # Header Styling
    header_fill = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center")
    border_thin = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    headers = [
        "Miền", "Vùng", "GSBH", "Nhà Phân Phối",
        "Mã NVBH", "Tên NVBH", "Mã Tuyến", "Tên Tuyến",
        "Mã KH", "Tên Khách Hàng", "Địa Chỉ", "Số Điện Thoại",
        "Kênh BH", "Loại Hình",
        "Tần Suất", "Tuần Viếng Thăm", "Ngày Viếng Thăm",
        "Vĩ Độ (Lat)", "Kinh Độ (Lon)", "Trạng Thái GPS"
    ]

    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    for r in rows:
        ws.append(list(r))

    # Auto fit column width
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)

    # Thêm sheet 2: Potential Loại A nếu có chọn NVBH hoặc NPP
    include_pot = args.get('include_potential', '1') == '1'
    if include_pot:
        pot_where, pot_params = build_pot_query(args)
        cur.execute(f"""
            SELECT 
                province_new, distributor_name, salesman_name,
                customer_code, customer_name, address,
                shop_type, sellout_month, latitude, longitude
            FROM outlets_potential
            WHERE {pot_where}
            LIMIT 5000
        """, pot_params)
        pot_rows = cur.fetchall()

        if pot_rows:
            ws_pot = wb.create_sheet(title="Potential_Loai_A")
            pot_headers = [
                "Tỉnh/Khu vực", "Nhà Phân Phối", "NVBH Phụ Trách",
                "Mã KH Tiềm Năng", "Tên Khách Hàng", "Địa Chỉ",
                "Phân Loại Shop", "Doanh Số Sellout/Tháng (ước tính)",
                "Vĩ Độ (Lat)", "Kinh Độ (Lon)"
            ]
            ws_pot.append(pot_headers)
            pot_fill = PatternFill(start_color="B9770E", end_color="B9770E", fill_type="solid")
            for col_num in range(1, len(pot_headers) + 1):
                cell = ws_pot.cell(row=1, column=col_num)
                cell.fill = pot_fill
                cell.font = header_font
                cell.alignment = align_center

            for pr in pot_rows:
                ws_pot.append(list(pr))

            for col in ws_pot.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                ws_pot.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)

    conn.close()

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = "CJ_Tuyen_Ban_Hang_DMS.xlsx"
    if args.get('salesman'):
        filename = f"Tuyen_DMS_{args['salesman'].replace(' ', '_')}.xlsx"

    return send_file(
        output,
        download_name=filename,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/api/sync', methods=['POST'])
def sync_database():
    try:
        # Chạy script etl_master.py để làm mới DB
        res = subprocess.run(['python', 'etl_master.py'], capture_output=True, text=True, check=True)
        return jsonify({
            'success': True,
            'message': 'Đã đồng bộ và cập nhật Master Database thành công!',
            'output': res.stdout
        })
    except subprocess.CalledProcessError as e:
        return jsonify({
            'success': False,
            'message': f'Lỗi khi đồng bộ: {e.stderr}'
        }), 500

def append_order_to_excel(excel_path, row_data):
    if not os.path.exists(excel_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Don_Hang_Tong_Hop"
        headers = [
            "Mã Đơn Hàng", "Thời Gian", "NVBH", "Nhà Phân Phối", "Loại Điểm",
            "Mã KH", "Tên Khách Hàng", "Địa Chỉ", "Số ĐT", "Chi Tiết Sản Phẩm Đặt",
            "Tiền Hàng (VNĐ)", "Chiết Khấu (VNĐ)", "Thực Thu (VNĐ)",
            "Hình Thức TT", "Ngày Hẹn Giao", "Ghi Chú", "GPS Check-in"
        ]
        ws.append(headers)
        
        header_fill = PatternFill(start_color="0A2E50", end_color="0A2E50", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style='thin', color='DDDDDD'),
            right=Side(style='thin', color='DDDDDD'),
            top=Side(style='thin', color='DDDDDD'),
            bottom=Side(style='thin', color='DDDDDD')
        )
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
        ws.row_dimensions[1].height = 28
    else:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active

    ws.append(row_data)
    new_row = ws.max_row
    data_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style='thin', color='EEEEEE'),
        right=Side(style='thin', color='EEEEEE'),
        top=Side(style='thin', color='EEEEEE'),
        bottom=Side(style='thin', color='EEEEEE')
    )
    for col_idx in range(1, len(row_data) + 1):
        cell = ws.cell(row=new_row, column=col_idx)
        cell.font = data_font
        cell.border = thin_border
        if col_idx in [11, 12, 13]:
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif col_idx in [1, 2, 5, 6, 9, 14, 15, 17]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(c.value or '')) for c in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 60)

    wb.save(excel_path)

@app.route('/api/orders/submit', methods=['POST'])
def submit_order():
    try:
        data = request.get_json(force=True) or {}
        now = datetime.now()
        
        # Order ID & Date key
        order_id = data.get('orderId') or f"PO-{now.strftime('%Y%m%d-%H%M%S')}"
        created_at = data.get('createdAt') or now.strftime('%Y-%m-%d %H:%M:%S')
        date_key = now.strftime('%Y-%m-%d')

        salesman = (data.get('salesman') or '').strip()
        distributor = (data.get('distributor') or '').strip()
        cust = data.get('customer') or {}
        cust_type = (data.get('customerType') or cust.get('type') or 'CJ').upper()
        cust_code = (cust.get('code') or '').strip()
        cust_name = (cust.get('name') or '').strip()
        cust_addr = (cust.get('address') or '').strip()
        cust_phone = (cust.get('phone') or '').strip()

        items = data.get('items') or []
        summary_parts = []
        for it in items:
            name = it.get('name', '')
            qty = it.get('qty', 0)
            unit = it.get('unit', 'gói')
            summary_parts.append(f"{name} (x{qty} {unit})")
        items_summary = "; ".join(summary_parts) if summary_parts else "Không có món"

        subtotal = float(data.get('subtotal') or 0)
        discount = float(data.get('discount') or 0)
        final_amount = float(data.get('finalAmount') or (subtotal - discount))
        payment_method = data.get('paymentMethod') or 'Tiền mặt (COD)'
        delivery_date = data.get('deliveryDate') or now.strftime('%Y-%m-%d')
        notes = (data.get('notes') or '').strip()

        gps = data.get('gpsCheckin') or {}
        gps_lat = float(gps.get('lat') or 0)
        gps_lon = float(gps.get('lon') or 0)
        gps_dist = float(gps.get('distanceMeters') or 0)
        gps_status = f"✓ {gps_dist:.0f}m" if gps.get('inRadius', True) else f"⚠ Cách {gps_dist:.0f}m"

        # 1. Save individual JSON in data_don_hang/{date_key}/
        daily_dir = os.path.join(DATA_ORDERS_DIR, date_key)
        os.makedirs(daily_dir, exist_ok=True)
        
        safe_code = cust_code.replace('/', '_').replace('\\', '_')
        json_filename = f"{order_id}_{safe_code}.json"
        json_path = os.path.join(daily_dir, json_filename)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # Save photos if any
        photos = data.get('photos') or []
        if photos:
            photos_dir = os.path.join(daily_dir, 'photos')
            os.makedirs(photos_dir, exist_ok=True)
            import base64
            for idx, p_b64 in enumerate(photos):
                try:
                    if ',' in p_b64:
                        p_b64 = p_b64.split(',', 1)[1]
                    img_data = base64.b64decode(p_b64)
                    photo_filename = f"{order_id}_photo_{idx+1}.jpg"
                    with open(os.path.join(photos_dir, photo_filename), 'wb') as pf:
                        pf.write(img_data)
                except Exception as pe:
                    print(f"Lỗi lưu ảnh chụp: {pe}")

        # 2. Append to daily summary Excel
        excel_path = os.path.join(daily_dir, f"Tong_Hop_Don_Hang_{date_key}.xlsx")
        row_data = [
            order_id, created_at, salesman, distributor, cust_type,
            cust_code, cust_name, cust_addr, cust_phone, items_summary,
            subtotal, discount, final_amount,
            payment_method, delivery_date, notes, gps_status
        ]
        append_order_to_excel(excel_path, row_data)

        # 3. Save to master_dms.db table orders
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            INSERT OR REPLACE INTO orders (
                order_id, created_at, date_key, salesman, distributor,
                customer_type, customer_code, customer_name, customer_address, customer_phone,
                items_json, items_summary, subtotal, discount, final_amount,
                payment_method, delivery_date, notes, gps_lat, gps_lon, gps_distance_meters
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            order_id, created_at, date_key, salesman, distributor,
            cust_type, cust_code, cust_name, cust_addr, cust_phone,
            json.dumps(items, ensure_ascii=False), items_summary, subtotal, discount, final_amount,
            payment_method, delivery_date, notes, gps_lat, gps_lon, gps_dist
        ))
        conn.commit()
        conn.close()

        # 4. Log to sync_log.txt
        log_path = os.path.join(DATA_ORDERS_DIR, 'sync_log.txt')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] ĐÃ NHẬN: {order_id} | NVBH: {salesman} | KH: {cust_name} ({cust_code}) | {final_amount:,.0f} đ\n")

        return jsonify({
            'success': True,
            'orderId': order_id,
            'message': 'Đã lưu đơn hàng thành công vào thư mục data_don_hang!',
            'excelFile': f"Tong_Hop_Don_Hang_{date_key}.xlsx",
            'jsonFile': json_filename
        })
    except Exception as e:
        return jsonify({'success': False, 'message': f'Lỗi khi lưu đơn hàng: {str(e)}'}), 500

@app.route('/api/orders/list', methods=['GET'])
def list_orders():
    try:
        date_key = request.args.get('date', datetime.now().strftime('%Y-%m-%d'))
        salesman = request.args.get('salesman', '').strip()
        conn = get_db()
        cur = conn.cursor()
        if salesman:
            cur.execute("SELECT * FROM orders WHERE date_key = ? AND salesman = ? ORDER BY id DESC", (date_key, salesman))
        else:
            cur.execute("SELECT * FROM orders WHERE date_key = ? ORDER BY id DESC", (date_key,))
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return jsonify({'success': True, 'orders': rows, 'total': len(rows)})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/orders/export-excel', methods=['GET'])
def export_orders_excel():
    try:
        date_key = request.args.get('date', datetime.now().strftime('%Y-%m-%d'))
        daily_dir = os.path.join(DATA_ORDERS_DIR, date_key)
        excel_path = os.path.join(daily_dir, f"Tong_Hop_Don_Hang_{date_key}.xlsx")
        if not os.path.exists(excel_path):
            return jsonify({'success': False, 'message': f'Chưa có file đơn hàng cho ngày {date_key}'}), 404
        return send_file(excel_path, as_attachment=True, download_name=f"Tong_Hop_Don_Hang_{date_key}.xlsx")
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    print("Khởi động CJ DMS Route Map Server tại http://127.0.0.1:5000 ...")
    app.run(host='0.0.0.0', port=5000, debug=False)
