import os
import sys
import time
import glob
import sqlite3
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

DB_PATH = 'master_dms.db'

def create_database():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    
    # 1. Outlets CJ Table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS outlets_cj (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        country_code TEXT,
        country_name TEXT,
        region_code TEXT,
        region_name TEXT,
        area_code TEXT,
        area_name TEXT,
        province_code TEXT,
        distributor_province TEXT,
        distributor_code TEXT,
        distributor_name TEXT,
        route_type TEXT,
        route_code TEXT,
        route_name TEXT,
        rsm_asm_code TEXT,
        asm_name TEXT,
        sup_code TEXT,
        sup_name TEXT,
        salesman_code TEXT,
        salesman_name TEXT,
        customer_code TEXT,
        customer_old_code TEXT,
        customer_name TEXT,
        customer_group TEXT,
        address TEXT,
        street_number TEXT,
        street_name TEXT,
        ward TEXT,
        province TEXT,
        phone TEXT,
        channel_code TEXT,
        channel_name TEXT,
        business_type_code TEXT,
        business_type_name TEXT,
        created_by TEXT,
        created_date TEXT,
        mcp_created_date TEXT,
        contact_person TEXT,
        status TEXT,
        mcp_status TEXT,
        longitude REAL,
        latitude REAL,
        frequency TEXT,
        visit_week TEXT,
        t2 INTEGER,
        t3 INTEGER,
        t4 INTEGER,
        t5 INTEGER,
        t6 INTEGER,
        t7 INTEGER,
        cn INTEGER,
        visit_days TEXT,
        has_gps INTEGER
    )
    ''')

    # 2. Outlets Potential Table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS outlets_potential (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        province_new TEXT,
        customer_code TEXT,
        customer_name TEXT,
        status TEXT,
        grade_area TEXT,
        grade_snack TEXT,
        grade_gum TEXT,
        address TEXT,
        province_old TEXT,
        district TEXT,
        ward TEXT,
        street TEXT,
        house_no TEXT,
        business_type TEXT,
        distributor_code TEXT,
        distributor_name TEXT,
        salesman_code TEXT,
        salesman_name TEXT,
        pie_wkday TEXT,
        longitude REAL,
        latitude REAL,
        sellout_3m REAL,
        sellout_month REAL,
        shop_type TEXT,
        has_gps INTEGER
    )
    ''')

    # 3. Create indexes for fast querying & filtering
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_region ON outlets_cj(region_name)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_area ON outlets_cj(area_name)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_sup ON outlets_cj(sup_name)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_dist ON outlets_cj(distributor_name)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_sm ON outlets_cj(salesman_name)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_route ON outlets_cj(route_code)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_freq ON outlets_cj(frequency)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_week ON outlets_cj(visit_week)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_t2 ON outlets_cj(t2)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_t3 ON outlets_cj(t3)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_t4 ON outlets_cj(t4)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_t5 ON outlets_cj(t5)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_t6 ON outlets_cj(t6)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_t7 ON outlets_cj(t7)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_cn ON outlets_cj(cn)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_cj_gps ON outlets_cj(has_gps)')
    
    cur.execute('CREATE INDEX IF NOT EXISTS idx_pot_type ON outlets_potential(shop_type)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_pot_dist ON outlets_potential(distributor_name)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_pot_sm ON outlets_potential(salesman_name)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_pot_gps ON outlets_potential(has_gps)')
    
    conn.commit()
    conn.close()

def clean_val(v):
    if v is None:
        return ""
    if isinstance(v, float) and (v != v): # NaN check
        return ""
    return str(v).strip()

def parse_float(v):
    if v is None:
        return 0.0
    try:
        f = float(v)
        if f != f: # NaN
            return 0.0
        return f
    except (ValueError, TypeError):
        return 0.0

def is_valid_gps(lat, lon):
    # Vietnam latitude: ~8.0 to ~24.0, longitude: ~102.0 to ~110.0
    if lat is None or lon is None:
        return 0
    if lat == 0.0 or lon == 0.0:
        return 0
    if 7.0 <= lat <= 25.0 and 101.0 <= lon <= 112.0:
        return 1
    return 0

def import_cj_data(excel):
    files = [f for f in glob.glob('Shop CJ/*.xls*') if not os.path.basename(f).startswith('~$')]
    if not files:
        print("Không tìm thấy file trong thư mục Shop CJ")
        return
    fpath = os.path.abspath(files[0])
    print(f"[1/2] Đang xử lý Shop CJ: {os.path.basename(fpath)}...")
    t0 = time.time()
    
    wb = excel.Workbooks.Open(fpath, ReadOnly=True)
    ws = wb.Sheets(1)
    total_rows = ws.UsedRange.Rows.Count
    print(f"Tổng số dòng trong sheet: {total_rows}")
    
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute('DELETE FROM outlets_cj') # Làm mới dữ liệu CJ
    
    chunk_size = 10000
    start_row = 6 # Dữ liệu bắt đầu từ dòng 6 (dòng 5 là header)
    inserted = 0
    
    while start_row <= total_rows:
        end_row = min(start_row + chunk_size - 1, total_rows)
        raw_chunk = ws.Range(ws.Cells(start_row, 1), ws.Cells(end_row, 52)).Value
        rows_to_insert = []
        
        for r in raw_chunk:
            cust_code = clean_val(r[21])
            if not cust_code: # Không có mã KH
                continue
            
            lon = parse_float(r[41])
            lat = parse_float(r[42])
            has_gps = is_valid_gps(lat, lon)
            
            freq = clean_val(r[43]).upper()
            week = clean_val(r[44]).upper()
            
            # Days: Col 46 (T2) -> Col 52 (CN)
            t2 = 1 if clean_val(r[45]).upper() == 'X' else 0
            t3 = 1 if clean_val(r[46]).upper() == 'X' else 0
            t4 = 1 if clean_val(r[47]).upper() == 'X' else 0
            t5 = 1 if clean_val(r[48]).upper() == 'X' else 0
            t6 = 1 if clean_val(r[49]).upper() == 'X' else 0
            t7 = 1 if clean_val(r[50]).upper() == 'X' else 0
            cn = 1 if clean_val(r[51]).upper() == 'X' else 0
            
            days = []
            if t2: days.append("T2")
            if t3: days.append("T3")
            if t4: days.append("T4")
            if t5: days.append("T5")
            if t6: days.append("T6")
            if t7: days.append("T7")
            if cn: days.append("CN")
            visit_days = ", ".join(days)
            
            row_data = (
                clean_val(r[0]),  # country_code
                clean_val(r[1]),  # country_name
                clean_val(r[2]),  # region_code
                clean_val(r[3]),  # region_name
                clean_val(r[4]),  # area_code
                clean_val(r[5]),  # area_name
                clean_val(r[6]),  # province_code
                clean_val(r[7]),  # distributor_province
                clean_val(r[10]), # distributor_code
                clean_val(r[11]), # distributor_name
                clean_val(r[12]), # route_type
                clean_val(r[13]), # route_code
                clean_val(r[14]), # route_name
                clean_val(r[15]), # rsm_asm_code
                clean_val(r[16]), # asm_name
                clean_val(r[17]), # sup_code
                clean_val(r[18]), # sup_name
                clean_val(r[19]), # salesman_code
                clean_val(r[20]), # salesman_name
                cust_code,        # customer_code
                clean_val(r[22]), # customer_old_code
                clean_val(r[23]), # customer_name
                clean_val(r[24]), # customer_group
                clean_val(r[25]), # address
                clean_val(r[26]), # street_number
                clean_val(r[27]), # street_name
                clean_val(r[28]), # ward
                clean_val(r[29]), # province
                clean_val(r[30]), # phone
                clean_val(r[31]), # channel_code
                clean_val(r[32]), # channel_name
                clean_val(r[33]), # business_type_code
                clean_val(r[34]), # business_type_name
                clean_val(r[35]), # created_by
                clean_val(r[36]), # created_date
                clean_val(r[37]), # mcp_created_date
                clean_val(r[38]), # contact_person
                clean_val(r[39]), # status
                clean_val(r[40]), # mcp_status
                lon,
                lat,
                freq,
                week,
                t2, t3, t4, t5, t6, t7, cn,
                visit_days,
                has_gps
            )
            rows_to_insert.append(row_data)
        
        cur.executemany('''
        INSERT INTO outlets_cj (
            country_code, country_name, region_code, region_name, area_code, area_name,
            province_code, distributor_province, distributor_code, distributor_name,
            route_type, route_code, route_name, rsm_asm_code, asm_name, sup_code, sup_name,
            salesman_code, salesman_name, customer_code, customer_old_code, customer_name,
            customer_group, address, street_number, street_name, ward, province, phone,
            channel_code, channel_name, business_type_code, business_type_name, created_by,
            created_date, mcp_created_date, contact_person, status, mcp_status,
            longitude, latitude, frequency, visit_week,
            t2, t3, t4, t5, t6, t7, cn, visit_days, has_gps
        ) VALUES (''' + ','.join(['?']*52) + ')', rows_to_insert)
        conn.commit()
        
        inserted += len(rows_to_insert)
        print(f"  -> Đã nạp {inserted}/{total_rows - 5} dòng CJ ({inserted*100//(total_rows-5)}%)...")
        start_row = end_row + 1
        
    wb.Close(False)
    conn.close()
    print(f"Hoàn thành nạp Shop CJ: {inserted} điểm bán trong {time.time()-t0:.2f}s")

def import_potential_data(excel):
    files = [f for f in glob.glob('Shop Potential/*.xls*') if not os.path.basename(f).startswith('~$')]
    if not files:
        print("Không tìm thấy file trong thư mục Shop Potential")
        return
    fpath = os.path.abspath(files[0])
    print(f"\n[2/2] Đang xử lý Shop Potential: {os.path.basename(fpath)}...")
    t0 = time.time()
    
    wb = excel.Workbooks.Open(fpath, ReadOnly=True)
    ws = wb.Sheets(1)
    total_rows = ws.UsedRange.Rows.Count
    print(f"Tổng số dòng trong sheet: {total_rows}")
    
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute('DELETE FROM outlets_potential') # Làm mới dữ liệu Potential
    
    chunk_size = 20000
    start_row = 2 # Dòng 1 là header
    inserted = 0
    type_a_count = 0
    
    while start_row <= total_rows:
        end_row = min(start_row + chunk_size - 1, total_rows)
        raw_chunk = ws.Range(ws.Cells(start_row, 1), ws.Cells(end_row, 24)).Value
        rows_to_insert = []
        
        for r in raw_chunk:
            raw_code = clean_val(r[1])
            if not raw_code:
                continue
            cust_code = f"GS_{raw_code}5"
            
            lon = parse_float(r[19])
            lat = parse_float(r[20])
            has_gps = is_valid_gps(lat, lon)
            
            shop_type = clean_val(r[23]).upper() # Cột 24 (Cột X)
            if shop_type == 'A':
                type_a_count += 1
                
            row_data = (
                clean_val(r[0]),  # province_new
                cust_code,        # customer_code
                clean_val(r[2]),  # customer_name
                clean_val(r[3]),  # status
                clean_val(r[4]),  # grade_area
                clean_val(r[5]),  # grade_snack
                clean_val(r[6]),  # grade_gum
                clean_val(r[7]),  # address
                clean_val(r[8]),  # province_old
                clean_val(r[9]),  # district
                clean_val(r[10]), # ward
                clean_val(r[11]), # street
                clean_val(r[12]), # house_no
                clean_val(r[13]), # business_type
                clean_val(r[14]), # distributor_code
                clean_val(r[15]), # distributor_name
                clean_val(r[16]), # salesman_code
                clean_val(r[17]), # salesman_name
                clean_val(r[18]), # pie_wkday
                lon,
                lat,
                parse_float(r[21]), # sellout_3m
                parse_float(r[22]), # sellout_month
                shop_type,          # shop_type (Cột X)
                has_gps
            )
            rows_to_insert.append(row_data)
            
        cur.executemany('''
        INSERT INTO outlets_potential (
            province_new, customer_code, customer_name, status,
            grade_area, grade_snack, grade_gum, address, province_old,
            district, ward, street, house_no, business_type,
            distributor_code, distributor_name, salesman_code, salesman_name,
            pie_wkday, longitude, latitude, sellout_3m, sellout_month,
            shop_type, has_gps
        ) VALUES (''' + ','.join(['?']*25) + ')', rows_to_insert)
        conn.commit()
        
        inserted += len(rows_to_insert)
        print(f"  -> Đã nạp {inserted}/{total_rows - 1} dòng Potential ({inserted*100//(total_rows-1)}%)... [Loại A: {type_a_count}]")
        start_row = end_row + 1
        
    wb.Close(False)
    conn.close()
    print(f"Hoàn thành nạp Shop Potential: {inserted} điểm bán trong {time.time()-t0:.2f}s")
    print(f"Tổng số KH Loại A: {type_a_count}")

area_keywords = {
    'Cần Thơ': ['Cần Thơ', 'An Giang', 'Cà Mau', 'Đồng Tháp', 'Vĩnh Long', 'Kiên Giang', 'Hậu Giang', 'Bến Tre', 'Bạc Liêu', 'Sóc Trăng', 'Trà Vinh', 'Tiền Giang', 'Long An'],
    'HCM': ['Hồ Chí Minh', 'Bình Dương', 'Đồng Nai', 'Bà Rịa', 'Vũng Tàu', 'Bình Phước', 'Tây Ninh'],
    'Đà Nẵng': ['Đà Nẵng', 'Thừa Thiên Huế', 'Quảng Trị', 'Quảng Bình', 'Quảng Nam', 'Quảng Ngãi', 'Bình Định'],
    'Nha Trang': ['Khánh Hòa', 'Phú Yên', 'Ninh Thuận', 'Bình Thuận', 'Lâm Đồng', 'Đắk Lắk', 'Đắk Nông', 'Gia Lai', 'Kon Tum'],
    'Hải Phòng': ['Hải Phòng', 'Quảng Ninh', 'Hải Dương', 'Hưng Yên', 'Hà Nam', 'Nam Định', 'Thái Bình', 'Ninh Bình', 'Thanh Hóa', 'Nghệ An', 'Hà Tĩnh'],
    'Hà Nội': ['Hà Nội', 'Thái Nguyên', 'Tuyên Quang', 'Phú Thọ', 'Cao Bằng', 'Sơn La', 'Lào Cai', 'Điện Biên', 'Bắc Kạn', 'Bắc Giang', 'Bắc Ninh', 'Hà Giang', 'Hòa Bình', 'Lai Châu', 'Lạng Sơn', 'Vĩnh Phúc', 'Yên Bái']
}

def map_province_to_area(prov):
    if not prov: return ''
    for a, kws in area_keywords.items():
        if any(k in prov for k in kws):
            return a
    return ''

def export_static_js():
    print("\n[3/3] Đang xuất file dms_data.js để chạy web trực tiếp không cần server...")
    t0 = time.time()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute('''
    SELECT id, customer_code, customer_name, latitude, longitude, salesman_name, distributor_name, 
           route_code, route_name, area_name, sup_name, frequency, visit_week,
           t2, t3, t4, t5, t6, t7, cn, address, phone, province
    FROM outlets_cj
    ''')
    cj_rows = cur.fetchall()

    cur.execute('''
    SELECT id, customer_code, customer_name, latitude, longitude, salesman_name, distributor_name,
           province_new, address, ROUND(sellout_month, 0)
    FROM outlets_potential
    WHERE shop_type = 'A' AND has_gps = 1
    ''')
    raw_pot = cur.fetchall()
    conn.close()

    pot_rows = []
    for p in raw_pot:
        area = map_province_to_area(p[7])
        pot_rows.append(list(p) + [area])

    with open('dms_data.js', 'w', encoding='utf-8') as f:
        f.write('window.DMS_RAW_DATA = ')
        import json
        json.dump({'cj': cj_rows, 'pot': pot_rows}, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')

    print(f"Hoàn thành xuất dms_data.js ({os.path.getsize('dms_data.js')/(1024*1024):.2f} MB) trong {time.time()-t0:.2f}s!")

def main():
    print("=== BẮT ĐẦU QUÁ TRÌNH TẠO MASTER DATABASE ===")
    t_start = time.time()
    create_database()
    
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        import_cj_data(excel)
        import_potential_data(excel)
    finally:
        excel.Quit()
        
    export_static_js()
    print(f"\n=== ĐÃ HOÀN TẤT TOÀN BỘ TRONG {time.time()-t_start:.2f} GIÂY! ===")
    
    # In thống kê từ Database
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*), SUM(has_gps) FROM outlets_cj")
    cj_total, cj_gps = cur.fetchone()
    
    cur.execute("SELECT COUNT(*), SUM(has_gps), SUM(CASE WHEN shop_type='A' THEN 1 ELSE 0 END) FROM outlets_potential")
    pot_total, pot_gps, pot_a = cur.fetchone()
    conn.close()
    
    print("\n--- BÁO CÁO TỔNG HỢP MASTER DATABASE ---")
    print(f"• Shop CJ (MCP): {cj_total:,} điểm | Có GPS chuẩn: {cj_gps:,} ({cj_gps*100//cj_total}%)")
    print(f"• Shop Potential: {pot_total:,} điểm | Có GPS chuẩn: {pot_gps:,} ({pot_gps*100//pot_total}%)")
    print(f"• Shop Potential Loại A (cột X): {pot_a:,} điểm")
    print(f"• File Database: {DB_PATH} ({os.path.getsize(DB_PATH)/(1024*1024):.2f} MB)")

if __name__ == '__main__':
    main()
