import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

def xl_to_date(val):
    try:
        v = float(val)
        if not np.isnan(v) and v > 1000:
            return datetime(1899, 12, 30) + timedelta(days=int(v))
    except:
        pass
    return None

def fmt_date(d):
    if isinstance(d, datetime):
        return d.strftime('%d/%m/%Y')
    return ''

def detect_department(filename):
    f = filename.lower()
    if 'fabric' in f: return 'Fabric'
    if 'merchandise' in f or 'merch' in f: return 'Merchandise'
    if 'delivery' in f or 'inspection' in f: return 'Delivery & QA'
    if 'master' in f: return 'Master Plan'
    if 'shipment' in f or 'tracking' in f: return 'Shipment'
    if 'daily_report' in f or 'daily report' in f: return 'Daily Report'
    return 'Autre'

def read_excel_safe(file_bytes, filename):
    ext = filename.lower().split('.')[-1]
    try:
        if ext == 'xlsb':
            return pd.ExcelFile(io.BytesIO(file_bytes), engine='pyxlsb')
        else:
            return pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception as e:
        print(f"Error reading {filename}: {e}")
        return None

def analyze_files(files_data):
    TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    JUNE_2026 = datetime(2026, 6, 1)

    dfs = {}

    # === LOAD ALL FILES ===
    for f in files_data:
        fname = f['filename']
        dept = detect_department(fname)
        xl = read_excel_safe(f['content'], fname)
        if xl is None:
            continue
        sheets = {}
        engine = 'pyxlsb' if fname.lower().endswith('.xlsb') else None
        for sheet in xl.sheet_names:
            try:
                if engine:
                    df = pd.read_excel(xl, sheet_name=sheet, header=None, engine=engine)
                else:
                    df = pd.read_excel(xl, sheet_name=sheet, header=None)
                sheets[sheet] = df
            except:
                pass
        dfs[dept] = {'filename': fname, 'sheets': sheets}

    anomalies = []
    stats = {'total': 0, 'critical': 0, 'risk': 0, 'watch': 0, 'ok': 0}

    # =============================================
    # 1. MASTER PLAN — header fixe ligne 3
    # =============================================
    master_orders = []
    master = dfs.get('Master Plan')
    if master:
        for sname, df in master['sheets'].items():
            if sname.lower() != 'plan':
                continue

            # Header is at row index 3
            HEADER_ROW = 3
            headers = [str(v).strip() for v in df.iloc[HEADER_ROW].values]
            data = df.iloc[HEADER_ROW+1:].reset_index(drop=True)
            data.columns = headers

            # Filter valid rows
            data = data[
                data['Customer'].notna() &
                (data['Customer'].astype(str).str.strip() != '') &
                (data['Customer'].astype(str).str.strip() != 'nan') &
                (data['Line'].astype(str).str.strip() != 'C1') &
                (data['Line'].astype(str).str.strip() != 'Line')
            ].reset_index(drop=True)

            print(f"Master Plan rows loaded: {len(data)}")

            for _, row in data.iterrows():
                try:
                    customer = str(row.get('Customer', '')).strip()
                    style = str(row.get('Style', '')).strip()
                    color = str(row.get('Color', '')).strip()
                    ship_date = xl_to_date(row.get('Ship Date'))
                    docket_plan = xl_to_date(row.get('Docket Plan'))
                    actual_docket = xl_to_date(row.get('Actual Docket'))
                    line = str(row.get('Line', '')).strip()

                    # Fabric codes from Master Plan
                    fabric_main = str(row.get('Main', '')).strip().upper()
                    fabric_contrast1 = str(row.get('Contrast 1', '')).strip().upper()
                    fabric_rib = str(row.get('Rib', '')).strip().upper()
                    fabric_codes = [c for c in [fabric_main, fabric_contrast1, fabric_rib] if c and c != 'NAN' and c != 'OK']

                    if not ship_date or not customer or customer == 'nan':
                        continue

                    # FOCUS JUIN+ UNIQUEMENT
                    if ship_date < JUNE_2026:
                        continue

                    days_to_ship = (ship_date - TODAY).days

                    master_orders.append({
                        'customer': customer,
                        'style': style,
                        'color': color,
                        'ship_date': ship_date,
                        'ship_date_str': fmt_date(ship_date),
                        'days_to_ship': days_to_ship,
                        'docket_plan': docket_plan,
                        'actual_docket': actual_docket,
                        'line': line,
                        'fabric_codes': fabric_codes,
                        'ck': f"{customer}|{style}|{color}".upper()
                    })
                except:
                    continue
            break

    print(f"Master Plan orders (June+): {len(master_orders)}")

    # =============================================
    # 2. FABRIC — codes reçus et en attente
    # =============================================
    fabric_received_codes = set()
    fabric_pending_codes = set()

    fabric = dfs.get('Fabric')
    if fabric:
        for sname, df in fabric['sheets'].items():
            if 'received' in sname.lower():
                for i in range(min(5, len(df))):
                    row = df.iloc[i]
                    vals = [str(v).lower().strip() for v in row.values]
                    if any('fabric' in v or 'code' in v or 'job' in v for v in vals):
                        data = df.iloc[i+1:].reset_index(drop=True)
                        data.columns = [str(c).strip() for c in df.iloc[i].values]
                        for col in data.columns:
                            if 'code' in str(col).lower() or 'fabric' in str(col).lower():
                                codes = data[col].dropna().astype(str).str.strip().str.upper()
                                fabric_received_codes.update(c for c in codes if c not in ['NAN', ''])
                                break
                        break

            if 'pending' in sname.lower():
                for i in range(min(5, len(df))):
                    row = df.iloc[i]
                    vals = [str(v).lower().strip() for v in row.values]
                    if any('fabric' in v or 'code' in v or 'job' in v for v in vals):
                        data = df.iloc[i+1:].reset_index(drop=True)
                        data.columns = [str(c).strip() for c in df.iloc[i].values]
                        for col in data.columns:
                            if 'code' in str(col).lower() or 'fabric' in str(col).lower():
                                codes = data[col].dropna().astype(str).str.strip().str.upper()
                                fabric_pending_codes.update(c for c in codes if c not in ['NAN', ''])
                                break
                        break

    print(f"Fabric received: {len(fabric_received_codes)} | Pending: {len(fabric_pending_codes)}")

    # =============================================
    # 3. MERCHANDISE — MER status (PLAN FULL FABRIC)
    # =============================================
    merch_status = {}
    merch = dfs.get('Merchandise')
    if merch:
        for sname, df in merch['sheets'].items():
            if 'tracking' not in sname.lower():
                continue
            for i in range(min(5, len(df))):
                row = df.iloc[i]
                vals = [str(v).lower().strip() for v in row.values]
                if any('customer' in v for v in vals):
                    data = df.iloc[i+1:].reset_index(drop=True)
                    data.columns = [str(c).strip() for c in df.iloc[i].values]
                    data = data.dropna(subset=['Customer']).reset_index(drop=True)
                    for _, r in data.iterrows():
                        try:
                            customer = str(r.get('Customer', '')).strip().upper()
                            style = str(r.get('Style', '')).strip().upper()
                            color = str(r.get('Color', '')).strip().upper()
                            plan_fab = str(r.get('PLAN FULL FABRIC', '')).strip()
                            stock_out = xl_to_date(r.get('Stock Out Date'))
                            docket_rcv = xl_to_date(r.get('Docket Receive Date'))
                            ck = f"{customer}|{style}|{color}"
                            merch_status[ck] = {
                                'mer_released': plan_fab == 'DONE',
                                'stock_out': stock_out,
                                'docket_rcv': docket_rcv,
                            }
                        except:
                            continue
                    break
            break

    print(f"Merchandise tracking: {len(merch_status)} entries")

    # =============================================
    # 4. DELIVERY PLAN — QA status
    # =============================================
    delivery_status = {}
    delivery = dfs.get('Delivery & QA')
    if delivery:
        for sname, df in delivery['sheets'].items():
            try:
                header_idx = None
                for i in range(min(10, len(df))):
                    vals = [str(v).lower().strip() for v in df.iloc[i].values]
                    if any('customer' in v for v in vals) and any('supplier' in v or 'fast' in v for v in vals):
                        header_idx = i
                        break
                if header_idx is None:
                    continue

                for offset in [1, 0]:
                    try:
                        cols = [str(c).strip() for c in df.iloc[header_idx+offset].values]
                        if 'Customer' not in cols:
                            continue
                        data = df.iloc[header_idx+offset+1:].reset_index(drop=True)
                        data.columns = cols
                        data = data[data['Customer'].notna() & ~data['Customer'].astype(str).str.strip().isin(['', 'nan', 'Unplan'])]
                        for _, r in data.iterrows():
                            try:
                                cust = str(r.get('Customer', '')).strip().upper()
                                fast_code = str(r.get('fast code', r.get('Fast Code', ''))).strip().upper()
                                received = str(r.get('RECEIVED/INSPECTED', '')).strip()
                                color = str(r.get('Color', '')).strip().upper()
                                ck = f"{cust}|{fast_code}|{color}"
                                delivery_status[ck] = {
                                    'received': received not in ['', 'nan'],
                                    'fast_code': fast_code
                                }
                            except:
                                continue
                        break
                    except:
                        continue
            except Exception as e:
                print(f"Delivery error: {e}")
                continue

    print(f"Delivery status: {len(delivery_status)} entries")

    # =============================================
    # 5. CROSS-CHECK PAR COMMANDE
    # =============================================
    for order in master_orders:
        days = order['days_to_ship']
        ck = order['ck']
        customer = order['customer']
        style = order['style']
        color = order['color']
        ship_str = order['ship_date_str']
        docket_plan = order['docket_plan']
        actual_docket = order['actual_docket']
        fabric_codes = order['fabric_codes']

        issues = []
        root_causes = []
        stats['total'] += 1

        merch_data = merch_status.get(ck, {})

    # --- CHECK 1: FABRIC RECEIVED AT WH ---
        fabric_received = False
        fabric_in_transit = False
        missing_codes = []

        if fabric_codes:
            for code in fabric_codes:
                if code in fabric_received_codes:
                    fabric_received = True
                elif code in fabric_pending_codes:
                    fabric_in_transit = True
                    missing_codes.append(code)
                else:
                    missing_codes.append(code)

        # Also check via merch stock_out date
        if merch_data.get('stock_out'):
            fabric_received = True
            missing_codes = []

        if not fabric_received:
            if missing_codes:
                codes_str = ', '.join(missing_codes[:3])
                if fabric_in_transit:
                    root_causes.append('fabric_in_transit')
                    issues.append(f'Vải đang vận chuyển, chưa về kho — mã: {codes_str}')
                else:
                    root_causes.append('fabric_not_received')
                    issues.append(f'Vải chưa nhận tại kho — mã: {codes_str}')
            elif not fabric_codes:
                root_causes.append('fabric_code_missing')
                issues.append('Không tìm thấy mã vải trong Master Plan')
        # --- CHECK 2: MER RELEASED ---
        mer_released = merch_data.get('mer_released', False)
        if fabric_received and not mer_released and ck in merch_status:
            root_causes.append('mer_not_released')
            issues.append('Vải đã về kho nhưng MER chưa được release cho sản xuất')

        # --- CHECK 3: INTERNAL BLOCKAGE ---
        docket_rcv = merch_data.get('docket_rcv')
        stock_out = merch_data.get('stock_out')
        if docket_rcv and not stock_out:
            days_stuck = (TODAY - docket_rcv).days
            if days_stuck >= 4:
                root_causes.append('internal_blockage')
                issues.append(f'Vải về xưởng {days_stuck} ngày nhưng kho chưa xác nhận (tắc nghẽn nội bộ)')

        # --- CHECK 4: PRODUCTION DELAY ---
        if docket_plan and actual_docket:
            delay = (actual_docket - docket_plan).days
            if delay > 3:
                root_causes.append('production_delay')
                issues.append(f'Docket trễ {delay} ngày (KH: {fmt_date(docket_plan)} → TT: {fmt_date(actual_docket)})')

        # --- CHECK 5: PRE-WARNING ---
        if not issues and days <= 30 and not fabric_received:
            issues.append(f'Cảnh báo sớm — sản xuất trong {days} ngày nhưng vải chưa sẵn sàng')
            root_causes.append('pre_warning')

        if not issues:
            stats['ok'] += 1
            continue

        # Classify
        if days < 0:
            level = 'CRITICAL'
            prefix = f'⚠️ Trễ {abs(days)} ngày — '
        elif days <= 7:
            level = 'CRITICAL'
            prefix = f'🔴 Còn {days} ngày — '
        elif days <= 14:
            level = 'CRITICAL'
            prefix = ''
        elif days <= 28:
            level = 'RISK'
            prefix = ''
        else:
            level = 'WATCH'
            prefix = ''

        stats[level.lower()] += 1

        anomalies.append({
            'level': level,
            'department': get_blocking_dept(root_causes),
            'customer': customer,
            'style': style,
            'color': color,
            'shipDate': ship_str,
            'daysToShip': int(days),
            'issue': prefix + ' | '.join(issues),
            'action': build_action(root_causes, customer, style, days),
            'rootCauses': root_causes,
            'fabricCodes': fabric_codes
            'fabricCodes': missing_codes
        })

    level_order = {'CRITICAL': 0, 'RISK': 1, 'WATCH': 2, 'OK': 3}
    anomalies.sort(key=lambda x: (level_order.get(x['level'], 9), x.get('daysToShip', 999)))

    dept_status = []
    for dept in ['Master Plan', 'Merchandise', 'Fabric', 'Delivery & QA', 'Shipment', 'Daily Report']:
        if dept not in dfs:
            continue
        dept_alerts = [a for a in anomalies if a['department'] == dept]
        critical_count = len([a for a in dept_alerts if a['level'] == 'CRITICAL'])
        status = 'CRITICAL' if critical_count > 0 else ('WARNING' if dept_alerts else 'OK')
        dept_status.append({
            'name': dept,
            'status': status,
            'summary': f"{len(dept_alerts)} vấn đề, {critical_count} khẩn cấp" if dept_alerts else "Không có vấn đề phát hiện"
        })

    print(f"Done: {len(anomalies)} anomalies | CRITICAL: {stats.get('critical',0)} | RISK: {stats.get('risk',0)} | WATCH: {stats.get('watch',0)} | OK: {stats['ok']}")

    return {
        'summary': {
            'totalOrders': stats['total'],
            'critical': len([a for a in anomalies if a['level'] == 'CRITICAL']),
            'risk': len([a for a in anomalies if a['level'] == 'RISK']),
            'watch': len([a for a in anomalies if a['level'] == 'WATCH']),
            'onTrack': stats['ok'],
            'depsReceived': len(files_data)
        },
        'alerts': anomalies[:150],
        'departmentStatus': dept_status,
        'recommendations': build_recommendations(anomalies),
        'globalSummary': build_global_summary(anomalies, stats, len(files_data), TODAY)
    }

def get_blocking_dept(root_causes):
    if 'fabric_not_received' in root_causes: return 'Fabric'
    if 'fabric_in_transit' in root_causes: return 'Delivery & QA'
    if 'fabric_code_missing' in root_causes: return 'Merchandising'
    if 'mer_not_released' in root_causes: return 'Merchandising'
    if 'internal_blockage' in root_causes: return 'Warehouse'
    if 'production_delay' in root_causes: return 'Master Plan'
    if 'pre_warning' in root_causes: return 'Fabric'
    return 'Merchandising'

def build_action(root_causes, customer, style, days):
    if 'fabric_code_missing' in root_causes:
        return f'Vérifier avec Merchandising — mã vải cho {customer}/{style} chưa có trong Master Plan'
    if 'fabric_not_received' in root_causes:
        return f'Liên hệ nhà cung cấp — vải cho {customer}/{style} chưa về kho'
    if 'fabric_in_transit' in root_causes:
        return f'Xác nhận ETD/ETA với nhà cung cấp cho {customer}/{style}'
    if 'mer_not_released' in root_causes:
        return f'Yêu cầu Merchandising release MER ngay cho {customer}/{style}'
    if 'internal_blockage' in root_causes:
        return f'Kiểm tra Warehouse — vải {customer}/{style} chưa được xác nhận nhận hàng'
    if 'production_delay' in root_causes:
        return f'Theo dõi tiến độ sản xuất cho {customer}/{style}'
    if 'pre_warning' in root_causes:
        return f'Đảm bảo vải {customer}/{style} về kho trước ngày sản xuất'
    return f'Theo dõi {customer}/{style}'

def build_recommendations(anomalies):
    recs = []
    overdue = [a for a in anomalies if a.get('daysToShip', 0) < 0]
    critical = [a for a in anomalies if a['level'] == 'CRITICAL']
    not_received = [a for a in anomalies if 'fabric_not_received' in a.get('rootCauses', [])]
    in_transit = [a for a in anomalies if 'fabric_in_transit' in a.get('rootCauses', [])]
    mer_pending = [a for a in anomalies if 'mer_not_released' in a.get('rootCauses', [])]
    blocked = [a for a in anomalies if 'internal_blockage' in a.get('rootCauses', [])]

    if overdue:
        custs = list(set([a['customer'] for a in overdue[:3]]))
        recs.append(f"Xử lý khẩn cấp {len(overdue)} đơn hàng đã TRỄ: {', '.join(custs)}")
    if critical:
        recs.append(f"Ưu tiên {len(critical)} đơn hàng KHẨN CẤP trong 14 ngày tới")
    if not_received:
        recs.append(f"Theo dõi {len(not_received)} đơn hàng chưa nhận vải tại kho")
    if in_transit:
        recs.append(f"Xác nhận ETD/ETA cho {len(in_transit)} lô vải đang vận chuyển")
    if mer_pending:
        recs.append(f"Release MER ngay cho {len(mer_pending)} đơn hàng — vải đã sẵn sàng")
    if blocked:
        recs.append(f"Kiểm tra {len(blocked)} mã vải tắc nghẽn giữa Xưởng và Kho")

    return recs[:5] if recs else ["Tất cả đơn hàng tháng 6+ đang trong tầm kiểm soát"]

def build_global_summary(anomalies, stats, dep_count, today):
    critical = len([a for a in anomalies if a['level'] == 'CRITICAL'])
    risk = len([a for a in anomalies if a['level'] == 'RISK'])
    overdue = len([a for a in anomalies if a.get('daysToShip', 0) < 0])
    date_str = today.strftime('%d/%m/%Y')

    if not anomalies:
        return f"Ngày {date_str}: Phân tích {dep_count} báo cáo. Tất cả đơn hàng tháng 6+ đang trong tầm kiểm soát."

    summary = f"Ngày {date_str}: Phân tích {dep_count} báo cáo, tập trung đơn hàng tháng 6+. "
    if overdue > 0:
        summary += f"{overdue} đơn hàng đã TRỄ ngày xuất. "
    if critical > 0:
        summary += f"{critical} đơn hàng KHẨN CẤP cần xử lý ngay. "
    if risk > 0:
        summary += f"{risk} đơn hàng có RỦI RO. "
    summary += "Xem chi tiết bên dưới."
    return summary