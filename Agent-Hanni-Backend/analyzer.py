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

def find_header_row(df, keywords):
    """Find the row index that contains the header with given keywords"""
    for i in range(min(10, len(df))):
        row = df.iloc[i]
        vals = [str(v).lower().strip() for v in row.values]
        matches = sum(1 for k in keywords if any(k.lower() in v for v in vals))
        if matches >= 2:
            return i
    return None

def analyze_files(files_data):
    TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    JUNE_2026 = datetime(2026, 6, 1)

    dfs = {}
    dept_info = {}

    # === LOAD ALL FILES ===
    for f in files_data:
        fname = f['filename']
        dept = detect_department(fname)
        xl = read_excel_safe(f['content'], fname)
        if xl is None:
            continue

        sheets = {}
        engine = 'pyxlsb' if fname.lower().endswith('.xlsb') else None
        for sheet in xl.sheet_names[:10]:
            try:
                if engine:
                    df = pd.read_excel(xl, sheet_name=sheet, header=None, engine=engine)
                else:
                    df = pd.read_excel(xl, sheet_name=sheet, header=None)
                sheets[sheet] = df
            except:
                pass

        dfs[dept] = {'filename': fname, 'sheets': sheets}
        dept_info[dept] = {'filename': fname, 'status': 'OK', 'issues': []}

    anomalies = []
    stats = {'total': 0, 'critical': 0, 'risk': 0, 'watch': 0, 'ok': 0}

    # =============================================
    # 1. MASTER PLAN — backbone de toute l'analyse
    # =============================================
    master_orders = []
    master = dfs.get('Master Plan')
    if master:
        for sname, df in master['sheets'].items():
            if sname.lower() != 'plan':
                continue
            # Try multiple header positions
            header_idx = None
            for i in range(min(8, len(df))):
                row = df.iloc[i]
                vals = [str(v).lower().strip() for v in row.values]
                if sum(1 for k in ['customer', 'style', 'ship'] if any(k in v for v in vals)) >= 2:
                    header_idx = i
                    break
            
            if header_idx is None:
                print(f"Master Plan: header not found in sheet {sname}")
                continue

            print(f"Master Plan: header found at row {header_idx}")
            data = df.iloc[header_idx+1:].reset_index(drop=True)
            data.columns = [str(c).strip() for c in df.iloc[header_idx].values]
            
            # Filter valid rows
            data = data[data.apply(lambda r: str(r.get('Customer', '')).strip() not in ['', 'nan', 'Customer', 'C1'], axis=1)]
            data = data.dropna(subset=['Customer']).reset_index(drop=True)
            
            print(f"Master Plan: {len(data)} rows after filter")
            print(f"Master Plan columns: {list(data.columns[:10])}")

            for _, row in data.iterrows():
                try:
                    customer = str(row.get('Customer', '')).strip()
                    style = str(row.get('Style', '')).strip()
                    color = str(row.get('Color', '')).strip()
                    ship_date = xl_to_date(row.get('Ship Date'))
                    docket_plan = xl_to_date(row.get('Docket Plan'))
                    actual_docket = xl_to_date(row.get('Actual Docket'))
                    line = str(row.get('Line', '')).strip()

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
                        'ck': f"{customer}|{style}|{color}".upper()
                    })
                except:
                    continue
            break

    print(f"Master Plan orders (June+): {len(master_orders)}")

    # =============================================
    # 2. FABRIC STATUS — tissu reçu en WH ou non
    # =============================================
    fabric_received_codes = set()
    fabric_pending_codes = set()

    fabric = dfs.get('Fabric')
    if fabric:
        # Received sheet
        for sname, df in fabric['sheets'].items():
            if 'received' in sname.lower():
                header_idx = find_header_row(df, ['fabric', 'code', 'job'])
                if header_idx is not None:
                    data = df.iloc[header_idx+1:].reset_index(drop=True)
                    data.columns = df.iloc[header_idx].values
                    for col in data.columns:
                        if 'code' in str(col).lower() or 'fabric' in str(col).lower():
                            codes = data[col].dropna().astype(str).str.strip().str.upper()
                            fabric_received_codes.update(codes[codes != 'NAN'])
                            break

        # Pending sheet — tissu commandé mais pas encore arrivé
        for sname, df in fabric['sheets'].items():
            if 'pending' in sname.lower():
                header_idx = find_header_row(df, ['fabric', 'code', 'job'])
                if header_idx is not None:
                    data = df.iloc[header_idx+1:].reset_index(drop=True)
                    data.columns = df.iloc[header_idx].values
                    for col in data.columns:
                        if 'code' in str(col).lower() or 'fabric' in str(col).lower():
                            codes = data[col].dropna().astype(str).str.strip().str.upper()
                            fabric_pending_codes.update(codes[codes != 'NAN'])
                            break

    print(f"Fabric received: {len(fabric_received_codes)} codes | Pending: {len(fabric_pending_codes)} codes")

    # =============================================
    # 3. MERCHANDISE — statut PLAN FULL FABRIC (MER)
    # =============================================
    merch_status = {}
    merch = dfs.get('Merchandise')
    if merch:
        for sname, df in merch['sheets'].items():
            if 'tracking' in sname.lower():
                header_idx = find_header_row(df, ['customer', 'style', 'shipment'])
                if header_idx is None:
                    header_idx = 0
                data = df.iloc[header_idx+1:].reset_index(drop=True)
                data.columns = df.iloc[header_idx].values
                data = data.dropna(subset=['Customer']).reset_index(drop=True)

                for _, row in data.iterrows():
                    try:
                        customer = str(row.get('Customer', '')).strip().upper()
                        style = str(row.get('Style', '')).strip().upper()
                        color = str(row.get('Color', '')).strip().upper()
                        plan_fab = str(row.get('PLAN FULL FABRIC', '')).strip()
                        stock_out = xl_to_date(row.get('Stock Out Date'))
                        docket_rcv = xl_to_date(row.get('Docket Receive Date'))
                        fabric_out = xl_to_date(row.get('Fabric Out Date'))

                        ck = f"{customer}|{style}|{color}"
                        merch_status[ck] = {
                            'plan_full_fabric': plan_fab,
                            'stock_out': stock_out,
                            'docket_rcv': docket_rcv,
                            'fabric_out': fabric_out,
                            'mer_released': plan_fab == 'DONE'
                        }
                    except:
                        continue
                break

    print(f"Merchandise tracking: {len(merch_status)} entries")

    # =============================================
    # 4. DELIVERY PLAN — QA / inspection status
    # =============================================
    delivery_status = {}
    delivery = dfs.get('Delivery & QA')
    if delivery:
        for sname, df in delivery['sheets'].items():
            try:
                # Try to find header row
                header_idx = None
                for i in range(min(10, len(df))):
                    row = df.iloc[i]
                    vals = [str(v).lower().strip() for v in row.values]
                    if any('customer' in v for v in vals) and any('supplier' in v or 'fast' in v or 'color' in v for v in vals):
                        header_idx = i
                        break

                if header_idx is None:
                    continue

                # Try header at header_idx, then header_idx+1
                for offset in [1, 0]:
                    try:
                        data = df.iloc[header_idx+offset+1:].reset_index(drop=True)
                        cols = [str(c).strip() for c in df.iloc[header_idx+offset].values]
                        data.columns = cols

                        if 'Customer' not in cols:
                            continue

                        data = data[data['Customer'].notna()]
                        data = data[data['Customer'].astype(str).str.strip().isin(['', 'nan', 'Unplan']) == False]
                        data = data.reset_index(drop=True)

                        for _, row in data.iterrows():
                            try:
                                cust = str(row.get('Customer', '')).strip().upper()
                                fast_code = str(row.get('fast code', row.get('Fast Code', ''))).strip().upper()
                                received = str(row.get('RECEIVED/INSPECTED', '')).strip()
                                color = str(row.get('Color', '')).strip().upper()
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
                print(f"Delivery Plan sheet {sname} error: {e}")
                continue

    print(f"Delivery status: {len(delivery_status)} entries")

    # =============================================
    # 5. CROSS-CHECK — analyse par commande
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

        issues = []
        root_causes = []
        stats['total'] += 1

        # --- CHECK 1: FABRIC RECEIVED AT WH ---
        fabric_in_received = any(style.upper() in code or code in style.upper() for code in fabric_received_codes) if fabric_received_codes else False
        fabric_in_pending = any(style.upper() in code or code in style.upper() for code in fabric_pending_codes) if fabric_pending_codes else False

        merch_data = merch_status.get(ck, {})
        wh_received = merch_data.get('stock_out') is not None or fabric_in_received

        if not wh_received:
            if fabric_in_pending:
                root_causes.append('fabric_in_transit')
                issues.append('Vải đã đặt hàng nhưng chưa về kho (đang vận chuyển)')
            else:
                root_causes.append('fabric_not_booked')
                issues.append('Vải chưa được đặt hàng hoặc chưa xuất hiện trong báo cáo')

        # --- CHECK 2: MER RELEASED ---
        mer_released = merch_data.get('mer_released', False)
        if not wh_received and not mer_released:
            # Normal — fabric not there yet so MER can't be released
            pass
        elif wh_received and not mer_released:
            root_causes.append('mer_not_released')
            issues.append('Vải đã về kho nhưng MER chưa được release cho sản xuất')

        # --- CHECK 3: INTERNAL SYNC — fabric arrived but WH not updated ---
        docket_rcv = merch_data.get('docket_rcv')
        stock_out = merch_data.get('stock_out')
        if docket_rcv and not stock_out:
            days_stuck = (TODAY - docket_rcv).days
            if days_stuck >= 4:
                root_causes.append('internal_blockage')
                issues.append(f'Vải đã về xưởng {days_stuck} ngày nhưng kho chưa xác nhận nhận hàng (tắc nghẽn nội bộ)')

        # --- CHECK 4: PRODUCTION DOCKET DELAY ---
        if docket_plan and actual_docket:
            delay = (actual_docket - docket_plan).days
            if delay > 3:
                root_causes.append('production_delay')
                issues.append(f'Docket thực tế trễ {delay} ngày so với kế hoạch ({fmt_date(docket_plan)} → {fmt_date(actual_docket)})')

        # --- CHECK 5: PRE-WARNING — production approaching without material ---
        if not issues and days <= 30 and not wh_received:
            issues.append(f'Sản xuất trong {days} ngày nhưng vải chưa sẵn sàng tại kho')
            root_causes.append('pre_warning')

        if not issues:
            stats['ok'] += 1
            continue

        # --- CLASSIFY BY URGENCY ---
        if days < 0:
            level = 'CRITICAL'
            prefix = f'⚠️ Đã trễ {abs(days)} ngày — '
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

        # Build action based on root cause
        action = build_action(root_causes, customer, style, days)

        anomalies.append({
            'level': level,
            'department': get_blocking_dept(root_causes),
            'customer': customer,
            'style': style,
            'color': color,
            'shipDate': ship_str,
            'daysToShip': int(days),
            'issue': prefix + ' | '.join(issues),
            'action': action,
            'rootCauses': root_causes
        })

    # Sort by urgency then days
    level_order = {'CRITICAL': 0, 'RISK': 1, 'WATCH': 2, 'OK': 3}
    anomalies.sort(key=lambda x: (level_order.get(x['level'], 9), x.get('daysToShip', 999)))

    # Department status
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
            'summary': f"{len(dept_alerts)} vấn đề phát hiện, {critical_count} khẩn cấp" if dept_alerts else "Không có vấn đề phát hiện"
        })

    print(f"Analysis complete: {len(anomalies)} anomalies | CRITICAL: {stats['critical']} | RISK: {stats['risk']} | WATCH: {stats['watch']} | OK: {stats['ok']}")

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
        'recommendations': build_recommendations(anomalies, TODAY),
        'globalSummary': build_global_summary(anomalies, stats, len(files_data), TODAY)
    }

def get_blocking_dept(root_causes):
    if 'fabric_not_booked' in root_causes: return 'Merchandising'
    if 'fabric_in_transit' in root_causes: return 'Delivery & QA'
    if 'mer_not_released' in root_causes: return 'Merchandising'
    if 'internal_blockage' in root_causes: return 'Warehouse'
    if 'production_delay' in root_causes: return 'Master Plan'
    if 'pre_warning' in root_causes: return 'Fabric'
    return 'Merchandise'

def build_action(root_causes, customer, style, days):
    if 'fabric_not_booked' in root_causes:
        return f'Kiểm tra ngay với Merchandising — vải cho {customer}/{style} chưa được đặt hàng hoặc chưa có trong hệ thống'
    if 'fabric_in_transit' in root_causes:
        return f'Liên hệ nhà cung cấp để xác nhận ETD/ETA của vải cho {customer}/{style}'
    if 'mer_not_released' in root_causes:
        return f'Yêu cầu Merchandising release MER ngay cho {customer}/{style} — vải đã sẵn sàng'
    if 'internal_blockage' in root_causes:
        return f'Kiểm tra bộ phận Warehouse — vải đã về xưởng nhưng chưa được xác nhận nhận hàng cho {customer}/{style}'
    if 'production_delay' in root_causes:
        return f'Theo dõi sát tiến độ sản xuất line cho {customer}/{style} — docket bị trễ'
    if 'pre_warning' in root_causes:
        return f'Cảnh báo sớm — đảm bảo vải cho {customer}/{style} về kho trước ngày sản xuất'
    return f'Theo dõi tình trạng {customer}/{style}'

def build_recommendations(anomalies, today):
    recs = []
    critical = [a for a in anomalies if a['level'] == 'CRITICAL']
    overdue = [a for a in anomalies if a.get('daysToShip', 0) < 0]
    not_booked = [a for a in anomalies if 'fabric_not_booked' in a.get('rootCauses', [])]
    in_transit = [a for a in anomalies if 'fabric_in_transit' in a.get('rootCauses', [])]
    mer_pending = [a for a in anomalies if 'mer_not_released' in a.get('rootCauses', [])]
    blocked = [a for a in anomalies if 'internal_blockage' in a.get('rootCauses', [])]

    if overdue:
        custs = list(set([a['customer'] for a in overdue[:3]]))
        recs.append(f"Xử lý khẩn cấp {len(overdue)} đơn hàng đã TRỄ ngày xuất: {', '.join(custs)}")
    if critical:
        recs.append(f"Ưu tiên theo dõi {len(critical)} đơn hàng KHẨN CẤP trong 14 ngày tới")
    if not_booked:
        recs.append(f"Merchandising cần đặt hàng vải cho {len(not_booked)} đơn hàng chưa có vải trong hệ thống")
    if in_transit:
        recs.append(f"Theo dõi {len(in_transit)} lô vải đang vận chuyển — xác nhận ETD/ETA với nhà cung cấp")
    if mer_pending:
        recs.append(f"Release MER ngay cho {len(mer_pending)} đơn hàng — vải đã về kho nhưng chưa release sản xuất")
    if blocked:
        recs.append(f"Kiểm tra {len(blocked)} mã vải bị tắc nghẽn giữa Xưởng và Kho (trên 4 ngày)")

    return recs[:5] if recs else ["Tất cả đơn hàng tháng 6+ đang trong tầm kiểm soát — tiếp tục theo dõi"]

def build_global_summary(anomalies, stats, dep_count, today):
    critical = len([a for a in anomalies if a['level'] == 'CRITICAL'])
    risk = len([a for a in anomalies if a['level'] == 'RISK'])
    overdue = len([a for a in anomalies if a.get('daysToShip', 0) < 0])
    date_str = today.strftime('%d/%m/%Y')

    if not anomalies:
        return f"Ngày {date_str}: Đã phân tích {dep_count} báo cáo. Tất cả đơn hàng tháng 6 trở đi đang trong tầm kiểm soát."

    summary = f"Ngày {date_str}: Phân tích {dep_count} báo cáo — tập trung vào đơn hàng tháng 6 trở đi. "
    if overdue > 0:
        summary += f"{overdue} đơn hàng đã TRỄ ngày xuất. "
    if critical > 0:
        summary += f"{critical} đơn hàng KHẨN CẤP cần xử lý ngay. "
    if risk > 0:
        summary += f"{risk} đơn hàng có RỦI RO trong 28 ngày tới. "
    summary += "Xem chi tiết bên dưới."

    return summary