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
    if 'erp' in f: return 'ERP'
    if 'fabric' in f and 'daily' in f: return 'Fabric'
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
        elif ext == 'xls':
            return pd.ExcelFile(io.BytesIO(file_bytes), engine='xlrd')
        else:
            return pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception as e:
        print(f"Error reading {filename}: {e}")
        return None

def analyze_files(files_data):
    TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    JUNE_2026 = datetime(2026, 6, 1)

    dfs = {}

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
        print(f"Loaded: {dept} ({fname}) — {len(sheets)} sheets")

    anomalies = []
    stats = {'total': 0, 'critical': 0, 'risk': 0, 'watch': 0, 'ok': 0}

    # =============================================
    # 1. MASTER PLAN — source of truth
    # =============================================
    master_orders = []
    master = dfs.get('Master Plan')
    if master:
        for sname, df in master['sheets'].items():
            if sname.lower() != 'plan':
                continue
            HEADER_ROW = 3
            headers = [str(v).strip() for v in df.iloc[HEADER_ROW].values]
            data = df.iloc[HEADER_ROW+1:].reset_index(drop=True)
            data.columns = headers
            data = data[
                data['Customer'].notna() &
                (data['Customer'].astype(str).str.strip() != '') &
                (data['Customer'].astype(str).str.strip() != 'nan') &
                (data['Line'].astype(str).str.strip() != 'C1') &
                (data['Line'].astype(str).str.strip() != 'Line')
            ].reset_index(drop=True)
            print(f"Master Plan rows: {len(data)}")
            for _, row in data.iterrows():
                try:
                    customer = str(row.get('Customer', '')).strip()
                    style = str(row.get('Style', '')).strip()
                    color = str(row.get('Color', '')).strip()
                    ship_date = xl_to_date(row.get('Ship Date'))
                    docket_plan = xl_to_date(row.get('Docket Plan'))
                    actual_docket = xl_to_date(row.get('Actual Docket'))
                    line = str(row.get('Line', '')).strip()
                    fabric_main = str(row.get('Main', '')).strip().upper()
                    fabric_codes = [
                        c for c in [fabric_main]
                        if c and c != 'NAN' and c != 'OK' and not c.replace('.', '').isdigit()
                    ]
                    if not ship_date or not customer or customer == 'nan':
                        continue
                    if ship_date < JUNE_2026:
                        continue
                    master_orders.append({
                        'customer': customer,
                        'style': style,
                        'color': color,
                        'ship_date': ship_date,
                        'ship_date_str': fmt_date(ship_date),
                        'days_to_ship': (ship_date - TODAY).days,
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
    # 2. ERP — supplier confirmation status
    # =============================================
    erp_status = {}
    erp = dfs.get('ERP')
    if erp:
        for sname, df in erp['sheets'].items():
            try:
                # Header at row 0
                data = pd.read_excel(
                    io.BytesIO(list(erp.get('_raw', {}).values() or [b''])[0]) if False else
                    next(iter([s for s in [df]]), df),
                    header=None
                ) if False else df
                # Find header
                for i in range(min(3, len(df))):
                    vals = [str(v).lower() for v in df.iloc[i].values]
                    if any('fabric fast' in v or 'fast code' in v for v in vals):
                        data = df.iloc[i+1:].reset_index(drop=True)
                        data.columns = [str(c).strip() for c in df.iloc[i].values]
                        for _, row in data.iterrows():
                            try:
                                code = str(row.get('Fabric FAST Code', '')).strip().upper()
                                status = str(row.get('Master PO Status', '')).strip()
                                confirm_date = row.get('Confirm Expect Stock in Date')
                                revised_date = row.get('Revised Confirmed Stock in Date')
                                if code and code != 'NAN':
                                    erp_status[code] = {
                                        'status': status,
                                        'confirm_date': pd.Timestamp(confirm_date).to_pydatetime() if pd.notna(confirm_date) else None,
                                        'revised_date': pd.Timestamp(revised_date).to_pydatetime() if pd.notna(revised_date) else None,
                                    }
                            except:
                                continue
                        break
            except Exception as e:
                print(f"ERP error: {e}")
                continue
    print(f"ERP status: {len(erp_status)} fabric codes")

    # =============================================
    # 3. FABRIC DAILY REPORT — CheckedIn (inspected)
    # =============================================
    fabric_checkedin = set()
    fabric = dfs.get('Fabric')
    if fabric:
        for sname, df in fabric['sheets'].items():
            if 'checkedin' in sname.lower():
                for i in range(min(6, len(df))):
                    vals = [str(v).lower() for v in df.iloc[i].values]
                    if any('fabric code' in v or 'fabric' in v for v in vals) and any('date' in v for v in vals):
                        data = df.iloc[i+1:].reset_index(drop=True)
                        data.columns = [str(c).strip() for c in df.iloc[i].values]
                        for col in data.columns:
                            if 'fabric code' in str(col).lower() or col == 'Fabric Code':
                                codes = data[col].dropna().astype(str).str.strip().str.upper()
                                fabric_checkedin.update(c for c in codes if c not in ['NAN', ''])
                                break
                        break
    print(f"Fabric CheckedIn: {len(fabric_checkedin)} codes")

    # =============================================
    # 4. DAILY REPORT — WH received (MãVậtTư)
    # =============================================
    wh_received = set()
    daily = dfs.get('Daily Report')
    if daily:
        for sname, df in daily['sheets'].items():
            if 'daily' in sname.lower():
                for i in range(min(5, len(df))):
                    vals = [str(v).lower() for v in df.iloc[i].values]
                    if any('mãvậttư' in v or 'mavat' in v or 'fabric' in v for v in vals):
                        data = df.iloc[i+1:].reset_index(drop=True)
                        data.columns = [str(c).strip() for c in df.iloc[i].values]
                        for col in data.columns:
                            if 'mãvậttư' in str(col).lower() or 'mavat' in str(col).lower():
                                codes = data[col].dropna().astype(str).str.strip().str.upper()
                                wh_received.update(c for c in codes if c not in ['NAN', ''])
                                break
                        break
    print(f"WH received: {len(wh_received)} codes")

    # =============================================
    # 5. MERCHANDISE — MER released
    # =============================================
    merch_status = {}
    merch = dfs.get('Merchandise')
    if merch:
        for sname, df in merch['sheets'].items():
            if 'tracking' in sname.lower():
                for i in range(min(3, len(df))):
                    vals = [str(v).lower() for v in df.iloc[i].values]
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
                                ck = f"{customer}|{style}|{color}"
                                merch_status[ck] = {
                                    'mer_released': plan_fab == 'DONE',
                                    'stock_out': stock_out,
                                }
                            except:
                                continue
                        break
            break
    print(f"Merchandise tracking: {len(merch_status)} entries")

    # =============================================
    # 6. DELIVERY PLAN — scheduled for delivery
    # =============================================
    delivery_planned = set()
    delivery = dfs.get('Delivery & QA')
    if delivery:
        for sname, df in delivery['sheets'].items():
            try:
                for i in range(min(8, len(df))):
                    vals = [str(v).lower() for v in df.iloc[i].values]
                    if any('fast code' in v or 'fast' in v for v in vals) and any('customer' in v for v in vals):
                        data = df.iloc[i+1:].reset_index(drop=True)
                        data.columns = [str(c).strip() for c in df.iloc[i].values]
                        for col in data.columns:
                            if 'fast code' in str(col).lower() or col == 'fast code':
                                codes = data[col].dropna().astype(str).str.strip().str.upper()
                                delivery_planned.update(c for c in codes if c not in ['NAN', ''])
                                break
                        break
            except:
                continue
    print(f"Delivery planned: {len(delivery_planned)} codes")

    # =============================================
    # 7. CROSS-CHECK — 5-step flow per order
    # =============================================
    PRIORITY_CUSTOMERS = {'ALD', 'GOLF WANG', 'RODD & GUNN', 'CORTEIZ', 'STUSSY', 'RAPHA'}

    for order in master_orders:
        days = order['days_to_ship']
        ck = order['ck']
        customer = order['customer']
        style = order['style']
        color = order['color']
        ship_str = order['ship_date_str']
        fabric_codes = order['fabric_codes']
        docket_plan = order['docket_plan']
        actual_docket = order['actual_docket']

        issues = []
        root_causes = []
        blocking_dept = None
        stats['total'] += 1

        merch_data = merch_status.get(ck, {})

        for code in fabric_codes:
            erp_data = erp_status.get(code, {})
            erp_st = erp_data.get('status', '')
            revised_date = erp_data.get('revised_date')
            confirm_date = erp_data.get('confirm_date')

            # STEP 1 — ERP: supplier confirmation
            if erp_st == 'Over-due':
                days_overdue = (TODAY - (revised_date or confirm_date)).days if (revised_date or confirm_date) else '?'
                issues.append(f'Nhà cung cấp TRỄ {days_overdue} ngày (ERP Over-due) — mã: {code}')
                root_causes.append('erp_overdue')
                blocking_dept = 'Purchasing'
                break

            if erp_st == 'On-due in next 10 days':
                issues.append(f'Vải sắp đến hạn trong 10 ngày (ERP) — mã: {code}')
                root_causes.append('erp_ondue')
                blocking_dept = 'Purchasing'
                break

            # STEP 2 — Delivery Plan: scheduled?
            if code not in delivery_planned and erp_st not in ['Done']:
                issues.append(f'Chưa có trong Delivery Plan — mã: {code}')
                root_causes.append('not_in_delivery')
                blocking_dept = 'Purchasing'
                break

            # STEP 3 — Fabric Daily Report: inspected?
            if code not in fabric_checkedin and erp_st not in ['Done']:
                issues.append(f'Chưa được kiểm tra QA — mã: {code}')
                root_causes.append('not_inspected')
                blocking_dept = 'QA'
                break

            # STEP 4 — Daily Report: WH received?
            if code not in wh_received and erp_st not in ['Done']:
                issues.append(f'Chưa nhận vào kho — mã: {code}')
                root_causes.append('not_in_wh')
                blocking_dept = 'Warehouse'
                break

        # STEP 5 — MER released?
        if not issues:
            mer_released = merch_data.get('mer_released', False)
            if not mer_released and ck in merch_status:
                issues.append('Vải đã sẵn sàng nhưng MER chưa được release cho sản xuất')
                root_causes.append('mer_not_released')
                blocking_dept = 'Merchandising'

        # Docket delay check
        if docket_plan and actual_docket:
            delay = (actual_docket - docket_plan).days
            if delay > 3:
                issues.append(f'Docket trễ {delay} ngày (KH: {fmt_date(docket_plan)} → TT: {fmt_date(actual_docket)})')
                root_causes.append('production_delay')
                if not blocking_dept:
                    blocking_dept = 'Production'

        if not issues:
            stats['ok'] += 1
            continue

        # Classify by urgency
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
            'department': blocking_dept or 'Unknown',
            'customer': customer,
            'style': style,
            'color': color,
            'shipDate': ship_str,
            'daysToShip': int(days),
            'issue': prefix + ' | '.join(issues),
            'action': build_action(root_causes, customer, style, days, blocking_dept),
            'rootCauses': root_causes,
            'fabricCodes': fabric_codes
        })

    # Sort — priority customers first, then by urgency
    level_order = {'CRITICAL': 0, 'RISK': 1, 'WATCH': 2, 'OK': 3}
    anomalies.sort(key=lambda x: (
        0 if x['customer'].upper() in PRIORITY_CUSTOMERS else 1,
        level_order.get(x['level'], 9),
        x.get('daysToShip', 999)
    ))

    dept_status = []
    for dept in ['ERP', 'Purchasing', 'QA', 'Warehouse', 'Merchandising', 'Production']:
        dept_alerts = [a for a in anomalies if a['department'] == dept]
        critical_count = len([a for a in dept_alerts if a['level'] == 'CRITICAL'])
        if not dept_alerts:
            continue
        status = 'CRITICAL' if critical_count > 0 else 'WARNING'
        dept_status.append({
            'name': dept,
            'status': status,
            'summary': f"{len(dept_alerts)} vấn đề, {critical_count} khẩn cấp"
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

def build_action(root_causes, customer, style, days, dept):
    if 'erp_overdue' in root_causes:
        return f'Liên hệ ngay Purchasing — nhà cung cấp trễ giao vải cho {customer}/{style}'
    if 'erp_ondue' in root_causes:
        return f'Theo dõi sát với Purchasing — vải {customer}/{style} sắp đến hạn'
    if 'not_in_delivery' in root_causes:
        return f'Yêu cầu Purchasing thêm vải {customer}/{style} vào Delivery Plan'
    if 'not_inspected' in root_causes:
        return f'Yêu cầu QA kiểm tra vải {customer}/{style} ngay'
    if 'not_in_wh' in root_causes:
        return f'Kiểm tra Warehouse — vải {customer}/{style} chưa được nhận vào kho'
    if 'mer_not_released' in root_causes:
        return f'Yêu cầu Merchandising release MER ngay cho {customer}/{style}'
    if 'production_delay' in root_causes:
        return f'Theo dõi tiến độ sản xuất cho {customer}/{style}'
    return f'Theo dõi {customer}/{style}'

def build_recommendations(anomalies):
    recs = []
    overdue = [a for a in anomalies if a.get('daysToShip', 0) < 0]
    erp_od = [a for a in anomalies if 'erp_overdue' in a.get('rootCauses', [])]
    not_del = [a for a in anomalies if 'not_in_delivery' in a.get('rootCauses', [])]
    not_insp = [a for a in anomalies if 'not_inspected' in a.get('rootCauses', [])]
    not_wh = [a for a in anomalies if 'not_in_wh' in a.get('rootCauses', [])]
    mer = [a for a in anomalies if 'mer_not_released' in a.get('rootCauses', [])]

    if overdue:
        recs.append(f"Xử lý khẩn cấp {len(overdue)} đơn hàng đã TRỄ ngày xuất")
    if erp_od:
        recs.append(f"Purchasing: {len(erp_od)} nhà cung cấp trễ giao vải — cần escalate ngay")
    if not_del:
        recs.append(f"Purchasing: thêm {len(not_del)} mã vải vào Delivery Plan")
    if not_insp:
        recs.append(f"QA: kiểm tra {len(not_insp)} lô vải đang chờ inspection")
    if not_wh:
        recs.append(f"Warehouse: xác nhận nhận hàng cho {len(not_wh)} lô vải")
    if mer:
        recs.append(f"Merchandising: release MER ngay cho {len(mer)} đơn hàng")

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
        summary += f"{critical} đơn hàng KHẨN CẤP. "
    if risk > 0:
        summary += f"{risk} đơn hàng RỦI RO. "
    summary += "Xem chi tiết bên dưới."
    return summary