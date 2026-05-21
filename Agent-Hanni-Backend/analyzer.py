import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

def xl_to_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.replace(hour=0, minute=0, second=0, microsecond=0)
    if hasattr(val, 'to_pydatetime'):
        return val.to_pydatetime().replace(hour=0, minute=0, second=0, microsecond=0)
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
    if 'fabric' in f: return 'Fabric'
    if 'merchandise' in f or 'merch' in f: return 'Merchandise'
    if 'delivery' in f or 'inspection' in f: return 'Delivery & QA'
    if 'master' in f: return 'Master Plan'
    if 'shipment' in f or 'tracking' in f: return 'Shipment'
    if 'daily' in f: return 'Daily Report'
    return 'Autre'

def read_sheet_direct(file_bytes, filename, sheet_name):
    ext = filename.lower().split('.')[-1]
    try:
        if ext == 'xlsb':
            return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, engine='pyxlsb')
        elif ext == 'xls':
            return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, engine='xlrd')
        else:
            return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
    except Exception as e:
        print(f"Error reading sheet '{sheet_name}' from {filename}: {e}")
        return None

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
        sheets = {}

        if dept == 'Shipment':
            dfs[dept] = {'filename': fname, 'sheets': {}}
            print(f"Skipped: {dept} ({fname})")
            continue

        try:
            if dept == 'Master Plan':
                df = read_sheet_direct(f['content'], fname, 'Plan')
                if df is not None:
                    sheets = {'Plan': df}
                print(f"Loaded: {dept} ({fname}) — {len(sheets)} sheets")

            elif dept == 'Merchandise':
                df = read_sheet_direct(f['content'], fname, 'Fabric Tracking')
                if df is not None:
                    sheets = {'Fabric Tracking': df}
                print(f"Loaded: {dept} ({fname}) — {len(sheets)} sheets")

            elif dept == 'Daily Report':
                df = read_sheet_direct(f['content'], fname, 'Daily Report')
                if df is not None:
                    sheets = {'Daily Report': df}
                print(f"Loaded: {dept} ({fname}) — {len(sheets)} sheets")

            else:
                xl = read_excel_safe(f['content'], fname)
                if xl:
                    ext = fname.lower().split('.')[-1]
                    engine = 'pyxlsb' if ext == 'xlsb' else None
                    for sname in xl.sheet_names[:10]:
                        try:
                            if engine:
                                df = pd.read_excel(xl, sheet_name=sname, header=None, engine=engine)
                            else:
                                df = pd.read_excel(xl, sheet_name=sname, header=None)
                            sheets[sname] = df
                        except:
                            pass
                    xl.close()
                print(f"Loaded: {dept} ({fname}) — {len(sheets)} sheets")

        except Exception as e:
            print(f"Error loading {dept} ({fname}): {e}")
            sheets = {}

        dfs[dept] = {'filename': fname, 'sheets': sheets}

    anomalies = []
    stats = {'total': 0, 'critical': 0, 'risk': 0, 'watch': 0, 'ok': 0}

    # 1. MASTER PLAN
    master_orders = []
    master = dfs.get('Master Plan')
    if master:
        for sname, df in master['sheets'].items():
            if sname.lower() != 'plan':
                continue
            headers = [str(v).strip() for v in df.iloc[3].values]
            data = df.iloc[5:].reset_index(drop=True)
            data.columns = headers
            data = data[
                data['Customer'].notna() &
                (data['Customer'].astype(str).str.strip().isin(['', 'nan', 'Customer', 'C1']) == False) &
                (data['Line'].astype(str).str.strip() != 'C1')
            ].reset_index(drop=True)
            print(f"Master Plan rows: {len(data)}")
            for _, row in data.iterrows():
                try:
                    customer = str(row.get('Customer', '')).strip()
                    style = str(row.get('Style', '')).strip()
                    color = str(row.get('Color', '')).strip()
                    season = str(row.get('Season', '')).strip()
                    drop = str(row.get('Drop', '')).strip()
                    ship_date = xl_to_date(row.get('Ship Date'))
                    qty = row.get('Qty', 0)
                    main_val = row.get('Main', '')
                    if hasattr(main_val, 'iloc'):
                        main_val = main_val.iloc[0]
                    fabric_main = str(main_val).strip().upper().split('\n')[0].strip()
                    fabric_codes = [c for c in [fabric_main] if c and c != 'NAN' and c != 'OK' and not c.replace('.', '').isdigit()]
                    if not ship_date or not customer or customer == 'nan':
                        continue
                    if ship_date < JUNE_2026:
                        continue
                    master_orders.append({
                        'customer': customer, 'style': style, 'color': color,
                        'season': season, 'drop': drop,
                        'ship_date': ship_date, 'ship_date_str': fmt_date(ship_date),
                        'days_to_ship': (ship_date - TODAY).days,
                        'qty_pcs': qty,
                        'docket_plan': xl_to_date(row.get('Docket Plan')),
                        'actual_docket': xl_to_date(row.get('Actual Docket')),
                        'fabric_codes': fabric_codes,
                        'ck': f"{customer}|{style}|{color}".upper()
                    })
                except:
                    continue
            break
    print(f"Master Plan orders (June+): {len(master_orders)}")

    # 2. ERP
    erp_data = {}
    erp = dfs.get('ERP')
    if erp:
        for sname, df in erp['sheets'].items():
            for i in range(min(3, len(df))):
                vals = [str(v).lower() for v in df.iloc[i].values]
                if any('fabric fast' in v or 'fast code' in v for v in vals):
                    data = df.iloc[i+1:].reset_index(drop=True)
                    data.columns = [str(c).strip() for c in df.iloc[i].values]
                    for _, row in data.iterrows():
                        try:
                            code = str(row.get('Fabric FAST Code', '')).strip().upper()
                            if not code or code == 'NAN':
                                continue
                            confirm_date = row.get('Confirm Expect Stock in Date')
                            revised_date = row.get('Revised Confirmed Stock in Date')
                            erp_data[code] = {
                                'status': str(row.get('Master PO Status', '')).strip(),
                                'actual_qty': row.get('Actual Order Qty', 0),
                                'confirm_date': pd.Timestamp(confirm_date).to_pydatetime() if pd.notna(confirm_date) else None,
                                'revised_date': pd.Timestamp(revised_date).to_pydatetime() if pd.notna(revised_date) else None,
                            }
                        except:
                            continue
                    break
            break
    print(f"ERP: {len(erp_data)} codes")

    # 3. DELIVERY PLAN
    delivery_data = {}
    delivery = dfs.get('Delivery & QA')
    if delivery:
        for sname, df in delivery['sheets'].items():
            if sname.lower() == 'summary':
                continue
            for i in range(min(8, len(df))):
                vals = [str(v).lower() for v in df.iloc[i].values]
                if any('fast code' in v for v in vals):
                    data = df.iloc[i+1:].reset_index(drop=True)
                    data.columns = [str(c).strip() for c in df.iloc[i].values]
                    for _, row in data.iterrows():
                        try:
                            code = str(row.get('fast code', '')).strip().upper()
                            if not code or code == 'NAN':
                                continue
                            ready = xl_to_date(row.get('READY TO SHIP DATE')) or (pd.Timestamp(row.get('READY TO SHIP DATE')).to_pydatetime() if pd.notna(row.get('READY TO SHIP DATE')) else None)
                            qty = row.get("DELIVER Q'TY", 0)
                            if code not in delivery_data:
                                delivery_data[code] = {'ready_date': ready, 'qty': qty}
                            else:
                                try:
                                    delivery_data[code]['qty'] = float(delivery_data[code]['qty'] or 0) + float(qty or 0)
                                except:
                                    pass
                        except:
                            continue
                    break
            print(f"  Sheet '{sname}': {len(delivery_data)} codes cumules")
    print(f"Delivery: {len(delivery_data)} codes")

    # 4. FABRIC CHECKEDIN
    fabric_checkedin = {}
    fabric = dfs.get('Fabric')
    if fabric:
        for sname, df in fabric['sheets'].items():
            if 'checkedin' in sname.lower():
                for i in range(min(6, len(df))):
                    vals = [str(v).lower() for v in df.iloc[i].values]
                    if any('fabric code' in v for v in vals):
                        data = df.iloc[i+1:].reset_index(drop=True)
                        data.columns = [str(c).strip() for c in df.iloc[i].values]
                        for _, row in data.iterrows():
                            try:
                                code = str(row.get('Fabric Code', '')).strip().upper()
                                if not code or code == 'NAN':
                                    continue
                                date_val = xl_to_date(row.get('DATE'))
                                qty = row.get("Q'TY (MET)", 0)
                                if code not in fabric_checkedin:
                                    fabric_checkedin[code] = {'date': date_val, 'qty': qty}
                                else:
                                    try:
                                        fabric_checkedin[code]['qty'] = float(fabric_checkedin[code]['qty'] or 0) + float(qty or 0)
                                    except:
                                        pass
                            except:
                                continue
                        break
    print(f"Fabric CheckedIn: {len(fabric_checkedin)} codes")

    # 5. DAILY REPORT (WH)
    wh_data = {}
    daily = dfs.get('Daily Report')
    if daily:
        print(f"Daily sheets: {list(daily['sheets'].keys())}")
        for sname, df in daily['sheets'].items():
            if 'daily' in sname.lower():
                header_row = 3
                data = df.iloc[header_row+1:].reset_index(drop=True)
                for _, row in data.iterrows():
                    try:
                        code = str(row.iloc[4]).strip().upper()
                        if not code or code == 'NAN':
                            continue
                        date_val = xl_to_date(row.iloc[3])
                        qty = row.iloc[9]
                        if code not in wh_data:
                            wh_data[code] = {'date': date_val, 'qty': qty}
                        else:
                            try:
                                wh_data[code]['qty'] = float(wh_data[code]['qty'] or 0) + float(qty or 0)
                            except:
                                pass
                    except:
                        continue
    print(f"WH received: {len(wh_data)} codes")

    # 6. MERCHANDISE
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
                                ck = f"{str(r.get('Customer','')).strip().upper()}|{str(r.get('Style','')).strip().upper()}|{str(r.get('Color','')).strip().upper()}"
                                plan_fab = str(r.get('PLAN FULL FABRIC', '')).strip()
                                stock_out = xl_to_date(r.get('Stock Out Date'))
                                shipment_date = xl_to_date(r.get('SHIPMENT DATE'))
                                qty = r.get('Qty', 0)
                                merch_status[ck] = {
                                    'mer_released': plan_fab == 'DONE',
                                    'stock_out': stock_out,
                                    'release_date': stock_out,
                                    'qty': qty,
                                    'shipment_date': shipment_date
                                }
                            except:
                                continue
                        break
    print(f"Merchandise: {len(merch_status)} entries")

    # 7. CROSS-CHECK
    PRIORITY_CUSTOMERS = {'ALD', 'GOLF WANG', 'RODD & GUNN', 'CORTEIZ', 'STUSSY', 'RAPHA'}
    SKIP = ['Done', 'In Progress', 'In Checking Process']

    for order in master_orders:
        try:
            days = order['days_to_ship']
            ck = order['ck']
            customer = order['customer']
            style = order['style']
            color = order['color']
            fabric_codes = order['fabric_codes']
            docket_plan = order['docket_plan']
            actual_docket = order['actual_docket']
            stats['total'] += 1

            issues = []
            root_causes = []
            blocking_dept = None
            merch_d = merch_status.get(ck, {})
            erp_info = {}
            delivery_info = {}
            qa_info = {}
            wh_info = {}

            for code in fabric_codes:
                e = erp_data.get(code, {})
                if e: erp_info = e
                d = delivery_data.get(code, {})
                if d: delivery_info = d
                q = fabric_checkedin.get(code, {})
                if q: qa_info = q
                w = wh_data.get(code, {})
                if w: wh_info = w

                erp_st = erp_info.get('status', '')
                revised_date = erp_info.get('revised_date')
                confirm_date = erp_info.get('confirm_date')

                if erp_st == 'Over-due':
                    ref_date = revised_date or confirm_date
                    days_od = (TODAY - ref_date).days if ref_date else '?'
                    issues.append(f'Nha cung cap TRE {days_od} ngay (ERP Over-due) - ma: {code}')
                    root_causes.append('erp_overdue')
                    blocking_dept = 'Purchasing'
                    break
                if erp_st == 'On-due in next 10 days':
                    issues.append(f'Vai sap den han trong 10 ngay - ma: {code}')
                    root_causes.append('erp_ondue')
                    blocking_dept = 'Purchasing'
                    break
                if not delivery_info and erp_st not in SKIP:
                    issues.append(f'Chua co trong Delivery Plan - ma: {code}')
                    root_causes.append('not_in_delivery')
                    blocking_dept = 'Purchasing'
                    break
                if not qa_info and erp_st not in SKIP and erp_st != '':
                    issues.append(f'Chua duoc kiem tra QA - ma: {code}')
                    root_causes.append('not_inspected')
                    blocking_dept = 'QA'
                    break
                if not wh_info and erp_st not in SKIP and erp_st != '':
                    issues.append(f'Chua nhan vao kho - ma: {code}')
                    root_causes.append('not_in_wh')
                    blocking_dept = 'Warehouse'
                    break

            if not issues:
                erp_st_final = erp_info.get('status', '')
                if not merch_d.get('mer_released', False) and bool(merch_d) and erp_st_final == 'Done':
                    issues.append('Vai san sang nhung MER chua release')
                    root_causes.append('mer_not_released')
                    blocking_dept = 'Merchandising'

            if docket_plan and actual_docket:
                delay = (actual_docket - docket_plan).days
                if delay > 3:
                    issues.append(f'Docket tre {delay} ngay')
                    root_causes.append('production_delay')
                    if not blocking_dept:
                        blocking_dept = 'Production'

            if not issues:
                stats['ok'] += 1
                if stats['ok'] <= 3:
                    print(f"DEBUG OK: {customer} | fabric={fabric_codes} | erp='{erp_info.get('status', '')}'")
                continue

            if stats['total'] <= 5:
                print(f"ISSUE {stats['total']}: {customer}|{style} fabric={fabric_codes} erp='{erp_info.get('status','')}' -> {issues}")

            try:
                if days < 0:
                    level = 'CRITICAL'
                    prefix = f'Tre {abs(int(days))} ngay - '
                elif days <= 7:
                    level = 'CRITICAL'
                    prefix = f'Con {int(days)} ngay - '
                elif days <= 14:
                    level = 'CRITICAL'
                    prefix = ''
                elif days <= 28:
                    level = 'RISK'
                    prefix = ''
                else:
                    level = 'WATCH'
                    prefix = ''
            except:
                level = 'WATCH'
                prefix = ''

            stats[level.lower()] += 1

            erp_qty = float(erp_info.get('actual_qty', 0) or 0)
            qa_qty = float(qa_info.get('qty', 0) or 0)
            wh_qty = float(wh_info.get('qty', 0) or 0)
            del_qty = float(delivery_info.get('qty', 0) or 0)
            master_qty = float(order.get('qty_pcs', 0) or 0)

            pct_delivery = round(del_qty / erp_qty * 100) if erp_qty > 0 else 0
            pct_qa = round(qa_qty / erp_qty * 100) if erp_qty > 0 else 0
            pct_wh = round(wh_qty / erp_qty * 100) if erp_qty > 0 else 0

            try:
                anomalies.append({
                    'level': level,
                    'department': blocking_dept or 'Unknown',
                    'customer': customer,
                    'style': style,
                    'color': color,
                    'season': order.get('season', ''),
                    'drop': order.get('drop', ''),
                    'shipDate': order['ship_date_str'],
                    'daysToShip': int(days) if days == days else 0,
                    'qtyPcs': int(master_qty) if master_qty and not (isinstance(master_qty, float) and master_qty != master_qty) else 0,
                    'issue': prefix + ' | '.join(issues),
                    'action': build_action(root_causes, customer, style, days, blocking_dept),
                    'rootCauses': root_causes,
                    'fabricCodes': fabric_codes,
                    'erp_actual_qty': int(erp_qty) if erp_qty and not (isinstance(erp_qty, float) and erp_qty != erp_qty) else 0,
                    'erp_confirm_date': fmt_date(erp_info.get('confirm_date')),
                    'erp_revised_date': fmt_date(erp_info.get('revised_date')),
                    'erp_status': erp_info.get('status', ''),
                    'erp_pct_arrived': pct_delivery,
                    'delivery_ready_date': fmt_date(delivery_info.get('ready_date')),
                    'delivery_qty': int(del_qty) if del_qty and not (isinstance(del_qty, float) and del_qty != del_qty) else 0,
                    'delivery_pct': pct_delivery,
                    'qa_date': fmt_date(qa_info.get('date')),
                    'qa_qty': round(qa_qty, 1) if qa_qty else 0,
                    'qa_pct': pct_qa,
                    'wh_date': fmt_date(wh_info.get('date')),
                    'wh_qty': round(wh_qty, 1) if wh_qty else 0,
                    'wh_pct': pct_wh,
                    'merch_release_date': fmt_date(merch_d.get('release_date')),
                    'merch_qty': 0,
                    'merch_status': 'DONE' if merch_d.get('mer_released') else ('PENDING' if ck in merch_status else 'N/A'),
                })
            except Exception as append_err:
                print(f"APPEND ERROR: {append_err} | customer={customer} style={style}")

        except Exception as order_err:
            print(f"ORDER ERROR: {order_err} | {order.get('customer','')} {order.get('style','')}")

        stats[level.lower()] += 1

        erp_qty = float(erp_info.get('actual_qty', 0) or 0)
        qa_qty = float(qa_info.get('qty', 0) or 0)
        wh_qty = float(wh_info.get('qty', 0) or 0)
        del_qty = float(delivery_info.get('qty', 0) or 0)
        master_qty = float(order.get('qty_pcs', 0) or 0)

        pct_delivery = round(del_qty / erp_qty * 100) if erp_qty > 0 and del_qty == del_qty else 0
        pct_qa = round(qa_qty / erp_qty * 100) if erp_qty > 0 and qa_qty == qa_qty else 0
        pct_wh = round(wh_qty / erp_qty * 100) if erp_qty > 0 and wh_qty == wh_qty else 0

        try:
            anomalies.append({
                'level': level,
                'department': blocking_dept or 'Unknown',
                'customer': customer,
                'style': style,
                'color': color,
                'season': order.get('season', ''),
                'drop': order.get('drop', ''),
                'shipDate': order['ship_date_str'],
                'daysToShip': int(days) if days == days else 0,
                'qtyPcs': int(master_qty) if master_qty and not (isinstance(master_qty, float) and master_qty != master_qty) else 0,
                'issue': prefix + ' | '.join(issues),
                'action': build_action(root_causes, customer, style, days, blocking_dept),
                'rootCauses': root_causes,
                'fabricCodes': fabric_codes,
                'erp_actual_qty': int(erp_qty) if erp_qty and not (isinstance(erp_qty, float) and erp_qty != erp_qty) else 0,
                'erp_confirm_date': fmt_date(erp_info.get('confirm_date')),
                'erp_revised_date': fmt_date(erp_info.get('revised_date')),
                'erp_status': erp_info.get('status', ''),
                'erp_pct_arrived': pct_delivery,
                'delivery_ready_date': fmt_date(delivery_info.get('ready_date')),
                'delivery_qty': int(del_qty) if del_qty and not (isinstance(del_qty, float) and del_qty != del_qty) else 0,
                'delivery_pct': pct_delivery,
                'qa_date': fmt_date(qa_info.get('date')),
                'qa_qty': round(qa_qty, 1) if qa_qty else 0,
                'qa_pct': pct_qa,
                'wh_date': fmt_date(wh_info.get('date')),
                'wh_qty': round(wh_qty, 1) if wh_qty else 0,
                'wh_pct': pct_wh,
                'merch_release_date': fmt_date(merch_d.get('release_date')),
                'merch_qty': 0,
                'merch_status': 'DONE' if merch_d.get('mer_released') else ('PENDING' if ck in merch_status else 'N/A'),
            })
        except Exception as append_err:
            print(f"APPEND ERROR: {append_err} | customer={customer} style={style}")

    level_order = {'CRITICAL': 0, 'RISK': 1, 'WATCH': 2, 'OK': 3}
    try:
        anomalies.sort(key=lambda x: (
            0 if x['customer'].upper() in PRIORITY_CUSTOMERS else 1,
            level_order.get(x['level'], 9),
            x.get('daysToShip', 999) if x.get('daysToShip', 999) == x.get('daysToShip', 999) else 999
        ))
    except Exception as sort_err:
        print(f"SORT ERROR: {sort_err}")

    dept_status = []
    for dept in ['Purchasing', 'QA', 'Warehouse', 'Merchandising', 'Production']:
        dept_alerts = [a for a in anomalies if a['department'] == dept]
        critical_count = len([a for a in dept_alerts if a['level'] == 'CRITICAL'])
        if not dept_alerts:
            continue
        dept_status.append({
            'name': dept,
            'status': 'CRITICAL' if critical_count > 0 else 'WARNING',
            'summary': f"{len(dept_alerts)} van de, {critical_count} khan cap"
        })

    print(f"Done: {len(anomalies)} | CRITICAL: {stats.get('critical',0)} | RISK: {stats.get('risk',0)} | WATCH: {stats.get('watch',0)} | OK: {stats['ok']}")

    return {
        'summary': {
            'totalOrders': stats['total'],
            'critical': len([a for a in anomalies if a['level'] == 'CRITICAL']),
            'risk': len([a for a in anomalies if a['level'] == 'RISK']),
            'watch': len([a for a in anomalies if a['level'] == 'WATCH']),
            'onTrack': stats['ok'],
            'depsReceived': len(files_data)
        },
        'alerts': anomalies[:200],
        'departmentStatus': dept_status,
        'recommendations': build_recommendations(anomalies),
        'globalSummary': build_global_summary(anomalies, stats, len(files_data), TODAY)
    }

def build_action(root_causes, customer, style, days, dept):
    if 'erp_overdue' in root_causes:
        return f'Lien he ngay Purchasing - nha cung cap tre giao vai cho {customer}/{style}'
    if 'erp_ondue' in root_causes:
        return f'Theo doi sat Purchasing - vai {customer}/{style} sap den han'
    if 'not_in_delivery' in root_causes:
        return f'Yeu cau Purchasing them vai {customer}/{style} vao Delivery Plan'
    if 'not_inspected' in root_causes:
        return f'Yeu cau QA kiem tra vai {customer}/{style} ngay'
    if 'not_in_wh' in root_causes:
        return f'Kiem tra Warehouse - vai {customer}/{style} chua nhan vao kho'
    if 'mer_not_released' in root_causes:
        return f'Yeu cau Merchandising release MER ngay cho {customer}/{style}'
    if 'production_delay' in root_causes:
        return f'Theo doi tien do san xuat cho {customer}/{style}'
    return f'Theo doi {customer}/{style}'

def build_recommendations(anomalies):
    recs = []
    overdue = [a for a in anomalies if a.get('daysToShip', 0) < 0]
    erp_od = [a for a in anomalies if 'erp_overdue' in a.get('rootCauses', [])]
    not_del = [a for a in anomalies if 'not_in_delivery' in a.get('rootCauses', [])]
    not_insp = [a for a in anomalies if 'not_inspected' in a.get('rootCauses', [])]
    not_wh = [a for a in anomalies if 'not_in_wh' in a.get('rootCauses', [])]
    mer = [a for a in anomalies if 'mer_not_released' in a.get('rootCauses', [])]

    if overdue:
        recs.append(f"Xu ly khan cap {len(overdue)} don hang da TRE ngay xuat")
    if erp_od:
        recs.append(f"Purchasing: {len(erp_od)} nha cung cap tre - can escalate ngay")
    if not_del:
        recs.append(f"Purchasing: them {len(not_del)} ma vai vao Delivery Plan")
    if not_insp:
        recs.append(f"QA: kiem tra {len(not_insp)} lo vai dang cho inspection")
    if not_wh:
        recs.append(f"Warehouse: xac nhan {len(not_wh)} lo vai chua nhan")
    if mer:
        recs.append(f"Merchandising: release MER cho {len(mer)} don hang")

    return recs[:5] if recs else ["Tat ca don hang thang 6+ dang trong tam kiem soat"]

def build_global_summary(anomalies, stats, dep_count, today):
    critical = len([a for a in anomalies if a['level'] == 'CRITICAL'])
    risk = len([a for a in anomalies if a['level'] == 'RISK'])
    overdue = len([a for a in anomalies if a.get('daysToShip', 0) < 0])
    date_str = today.strftime('%d/%m/%Y')
    if not anomalies:
        return f"Ngay {date_str}: Phan tich {dep_count} bao cao. Tat ca don hang thang 6+ dang trong tam kiem soat."
    summary = f"Ngay {date_str}: Phan tich {dep_count} bao cao, tap trung don hang thang 6+. "
    if overdue > 0:
        summary += f"{overdue} don hang da TRE. "
    if critical > 0:
        summary += f"{critical} don hang KHAN CAP. "
    if risk > 0:
        summary += f"{risk} don hang RUI RO. "
    summary += "Xem chi tiet ben duoi."
    return summary