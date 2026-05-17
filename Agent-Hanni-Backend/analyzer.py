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
        return None

def analyze_files(files_data):
    """
    files_data: list of { filename, content (bytes) }
    Returns: dict with anomalies, summary, department_status
    """
    TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
    dfs = {}
    dept_info = {}
    errors = []

    # === LOAD ALL FILES ===
    for f in files_data:
        fname = f['filename']
        dept = detect_department(fname)
        xl = read_excel_safe(f['content'], fname)
        if xl is None:
            errors.append(f"Không đọc được file: {fname}")
            continue
        
        sheets = {}
        for sheet in xl.sheet_names[:8]:
            try:
                engine = 'pyxlsb' if fname.lower().endswith('.xlsb') else None
                if engine:
                    df = pd.read_excel(xl, sheet_name=sheet, header=None, engine=engine)
                else:
                    df = pd.read_excel(xl, sheet_name=sheet, header=None)
                sheets[sheet] = df
            except:
                pass
        
        dfs[dept] = {'filename': fname, 'sheets': sheets}
        dept_info[dept] = {'filename': fname, 'sheet_count': len(sheets), 'status': 'OK', 'issues': []}

    anomalies = []
    stats = {'total': 0, 'critical': 0, 'risk': 0, 'watch': 0, 'ok': 0}

    # === ANALYZE MERCHANDISE ===
    merch = dfs.get('Merchandise')
    if merch:
        # Find Plan sheet
        plan_df = None
        tracking_df = None
        for sname, df in merch['sheets'].items():
            if 'plan' in sname.lower() and plan_df is None:
                # Find header row
                for i, row in df.iterrows():
                    vals = [str(v).lower() for v in row.values]
                    if 'customer' in vals and 'style' in vals:
                        plan_df = df.iloc[i+1:].reset_index(drop=True)
                        plan_df.columns = df.iloc[i].values
                        plan_df = plan_df.dropna(subset=[df.iloc[i].values[0]])
                        break
            if 'tracking' in sname.lower() and tracking_df is None:
                for i, row in df.iterrows():
                    vals = [str(v).lower() for v in row.values]
                    if 'customer' in vals:
                        tracking_df = df.iloc[i+1:].reset_index(drop=True)
                        tracking_df.columns = df.iloc[i].values
                        break

        if tracking_df is not None:
            for _, row in tracking_df.iterrows():
                try:
                    plan_fab = str(row.get('PLAN FULL FABRIC', ''))
                    ship_date_raw = row.get('SHIPMENT DATE')
                    stock_out = row.get('Stock Out Date')
                    customer = str(row.get('Customer', ''))
                    style = str(row.get('Style', ''))
                    color = str(row.get('Color', ''))

                    ship_date = xl_to_date(ship_date_raw)
                    if not ship_date or not customer or customer == 'nan':
                        continue

                    days = (ship_date - TODAY).days
                    stats['total'] += 1
                    issues = []

                    # Fabric not released
                    if plan_fab not in ('DONE', 'nan', '') and plan_fab != 'DONE':
                        try:
                            planned = xl_to_date(float(plan_fab))
                            planned_str = fmt_date(planned) if planned else plan_fab
                        except:
                            planned_str = plan_fab
                        issues.append(f'Vải chưa được release (dự kiến: {planned_str})')

                    # Fabric not received
                    stock_dt = xl_to_date(stock_out)
                    if not stock_dt:
                        issues.append('Vải chưa về xưởng')

                    if issues:
                        level = 'CRITICAL' if days <= 14 else ('RISK' if days <= 28 else 'WATCH')
                        if days < 0:
                            level = 'CRITICAL'
                            issues.insert(0, f'Đã trễ hạn {abs(days)} ngày!')
                        
                        stats[level.lower()] += 1
                        anomalies.append({
                            'level': level,
                            'department': 'Merchandising',
                            'customer': customer,
                            'style': style,
                            'color': color,
                            'shipDate': fmt_date(ship_date),
                            'daysToShip': int(days),
                            'issue': ' | '.join(issues),
                            'action': build_action(issues, customer, style, days)
                        })
                    else:
                        stats['ok'] += 1
                except:
                    continue

            if anomalies:
                dept_info['Merchandise']['status'] = 'WARNING'
                dept_info['Merchandise']['issues'].append(f"{len([a for a in anomalies if a['department']=='Merchandising'])} vấn đề phát hiện")

    # === ANALYZE DELIVERY PLAN ===
    delivery = dfs.get('Delivery & QA')
    if delivery:
        for sname, df in delivery['sheets'].items():
            if 'delivery' in sname.lower() or 'uninspect' in sname.lower():
                for i, row in df.iterrows():
                    vals = [str(v).lower() for v in row.values]
                    if 'customer' in vals and 'supplier' in vals:
                        data = df.iloc[i+2:].reset_index(drop=True)
                        data.columns = df.iloc[i+1].values
                        data = data[data['Customer'].notna() & (data['Customer'] != 'Unplan')]
                        
                        not_received = data[
                            data['RECEIVED/INSPECTED'].isna() | 
                            (data['RECEIVED/INSPECTED'].astype(str).str.upper() == 'NAN')
                        ]
                        
                        for _, row2 in not_received.iterrows():
                            try:
                                ship_raw = row2.get('READY TO SHIP DATE')
                                if ship_raw == 'Unplan' or pd.isna(ship_raw):
                                    continue
                                ship_date = pd.to_datetime(ship_raw, errors='coerce')
                                if pd.isna(ship_date):
                                    continue
                                ship_date = ship_date.to_pydatetime()
                                days = (ship_date - TODAY).days
                                customer = str(row2.get('Customer', ''))
                                fast_code = str(row2.get('fast code', ''))
                                
                                if days <= 35 and customer and customer != 'nan':
                                    level = 'CRITICAL' if days <= 14 else 'RISK'
                                    anomalies.append({
                                        'level': level,
                                        'department': 'QA / Delivery',
                                        'customer': customer,
                                        'style': fast_code,
                                        'color': str(row2.get('Color', '')),
                                        'shipDate': ship_date.strftime('%d/%m/%Y'),
                                        'daysToShip': int(days),
                                        'issue': 'Vải chưa được kiểm tra QA và chưa nhận vào kho',
                                        'action': f'Liên hệ ngay bộ phận QA để kiểm tra vải cho {customer} - {fast_code}'
                                    })
                                    stats[level.lower()] += 1
                            except:
                                continue
                        break

    # === ANALYZE MASTER PLAN ===
    master = dfs.get('Master Plan')
    if master:
        for sname, df in master['sheets'].items():
            if sname.lower() == 'plan':
                for i, row in df.iterrows():
                    vals = [str(v).lower() for v in row.values]
                    if 'line' in vals and 'customer' in vals and 'ship date' in vals:
                        data = df.iloc[i+2:].reset_index(drop=True)
                        data.columns = df.iloc[i].values
                        data = data[data['Line'] != 'C1'].dropna(subset=['Customer', 'Style'])
                        
                        for _, row2 in data.iterrows():
                            try:
                                ship_raw = row2.get('Ship Date')
                                ship_date = xl_to_date(ship_raw)
                                if not ship_date:
                                    continue
                                days = (ship_date - TODAY).days
                                if days > 60 or days < -30:
                                    continue
                                
                                docket_plan = xl_to_date(row2.get('Docket Plan'))
                                actual_docket = xl_to_date(row2.get('Actual Docket'))
                                
                                if docket_plan and actual_docket:
                                    delay = (actual_docket - docket_plan).days
                                    if delay > 3 and days <= 28:
                                        customer = str(row2.get('Customer', ''))
                                        style = str(row2.get('Style', ''))
                                        anomalies.append({
                                            'level': 'RISK' if days > 14 else 'CRITICAL',
                                            'department': 'Production',
                                            'customer': customer,
                                            'style': style,
                                            'color': str(row2.get('Color', '')),
                                            'shipDate': fmt_date(ship_date),
                                            'daysToShip': int(days),
                                            'issue': f'Docket thực tế trễ {delay} ngày so với kế hoạch (KH: {fmt_date(docket_plan)}, TT: {fmt_date(actual_docket)})',
                                            'action': f'Kiểm tra tiến độ sản xuất line {row2.get("Line","")} cho {customer} - {style}'
                                        })
                                        stats['risk'] += 1
                            except:
                                continue
                break

    # Sort by urgency
    level_order = {'CRITICAL': 0, 'RISK': 1, 'WATCH': 2, 'OK': 3}
    anomalies.sort(key=lambda x: (level_order.get(x['level'], 9), x.get('daysToShip', 999)))

    # Department status summary
    dept_status = []
    for dept, info in dept_info.items():
        dept_issues = [a for a in anomalies if a['department'].lower() in dept.lower() or dept.lower() in a['department'].lower()]
        critical_count = len([a for a in dept_issues if a['level'] == 'CRITICAL'])
        status = 'CRITICAL' if critical_count > 0 else ('WARNING' if dept_issues else 'OK')
        dept_status.append({
            'name': dept,
            'status': status,
            'summary': f"{len(dept_issues)} vấn đề, {critical_count} khẩn cấp" if dept_issues else "Không có vấn đề phát hiện"
        })

    return {
        'summary': {
            'totalOrders': stats['total'],
            'critical': len([a for a in anomalies if a['level'] == 'CRITICAL']),
            'risk': len([a for a in anomalies if a['level'] == 'RISK']),
            'watch': len([a for a in anomalies if a['level'] == 'WATCH']),
            'onTrack': stats['ok'],
            'depsReceived': len(files_data),
            'errors': errors
        },
        'alerts': anomalies[:100],  # max 100 alerts
        'departmentStatus': dept_status,
        'recommendations': build_recommendations(anomalies, TODAY),
        'globalSummary': build_global_summary(anomalies, stats, len(files_data), TODAY)
    }

def build_action(issues, customer, style, days):
    if days < 0:
        return f'KHẨN CẤP: Liên hệ ngay {customer} để thông báo trễ hàng và tìm giải pháp'
    if 'chưa về xưởng' in ' '.join(issues):
        return f'Kiểm tra với bộ phận Purchasing tình trạng vải cho {customer} - {style}'
    if 'chưa được release' in ' '.join(issues):
        return f'Yêu cầu Merchandising release vải ngay cho {customer} - {style}'
    return f'Theo dõi sát tình trạng {customer} - {style}'

def build_recommendations(anomalies, today):
    recs = []
    critical = [a for a in anomalies if a['level'] == 'CRITICAL']
    if critical:
        customers = list(set([a['customer'] for a in critical[:3]]))
        recs.append(f"Ưu tiên xử lý ngay {len(critical)} đơn hàng khẩn cấp: {', '.join(customers)}")
    
    overdue = [a for a in anomalies if a.get('daysToShip', 0) < 0]
    if overdue:
        recs.append(f"Có {len(overdue)} đơn hàng đã trễ ngày xuất — cần liên hệ khách hàng ngay hôm nay")
    
    no_fabric = [a for a in anomalies if 'chưa về xưởng' in a.get('issue', '')]
    if no_fabric:
        recs.append(f"Theo dõi {len(no_fabric)} đơn hàng chưa nhận vải về xưởng với bộ phận Purchasing")
    
    no_qa = [a for a in anomalies if 'QA' in a.get('issue', '')]
    if no_qa:
        recs.append(f"Nhắc bộ phận QA kiểm tra {len(no_qa)} lô vải đang chờ")
    
    if not recs:
        recs.append("Tất cả đơn hàng đang trong tầm kiểm soát — tiếp tục theo dõi theo lịch")
    
    return recs[:5]

def build_global_summary(anomalies, stats, dep_count, today):
    critical = len([a for a in anomalies if a['level'] == 'CRITICAL'])
    risk = len([a for a in anomalies if a['level'] == 'RISK'])
    overdue = len([a for a in anomalies if a.get('daysToShip', 0) < 0])
    
    date_str = today.strftime('%d/%m/%Y')
    
    if critical == 0 and risk == 0:
        return f"Ngày {date_str}: Đã phân tích {dep_count} báo cáo từ các bộ phận. Tình hình chuỗi cung ứng ổn định, không có vấn đề nghiêm trọng được phát hiện."
    
    summary = f"Ngày {date_str}: Đã phân tích {dep_count} báo cáo. "
    if overdue > 0:
        summary += f"Có {overdue} đơn hàng đã TRỄ ngày xuất — cần xử lý khẩn cấp. "
    if critical > 0:
        summary += f"{critical} đơn hàng ở mức KHẨN CẤP (dưới 14 ngày). "
    if risk > 0:
        summary += f"{risk} đơn hàng có RỦI RO (15-28 ngày). "
    summary += "Vui lòng xem chi tiết bên dưới và thực hiện các hành động được đề xuất."
    
    return summary
