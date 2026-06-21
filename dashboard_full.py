import random
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
from faker import Faker
import json

# ─────────────────────────────────────────────
# 1. تولید داده‌ها
# ─────────────────────────────────────────────
np.random.seed(42)
random.seed(42)
fake = Faker()
Faker.seed(42)

N_SALES = 5000
START_DATE = datetime(2024, 1, 1)
END_DATE = datetime(2024, 12, 31)

REGIONS = ["Tehran", "Isfahan", "Mashhad", "Shiraz", "Tabriz"]
CHANNELS = ["Modern Trade", "Traditional Trade", "E-Commerce", "HoReCa"]
CATEGORIES = ["Dairy", "Beverages", "Snacks", "Personal Care", "Home Care"]
BRANDS = ["BrandA", "BrandB", "BrandC", "BrandD"]
CUSTOMERS = [f"CUST_{i:04d}" for i in range(1, 201)]
PRODUCTS = [f"SKU_{i:03d}" for i in range(1, 51)]
SALESREPS = [f"REP_{i:02d}" for i in range(1, 16)]

def rand_dates(n, start=START_DATE, end=END_DATE):
    delta = (end - start).days
    return [start + timedelta(days=int(d)) for d in np.random.randint(0, delta + 1, n)]

def make_sales():
    print("📊 در حال تولید داده‌های فروش...")
    dates = rand_dates(N_SALES)
    customers = np.random.choice(CUSTOMERS, N_SALES)
    products = np.random.choice(PRODUCTS, N_SALES)
    reps = np.random.choice(SALESREPS, N_SALES)
    regions = np.random.choice(REGIONS, N_SALES)
    channels = np.random.choice(CHANNELS, N_SALES)
    categories = np.random.choice(CATEGORIES, N_SALES)
    brands = np.random.choice(BRANDS, N_SALES)
    quantity = np.random.randint(1, 200, N_SALES)
    unit_price = np.round(np.random.uniform(5, 500, N_SALES), 2)
    discount_rate = np.round(np.random.uniform(0, 0.25, N_SALES), 4)
    gross_sales = np.round(quantity * unit_price, 2)
    discount_amt = np.round(gross_sales * discount_rate, 2)
    net_sales = np.round(gross_sales - discount_amt, 2)

    df = pd.DataFrame({
        "Date": dates,
        "Month": [d.month for d in dates],
        "Quarter": [f"Q{((d.month-1)//3)+1}" for d in dates],
        "Customer_ID": customers,
        "Product_ID": products,
        "SalesRep_ID": reps,
        "Region": regions,
        "Channel": channels,
        "Category": categories,
        "Brand": brands,
        "Quantity": quantity,
        "Unit_Price": unit_price,
        "Discount_Rate": discount_rate,
        "Discount_Amount": discount_amt,
        "Net_Sales": net_sales,
    })
    return df.sort_values("Date").reset_index(drop=True)

# ─────────────────────────────────────────────
# 2. محاسبه شاخص‌ها
# ─────────────────────────────────────────────
def calculate_kpis(df):
    total_revenue = df['Net_Sales'].sum()
    total_quantity = df['Quantity'].sum()
    avg_price = df['Unit_Price'].mean()
    total_discount = df['Discount_Amount'].sum()
    discount_rate = (total_discount / (total_revenue + total_discount) * 100) if (total_revenue + total_discount) > 0 else 0
    row_count = len(df)
    unique_customers = df['Customer_ID'].nunique()
    
    monthly = df.groupby('Month')['Net_Sales'].sum().reindex(range(1, 13), fill_value=0).tolist()
    months = ['فروردین', 'اردیبهشت', 'خرداد', 'تیر', 'مرداد', 'شهریور', 
              'مهر', 'آبان', 'آذر', 'دی', 'بهمن', 'اسفند']
    
    channel_data = df.groupby('Channel')['Net_Sales'].sum().to_dict()
    channel_labels = list(channel_data.keys())
    channel_values = list(channel_data.values())
    
    quarterly = df.groupby('Quarter')['Net_Sales'].sum().reindex(['Q1', 'Q2', 'Q3', 'Q4'], fill_value=0).tolist()
    
    category_data = df.groupby('Category')['Net_Sales'].sum().to_dict()
    category_labels = list(category_data.keys())
    category_values = list(category_data.values())
    
    region_data = df.groupby('Region')['Net_Sales'].sum().to_dict()
    region_labels = list(region_data.keys())
    region_values = list(region_data.values())
    
    return {
        'total_revenue': total_revenue,
        'total_quantity': total_quantity,
        'avg_price': avg_price,
        'discount_rate': discount_rate,
        'total_discount': total_discount,
        'row_count': row_count,
        'unique_customers': unique_customers,
        'monthly': monthly,
        'months': months,
        'channel_labels': channel_labels,
        'channel_values': channel_values,
        'quarterly': quarterly,
        'category_labels': category_labels,
        'category_values': category_values,
        'region_labels': region_labels,
        'region_values': region_values,
        'recent': df.head(10)[['Date', 'Customer_ID', 'Product_ID', 'Channel', 'Net_Sales']].copy()
    }

# ─────────────────────────────────────────────
# 3. ساخت HTML
# ─────────────────────────────────────────────
def build_pro_dashboard(kpis):
    k = kpis
    recent = k['recent']
    recent['Date'] = recent['Date'].dt.strftime('%Y-%m-%d')
    
    table_rows = ''
    for _, row in recent.iterrows():
        table_rows += f'<tr><td>{row["Date"]}</td><td>{row["Customer_ID"]}</td><td>{row["Product_ID"]}</td><td><span class="badge">{row["Channel"]}</span></td><td>{row["Net_Sales"]:,.0f}</td></tr>'
    
    html = f'''<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>داشبورد مدیریتی FMCG</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, sans-serif; }}
        body {{ background: #f0f2f5; padding: 20px; }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        .header {{ background: linear-gradient(135deg, #1a365d, #2b6cb0); color: white; padding: 25px 30px; 
                    border-radius: 15px; margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; }}
        .header h1 {{ font-size: 28px; font-weight: 300; }}
        .header p {{ opacity: 0.8; margin-top: 5px; }}
        .header-badge {{ background: rgba(255,255,255,0.2); padding: 8px 16px; border-radius: 20px; font-size: 13px; }}
        
        .nav-tabs {{ display: flex; gap: 4px; background: white; padding: 8px 12px; border-radius: 12px; 
                     box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 25px; flex-wrap: wrap; }}
        .nav-tab {{ padding: 8px 20px; border-radius: 8px; font-size: 13px; font-weight: 500; color: #4a5568; 
                   cursor: pointer; transition: all 0.2s; border: none; background: transparent; }}
        .nav-tab:hover {{ background: #edf2f7; }}
        .nav-tab.active {{ background: #2b6cb0; color: white; }}
        
        .page {{ display: none; animation: fadeIn 0.3s ease; }}
        .page.active {{ display: block; }}
        @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(10px); }} to {{ opacity: 1; transform: translateY(0); }} }}
        
        .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 25px; }}
        .card {{ background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); 
                transition: transform 0.2s; }}
        .card:hover {{ transform: translateY(-3px); box-shadow: 0 4px 20px rgba(0,0,0,0.1); }}
        .card .icon {{ font-size: 24px; margin-bottom: 8px; }}
        .card h3 {{ font-size: 12px; color: #718096; text-transform: uppercase; letter-spacing: 0.5px; }}
        .card .value {{ font-size: 26px; font-weight: 700; color: #2d3748; margin: 8px 0; }}
        .card .sub {{ font-size: 13px; }}
        .card .sub.green {{ color: #48bb78; }}
        .card .sub.red {{ color: #fc8181; }}
        .card .sub.orange {{ color: #ed8936; }}
        
        .alert-strip {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }}
        .alert-card {{ background: white; padding: 15px 20px; border-radius: 10px; border-right: 4px solid; 
                       box-shadow: 0 2px 8px rgba(0,0,0,0.04); display: flex; align-items: center; gap: 12px; }}
        .alert-card.danger {{ border-color: #fc8181; }}
        .alert-card.warning {{ border-color: #ed8936; }}
        .alert-card.success {{ border-color: #48bb78; }}
        .alert-icon {{ font-size: 20px; }}
        .alert-title {{ font-weight: 600; font-size: 14px; }}
        .alert-desc {{ font-size: 12px; color: #718096; }}
        
        .chart-grid {{ display: grid; grid-template-columns: 2fr 1fr; gap: 20px; margin-bottom: 25px; }}
        .chart-box {{ background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); }}
        .chart-box h3 {{ margin-bottom: 15px; color: #2d3748; font-size: 16px; }}
        .chart-box .subtitle {{ font-size: 12px; color: #718096; font-weight: normal; }}
        
        .table-box {{ background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th {{ text-align: right; padding: 12px 10px; background: #edf2f7; color: #4a5568; font-weight: 600; }}
        td {{ padding: 10px; border-bottom: 1px solid #e2e8f0; }}
        .badge {{ background: #ebf4ff; color: #2b6cb0; padding: 2px 10px; border-radius: 20px; font-size: 12px; }}
        
        @media (max-width: 768px) {{ 
            .chart-grid {{ grid-template-columns: 1fr; }}
            .alert-strip {{ grid-template-columns: 1fr; }}
            .header {{ flex-direction: column; text-align: center; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div>
                <h1>📊 داشبورد مدیریتی FMCG</h1>
                <p>گزارش تحلیلی فروش و توزیع • {k['row_count']:,} رکورد • {k['unique_customers']} مشتری فعال</p>
            </div>
            <div class="header-badge">📅 سال مالی ۱۴۰۳</div>
        </div>

        <div class="nav-tabs">
            <button class="nav-tab active" onclick="showPage('overview')">📊 Overview</button>
            <button class="nav-tab" onclick="showPage('pl')">📈 P&L Analysis</button>
            <button class="nav-tab" onclick="showPage('sales')">💰 Sales Drill</button>
            <button class="nav-tab" onclick="showPage('customers')">👥 Customers</button>
            <button class="nav-tab" onclick="showPage('inventory')">📦 Inventory</button>
        </div>

        <div class="page active" id="page-overview">
            <div class="alert-strip">
                <div class="alert-card danger">
                    <span class="alert-icon">🔴</span>
                    <div><div class="alert-title">هشدار: کاهش فروش در منطقه تهران</div>
                    <div class="alert-desc">فروش منطقه تهران ۸٪ نسبت به ماه قبل کاهش داشته است</div></div>
                </div>
                <div class="alert-card warning">
                    <span class="alert-icon">⚠️</span>
                    <div><div class="alert-title">تخفیف‌های غیرعادی</div>
                    <div class="alert-desc">نرخ تخفیف در کانال مدرن به ۱۸٪ رسیده است</div></div>
                </div>
                <div class="alert-card success">
                    <span class="alert-icon">✅</span>
                    <div><div class="alert-title">رشد فروش آنلاین</div>
                    <div class="alert-desc">فروش کانال E-Commerce ۳۲٪ رشد داشته است</div></div>
                </div>
            </div>

            <div class="grid">
                <div class="card"><div class="icon">💰</div><h3>فروش خالص</h3><div class="value">{k['total_revenue']:,.0f}</div><div class="sub green">▲ ۱۱.۴٪ نسبت به سال قبل</div></div>
                <div class="card"><div class="icon">📦</div><h3>حجم فروش</h3><div class="value">{k['total_quantity']:,.0f}</div><div class="sub green">▲ ۶.۸٪ رشد</div></div>
                <div class="card"><div class="icon">🏷️</div><h3>میانگین قیمت</h3><div class="value">{k['avg_price']:,.0f}</div><div class="sub green">▲ ۴.۶٪ افزایش</div></div>
                <div class="card"><div class="icon">🎯</div><h3>نرخ تخفیف</h3><div class="value">{k['discount_rate']:.1f}%</div><div class="sub orange">▲ ۲.۱٪ نسبت به هدف</div></div>
                <div class="card"><div class="icon">👥</div><h3>مشتریان فعال</h3><div class="value">{k['unique_customers']}</div><div class="sub green">▲ ۱۲ مشتری جدید</div></div>
            </div>

            <div class="chart-grid">
                <div class="chart-box">
                    <h3>📈 روند فروش ماهانه <span class="subtitle">(ریال)</span></h3>
                    <canvas id="monthlyChart" height="200"></canvas>
                </div>
                <div class="chart-box">
                    <h3>🍩 توزیع فروش بر اساس کانال</h3>
                    <canvas id="channelChart" height="200"></canvas>
                </div>
            </div>
            
            <div class="chart-grid" style="margin-top:20px;">
                <div class="chart-box">
                    <h3>📊 فروش فصلی <span class="subtitle">(ریال)</span></h3>
                    <canvas id="quarterlyChart" height="180"></canvas>
                </div>
                <div class="chart-box">
                    <h3>🏷️ فروش بر اساس دسته‌بندی</h3>
                    <canvas id="categoryChart" height="180"></canvas>
                </div>
            </div>

            <div class="table-box">
                <h3 style="margin-bottom: 15px;">🔍 ۱۰ معامله آخر</h3>
                <table>
                    <thead><tr><th>تاریخ</th><th>مشتری</th><th>محصول</th><th>کانال</th><th>مبلغ (ریال)</th></tr></thead>
                    <tbody>{table_rows}</tbody>
                </table>
            </div>
        </div>

        <div class="page" id="page-pl">
            <h2 style="margin-bottom:20px;">📈 تحلیل سود و زیان</h2>
            <div class="grid">
                <div class="card"><div class="icon">💰</div><h3>درآمد ناخالص</h3><div class="value">{k['total_revenue'] + k['total_discount']:,.0f}</div><div class="sub green">▲ ۸.۲٪</div></div>
                <div class="card"><div class="icon">📉</div><h3>تخفیف</h3><div class="value">{k['total_discount']:,.0f}</div><div class="sub red">▲ ۱۲.۴٪</div></div>
                <div class="card"><div class="icon">💰</div><h3>درآمد خالص</h3><div class="value">{k['total_revenue']:,.0f}</div><div class="sub green">▲ ۱۱.۴٪</div></div>
                <div class="card"><div class="icon">📊</div><h3>حاشیه سود ناخالص</h3><div class="value">۳۸.۲%</div><div class="sub green">▲ ۱۴۰ واحد</div></div>
            </div>
            <div class="chart-box">
                <h3>📊 مقایسه درآمد و تخفیف ماهانه</h3>
                <canvas id="plChart" height="200"></canvas>
            </div>
        </div>

        <div class="page" id="page-sales">
            <h2 style="margin-bottom:20px;">💰 تحلیل فروش</h2>
            <div class="chart-grid">
                <div class="chart-box">
                    <h3>🌍 فروش بر اساس منطقه</h3>
                    <canvas id="regionChart" height="200"></canvas>
                </div>
                <div class="chart-box">
                    <h3>📈 روند فروش کانال‌ها</h3>
                    <canvas id="channelTrendChart" height="200"></canvas>
                </div>
            </div>
        </div>

        <div class="page" id="page-customers">
            <h2 style="margin-bottom:20px;">👥 تحلیل مشتریان</h2>
            <div class="grid">
                <div class="card"><div class="icon">👥</div><h3>تعداد مشتریان</h3><div class="value">{k['unique_customers']}</div><div class="sub green">▲ ۶.۴٪ رشد</div></div>
                <div class="card"><div class="icon">🔄</div><h3>میانگین تعداد خرید</h3><div class="value">۴.۲</div><div class="sub green">▲ ۰.۸</div></div>
                <div class="card"><div class="icon">💰</div><h3>میانگین ارزش خرید</h3><div class="value">{(k['total_revenue'] / k['unique_customers']):,.0f}</div><div class="sub green">▲ ۵.۲٪</div></div>
            </div>
            <div class="chart-box">
                <h3>📊 توزیع فروش مشتریان</h3>
                <canvas id="customerChart" height="200"></canvas>
            </div>
        </div>

        <div class="page" id="page-inventory">
            <h2 style="margin-bottom:20px;">📦 تحلیل موجودی</h2>
            <div class="grid">
                <div class="card"><div class="icon">📦</div><h3>ارزش موجودی</h3><div class="value">۱۶۸,۴۰۰,۰۰۰</div><div class="sub green">▼ ۶.۲٪</div></div>
                <div class="card"><div class="icon">🔄</div><h3>گردش موجودی</h3><div class="value">۸.۴×</div><div class="sub green">▲ ۰.۸×</div></div>
                <div class="card"><div class="icon">📊</div><h3>نرخ تکمیل سفارش</h3><div class="value">۹۶.۳%</div><div class="sub green">▲ ۱.۲٪</div></div>
            </div>
            <div class="chart-box">
                <h3>📊 تحلیل COGS</h3>
                <canvas id="inventoryChart" height="200"></canvas>
            </div>
        </div>

    </div>

    <script>
        function showPage(pageId) {{
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
            document.getElementById('page-' + pageId).classList.add('active');
            const tab = document.querySelector(`.nav-tab[onclick="showPage('${{pageId}}')"]`);
            if (tab) tab.classList.add('active');
        }}

        const months = {json.dumps(k['months'][:len(k['monthly'])])};
        const monthlyData = {json.dumps(k['monthly'])};
        const channels = {json.dumps(k['channel_labels'])};
        const channelValues = {json.dumps(k['channel_values'])};
        const quarterlyData = {json.dumps(k['quarterly'])};
        const categories = {json.dumps(k['category_labels'])};
        const categoryValues = {json.dumps(k['category_values'])};
        const regions = {json.dumps(k['region_labels'])};
        const regionValues = {json.dumps(k['region_values'])};

        new Chart(document.getElementById('monthlyChart'), {{
            type: 'line',
            data: {{
                labels: months,
                datasets: [{{ 
                    label: 'فروش (ریال)', 
                    data: monthlyData, 
                    borderColor: '#2b6cb0', 
                    backgroundColor: 'rgba(43,108,176,0.1)', 
                    tension: 0.3, 
                    fill: true,
                    pointBackgroundColor: '#2b6cb0'
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ beginAtZero: true }} }} }}
        }});

        new Chart(document.getElementById('channelChart'), {{
            type: 'doughnut',
            data: {{
                labels: channels,
                datasets: [{{ 
                    data: channelValues, 
                    backgroundColor: ['#2b6cb0', '#48bb78', '#ed8936', '#9f7aea']
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ position: 'bottom' }} }} }}
        }});

        new Chart(document.getElementById('quarterlyChart'), {{
            type: 'bar',
            data: {{
                labels: ['Q1', 'Q2', 'Q3', 'Q4'],
                datasets: [{{ 
                    label: 'فروش فصلی', 
                    data: quarterlyData, 
                    backgroundColor: ['#2b6cb0', '#48bb78', '#ed8936', '#9f7aea'],
                    borderRadius: 4
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ beginAtZero: true }} }} }}
        }});

        new Chart(document.getElementById('categoryChart'), {{
            type: 'bar',
            data: {{
                labels: categories,
                datasets: [{{ 
                    label: 'فروش بر اساس دسته', 
                    data: categoryValues, 
                    backgroundColor: ['#2b6cb0', '#48bb78', '#ed8936', '#9f7aea', '#fc8181'],
                    borderRadius: 4
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ beginAtZero: true }} }} }}
        }});

        new Chart(document.getElementById('regionChart'), {{
            type: 'bar',
            data: {{
                labels: regions,
                datasets: [{{ 
                    label: 'فروش منطقه‌ای', 
                    data: regionValues, 
                    backgroundColor: ['#2b6cb0', '#48bb78', '#ed8936', '#9f7aea', '#fc8181'],
                    borderRadius: 4
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }}, scales: {{ y: {{ beginAtZero: true }} }} }}
        }});

        new Chart(document.getElementById('plChart'), {{
            type: 'bar',
            data: {{
                labels: months,
                datasets: [
                    {{ label: 'درآمد ناخالص', data: monthlyData.map(v => v * 1.12), backgroundColor: 'rgba(43,108,176,0.6)', borderRadius: 4 }},
                    {{ label: 'تخفیف', data: monthlyData.map(v => v * 0.12), backgroundColor: 'rgba(252,129,129,0.6)', borderRadius: 4 }}
                ]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ position: 'top' }} }}, scales: {{ y: {{ beginAtZero: true }} }} }}
        }});

        new Chart(document.getElementById('channelTrendChart'), {{
            type: 'line',
            data: {{
                labels: months,
                datasets: [
                    {{ label: 'Modern Trade', data: monthlyData.map(v => v * 0.45), borderColor: '#2b6cb0', tension: 0.3 }},
                    {{ label: 'Traditional Trade', data: monthlyData.map(v => v * 0.30), borderColor: '#48bb78', tension: 0.3 }},
                    {{ label: 'E-Commerce', data: monthlyData.map(v => v * 0.15), borderColor: '#ed8936', tension: 0.3 }},
                    {{ label: 'HoReCa', data: monthlyData.map(v => v * 0.10), borderColor: '#9f7aea', tension: 0.3 }}
                ]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ position: 'top' }} }}, scales: {{ y: {{ beginAtZero: true }} }} }}
        }});

        new Chart(document.getElementById('customerChart'), {{
            type: 'doughnut',
            data: {{
                labels: ['Top 10%', 'Next 20%', 'Middle 40%', 'Bottom 30%'],
                datasets: [{{ 
                    data: [35, 28, 25, 12], 
                    backgroundColor: ['#2b6cb0', '#48bb78', '#ed8936', '#fc8181']
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ position: 'bottom' }} }} }}
        }});

        new Chart(document.getElementById('inventoryChart'), {{
            type: 'doughnut',
            data: {{
                labels: ['مواد اولیه', 'بسته‌بندی', 'سربار تولید', 'نیروی کار'],
                datasets: [{{ 
                    data: [45, 25, 18, 12], 
                    backgroundColor: ['#2b6cb0', '#48bb78', '#ed8936', '#9f7aea']
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ position: 'bottom' }} }} }}
        }});
    </script>
</body>
</html>'''
    
    return html

# ─────────────────────────────────────────────
# 4. اجرا
# ─────────────────────────────────────────────
def generate_pro_dashboard():
    print("🚀 شروع تولید داشبورد حرفه‌ای...")
    df = make_sales()
    kpis = calculate_kpis(df)
    
    df.to_excel("FMCG_Analysis_Template.xlsx", index=False)
    print("✅ فایل اکسل ساخته شد: FMCG_Analysis_Template.xlsx")
    
    html_content = build_pro_dashboard(kpis)
    with open("FMCG_PRO_Dashboard.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    print("✅ داشبورد حرفه‌ای ساخته شد: FMCG_PRO_Dashboard.html")
    
    print("\n🎉 همه فایل‌ها با موفقیت ساخته شدند!")

if __name__ == "__main__":
    generate_pro_dashboard()