import os
import requests
import msal
from flask import Flask, request, redirect, session, jsonify, render_template_string
from datetime import datetime, timedelta

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "rahasia-banget-2024")

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "User.Read.All", "Organization.Read.All", "AuditLog.Read.All"]

# HTML Dashboard FINAL VERSION (Tanpa MFA)
DASHBOARD_HTML = '''
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 License Monitor Pro</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f0f2f5;
        }
        .navbar {
            background: white;
            padding: 15px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            position: sticky;
            top: 0;
            z-index: 100;
        }
        .logo { font-size: 24px; font-weight: bold; color: #0078D4; }
        .user-info { display: flex; align-items: center; gap: 20px; flex-wrap: wrap; }
        .logout-btn {
            background: #dc3545;
            color: white;
            padding: 8px 20px;
            border-radius: 6px;
            text-decoration: none;
        }
        .container { max-width: 1400px; margin: 0 auto; padding: 30px; }
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 15px;
            margin-bottom: 30px;
        }
        .stat-card {
            background: white;
            border-radius: 12px;
            padding: 15px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            cursor: pointer;
            transition: transform 0.2s;
        }
        .stat-card:hover { transform: translateY(-3px); }
        .stat-value { font-size: 28px; font-weight: bold; margin: 5px 0; }
        .stat-label { color: #666; font-size: 12px; }
        .stat-card.warning .stat-value { color: #dc3545; }
        .stat-card.success .stat-value { color: #28a745; }
        .stat-card.primary .stat-value { color: #0078D4; }
        .stat-card.info .stat-value { color: #17a2b8; }
        .charts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .chart-card {
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .filter-bar {
            background: white;
            border-radius: 12px;
            padding: 15px 20px;
            margin-bottom: 20px;
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .search-box {
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 6px;
            flex: 1;
            min-width: 200px;
        }
        .filter-group {
            display: flex;
            gap: 5px;
            flex-wrap: wrap;
        }
        .filter-btn {
            padding: 8px 16px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            background: #e9ecef;
            transition: all 0.2s;
            font-size: 13px;
        }
        .filter-btn.active {
            background: #0078D4;
            color: white;
        }
        .filter-btn.warning.active {
            background: #dc3545;
            color: white;
        }
        .filter-btn.info.active {
            background: #17a2b8;
            color: white;
        }
        .export-btn {
            background: #28a745;
            color: white;
            padding: 8px 16px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
        }
        .table-container {
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            overflow-x: auto;
        }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #f8f9fa; font-weight: 600; position: sticky; top: 0; }
        tr:hover { background: #f5f5f5; }
        .badge {
            background: #0078D4;
            color: white;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 11px;
            display: inline-block;
            margin: 2px;
        }
        .badge-warning { background: #dc3545; }
        .badge-success { background: #28a745; }
        .badge-info { background: #17a2b8; }
        .badge-dark { background: #6c757d; }
        .sign-blocked { background: #fff5f5; border-left: 3px solid #dc3545; }
        .inactive-user { background: #fffbf0; border-left: 3px solid #ffc107; }
        .loading { text-align: center; padding: 50px; }
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #0078D4;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        @media (max-width: 768px) {
            .navbar { flex-direction: column; gap: 10px; }
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .charts-grid { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; }
        }
    </style>
</head>
<body>
    <div class="navbar">
        <div class="logo">📊 M365 License Monitor Pro</div>
        <div class="user-info">
            <span>👋 {{ user.name }} ({{ user.email }})</span>
            <a href="/logout" class="logout-btn">🚪 Logout</a>
        </div>
    </div>
    <div class="container">
        <div id="loading" class="loading"><div class="spinner"></div><p>Loading data dari Microsoft 365...</p></div>
        <div id="content" style="display:none;">
            <!-- Stats Cards -->
            <div class="stats-grid" id="statsGrid"></div>
            
            <!-- Charts -->
            <div class="charts-grid">
                <div class="chart-card"><h3>📈 Top 10 License Distribution</h3><canvas id="licenseChart"></canvas></div>
                <div class="chart-card"><h3>👥 User Status Overview</h3><canvas id="statusChart"></canvas></div>
            </div>
            
            <!-- Filter Bar -->
            <div class="filter-bar">
                <input type="text" id="searchInput" class="search-box" placeholder="🔍 Cari nama, email, atau department...">
                <div class="filter-group">
                    <button id="filterAll" class="filter-btn active">📋 All Users</button>
                    <button id="filterLicensed" class="filter-btn">✅ Licensed</button>
                    <button id="filterUnlicensed" class="filter-btn">❌ Unlicensed</button>
                    <button id="filterBlockedAll" class="filter-btn warning">🚫 All Blocked</button>
                    <button id="filterInactiveLicensed" class="filter-btn info">⏰ Inactive + Licensed</button>
                </div>
                <button id="exportBtn" class="export-btn">📥 Export CSV</button>
            </div>
            
            <!-- User Table -->
            <div class="table-container">
                <table id="userTable">
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th>Email</th>
                            <th>Department</th>
                            <th>User Type</th>
                            <th>Sign Status</th>
                            <th>Last Sign In</th>
                            <th>Licenses</th>
                            <th>Count</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>
        </div>
    </div>
    <script>
        let allUsers = [];
        let licenseStats = {};
        let currentFilter = 'all';
        
        async function loadData() {
            const res = await fetch('/api/license-data');
            const data = await res.json();
            if(data.error){ alert('Session expired'); window.location='/logout'; return; }
            
            allUsers = data.users;
            licenseStats = data.license_stats;
            
            updateStats(data.summary);
            renderCharts();
            renderTable();
            
            document.getElementById('loading').style.display = 'none';
            document.getElementById('content').style.display = 'block';
        }
        
        function updateStats(summary) {
            const statsGrid = document.getElementById('statsGrid');
            statsGrid.innerHTML = `
                <div class="stat-card primary" onclick="setFilter('all')">
                    <div class="stat-label">Total Users</div>
                    <div class="stat-value">${summary.total_users}</div>
                </div>
                <div class="stat-card success" onclick="setFilter('licensed')">
                    <div class="stat-label">Licensed</div>
                    <div class="stat-value">${summary.licensed_users}</div>
                </div>
                <div class="stat-card warning" onclick="setFilter('unlicensed')">
                    <div class="stat-label">Unlicensed</div>
                    <div class="stat-value">${summary.unlicensed_users}</div>
                </div>
                <div class="stat-card warning" onclick="setFilter('blocked_all')">
                    <div class="stat-label">Blocked</div>
                    <div class="stat-value">${summary.blocked_users}</div>
                </div>
                <div class="stat-card info" onclick="setFilter('inactive_licensed')">
                    <div class="stat-label">Inactive + Licensed</div>
                    <div class="stat-value">${summary.inactive_licensed}</div>
                </div>
                <div class="stat-card primary" onclick="setFilter('all')">
                    <div class="stat-label">Coverage</div>
                    <div class="stat-value">${summary.coverage}%</div>
                </div>
            `;
        }
        
        function setFilter(filter) {
            currentFilter = filter;
            updateFilterButtons();
            renderTable();
        }
        
        function renderCharts() {
            const labels = Object.keys(licenseStats).slice(0, 10);
            const values = Object.values(licenseStats).slice(0, 10);
            new Chart(document.getElementById('licenseChart'), {
                type: 'bar',
                data: { labels, datasets: [{ label: 'Users', data: values, backgroundColor: '#0078D4' }] },
                options: { responsive: true, maintainAspectRatio: true }
            });
            
            const licensed = allUsers.filter(u => u.license_count > 0 && !u.sign_blocked).length;
            const unlicensed = allUsers.filter(u => u.license_count === 0 && !u.sign_blocked).length;
            const blocked = allUsers.filter(u => u.sign_blocked).length;
            new Chart(document.getElementById('statusChart'), {
                type: 'doughnut',
                data: { labels: ['Licensed Active', 'Unlicensed Active', 'Blocked'], datasets: [{ data: [licensed, unlicensed, blocked], backgroundColor: ['#28a745', '#ffc107', '#dc3545'] }] },
                options: { responsive: true }
            });
        }
        
        function getFilteredUsers() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            let filtered = allUsers;
            
            if (currentFilter === 'licensed') filtered = allUsers.filter(u => u.license_count > 0 && !u.sign_blocked);
            else if (currentFilter === 'unlicensed') filtered = allUsers.filter(u => u.license_count === 0 && !u.sign_blocked);
            else if (currentFilter === 'blocked_all') filtered = allUsers.filter(u => u.sign_blocked);
            else if (currentFilter === 'inactive_licensed') filtered = allUsers.filter(u => u.inactive_days > 90 && u.license_count > 0 && !u.sign_blocked);
            
            if (searchTerm) {
                filtered = filtered.filter(u => 
                    u.name.toLowerCase().includes(searchTerm) || 
                    u.email.toLowerCase().includes(searchTerm) ||
                    u.department.toLowerCase().includes(searchTerm)
                );
            }
            return filtered;
        }
        
        function renderTable() {
            const filtered = getFilteredUsers();
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';
            
            filtered.forEach(user => {
                const row = tbody.insertRow();
                let rowClass = '';
                if (user.sign_blocked) rowClass = 'sign-blocked';
                else if (user.inactive_days > 90 && user.license_count > 0) rowClass = 'inactive-user';
                if (rowClass) row.classList.add(rowClass);
                
                row.insertCell(0).innerHTML = user.name;
                row.insertCell(1).innerHTML = `<a href="mailto:${user.email}">${user.email}</a>`;
                row.insertCell(2).innerHTML = user.department || '-';
                row.insertCell(3).innerHTML = user.user_type === 'Member' ? '👤 Member' : '👥 Guest';
                
                let signStatus = '';
                if (user.sign_blocked) signStatus = '<span class="badge badge-warning">🚫 Blocked</span>';
                else if (user.inactive_days > 90) signStatus = '<span class="badge badge-info">⏰ Inactive ' + user.inactive_days + ' days</span>';
                else signStatus = '<span class="badge badge-success">✅ Active</span>';
                row.insertCell(4).innerHTML = signStatus;
                
                row.insertCell(5).innerHTML = user.last_sign_in || '<span class="badge badge-dark">Never</span>';
                row.insertCell(6).innerHTML = user.licenses.map(l => `<span class="badge">${l}</span>`).join(' ') || '<span class="badge badge-warning">No License</span>';
                row.insertCell(7).innerHTML = user.license_count;
            });
        }
        
        function exportCSV() {
            const filtered = getFilteredUsers();
            let csv = "Name,Email,Department,User Type,Sign Status,Last Sign In,Inactive Days,Licenses,License Count\\n";
            filtered.forEach(u => {
                csv += `"${u.name}","${u.email}","${u.department}","${u.user_type}","${u.sign_blocked ? 'Blocked' : (u.inactive_days > 90 ? 'Inactive ' + u.inactive_days + ' days' : 'Active')}","${u.last_sign_in || 'Never'}","${u.inactive_days}","${u.licenses.join('; ')}",${u.license_count}\\n`;
            });
            const blob = new Blob([csv], {type:'text/csv'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `m365_full_report_${new Date().toISOString().split('T')[0]}.csv`;
            a.click();
        }
        
        function updateFilterButtons() {
            const btns = ['filterAll', 'filterLicensed', 'filterUnlicensed', 'filterBlockedAll', 'filterInactiveLicensed'];
            const mapping = {
                'filterAll': 'all', 'filterLicensed': 'licensed', 'filterUnlicensed': 'unlicensed',
                'filterBlockedAll': 'blocked_all', 'filterInactiveLicensed': 'inactive_licensed'
            };
            btns.forEach(btnId => {
                const btn = document.getElementById(btnId);
                if (mapping[btnId] === currentFilter) btn.classList.add('active');
                else btn.classList.remove('active');
            });
        }
        
        // Event listeners
        document.getElementById('searchInput').addEventListener('keyup', renderTable);
        document.getElementById('filterAll').onclick = () => setFilter('all');
        document.getElementById('filterLicensed').onclick = () => setFilter('licensed');
        document.getElementById('filterUnlicensed').onclick = () => setFilter('unlicensed');
        document.getElementById('filterBlockedAll').onclick = () => setFilter('blocked_all');
        document.getElementById('filterInactiveLicensed').onclick = () => setFilter('inactive_licensed');
        document.getElementById('exportBtn').onclick = exportCSV;
        
        loadData();
        setInterval(loadData, 300000);
    </script>
</body>
</html>
'''

LOGIN_HTML = '''
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 License Monitor Pro</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .card {
            background: white;
            border-radius: 20px;
            padding: 50px;
            max-width: 550px;
            text-align: center;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }
        .logo { font-size: 64px; margin-bottom: 20px; }
        h1 { color: #333; margin-bottom: 10px; font-size: 28px; }
        .subtitle { color: #666; margin-bottom: 30px; }
        .features { text-align: left; margin: 30px 0; list-style: none; }
        .features li { margin: 12px 0; color: #555; }
        .btn {
            background: #0078D4;
            color: white;
            padding: 15px 40px;
            border-radius: 10px;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            transition: transform 0.2s;
        }
        .btn:hover { transform: translateY(-2px); background: #005a9e; }
        .footer { margin-top: 30px; font-size: 12px; color: #999; }
        .version { margin-top: 15px; font-size: 11px; color: #bbb; }
    </style>
</head>
<body>
    <div class="card">
        <div class="logo">📊</div>
        <h1>Microsoft 365 License Monitor Pro</h1>
        <p class="subtitle">License & User Activity Monitoring</p>
        <ul class="features">
            <li>✅ Full user license assignment tracking</li>
            <li>🚫 Sign-blocked users detection</li>
            <li>⏰ Inactive users with licenses (cost saving)</li>
            <li>📊 Advanced analytics & charts</li>
            <li>📥 Export full report to CSV</li>
        </ul>
        <a href="/login" class="btn">🔐 Login dengan Microsoft 365</a>
        <div class="footer">
            <p>Akses menggunakan akun Microsoft 365 perusahaan Anda</p>
            <div class="version">Version 5.0 - Fast Loading</div>
        </div>
    </div>
</body>
</html>
'''

def get_msal_app():
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY
    )

@app.route('/')
def home():
    if 'user' in session:
        return redirect('/dashboard')
    return LOGIN_HTML

@app.route('/login')
def login():
    app_msal = get_msal_app()
    auth_url = app_msal.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    return redirect(auth_url)

@app.route('/callback')
def callback():
    code = request.args.get('code')
    if not code:
        return "Error: No code provided", 400
    
    app_msal = get_msal_app()
    result = app_msal.acquire_token_by_authorization_code(
        code=code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    
    if "access_token" in result:
        session['access_token'] = result['access_token']
        headers = {"Authorization": f"Bearer {result['access_token']}"}
        me_response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
        user_info = me_response.json() if me_response.status_code == 200 else {}
        session['user'] = {
            'name': user_info.get('displayName', 'Admin'),
            'email': user_info.get('userPrincipalName', '')
        }
        return redirect('/dashboard')
    else:
        return f"Login failed: {result.get('error_description', 'Unknown error')}", 400

@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect('/')
    return render_template_string(DASHBOARD_HTML, user=session['user'])

@app.route('/api/license-data')
def api_license_data():
    token = session.get('access_token')
    if not token:
        return jsonify({'error': 'Unauthorized'}), 401
    
    headers = {"Authorization": f"Bearer {token}"}
    
    # Ambil semua user dengan detail lengkap
    users = []
    url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,assignedLicenses,userType,department,accountEnabled,signInActivity&$top=999"
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            users.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        else:
            break
    
    # Ambil daftar SKU license
    skus_response = requests.get("https://graph.microsoft.com/v1.0/subscribedSkus", headers=headers)
    skus = skus_response.json().get("value", []) if skus_response.status_code == 200 else []
    sku_map = {sku.get("skuId"): sku.get("skuPartNumber", "Unknown") for sku in skus}
    
    # Proses data user (TANPA MFA)
    processed_users = []
    license_stats = {}
    blocked_count = 0
    inactive_licensed_count = 0
    today = datetime.now()
    
    for user in users:
        # Proses license
        licenses = user.get("assignedLicenses", [])
        license_names = []
        for lic in licenses:
            sku_id = lic.get("skuId")
            if sku_id and sku_id in sku_map:
                license_names.append(sku_map[sku_id])
                license_stats[sku_map[sku_id]] = license_stats.get(sku_map[sku_id], 0) + 1
        
        # Deteksi sign blocked
        sign_blocked = user.get("accountEnabled") == False
        
        # Deteksi last sign in & inactive days
        last_sign_in = None
        inactive_days = 999
        sign_in_activity = user.get("signInActivity", {})
        last_sign_in_date = sign_in_activity.get("lastSignInDateTime")
        
        if last_sign_in_date:
            try:
                last_sign_in = datetime.fromisoformat(last_sign_in_date.replace('Z', '+00:00'))
                inactive_days = (today - last_sign_in).days
                last_sign_in_str = last_sign_in.strftime('%Y-%m-%d')
            except:
                last_sign_in_str = last_sign_in_date[:10] if last_sign_in_date else 'Never'
                inactive_days = 999
        else:
            last_sign_in_str = 'Never'
            inactive_days = 999
        
        # Count inactive licensed users (inactive >90 days, has license, not blocked)
        if inactive_days > 90 and len(license_names) > 0 and not sign_blocked:
            inactive_licensed_count += 1
        
        if sign_blocked:
            blocked_count += 1
        
        processed_users.append({
            "name": user.get("displayName", "N/A"),
            "email": user.get("userPrincipalName", "N/A"),
            "department": user.get("department", "N/A"),
            "user_type": user.get("userType", "Member"),
            "license_count": len(license_names),
            "licenses": license_names,
            "sign_blocked": sign_blocked,
            "last_sign_in": last_sign_in_str if last_sign_in_str != 'Never' else 'Never',
            "inactive_days": inactive_days if inactive_days < 900 else 999
        })
    
    licensed_count = len([u for u in processed_users if u['license_count'] > 0 and not u['sign_blocked']])
    unlicensed_count = len([u for u in processed_users if u['license_count'] == 0 and not u['sign_blocked']])
    
    return jsonify({
        'users': processed_users,
        'license_stats': dict(sorted(license_stats.items(), key=lambda x: x[1], reverse=True)),
        'summary': {
            'total_users': len(processed_users),
            'licensed_users': licensed_count,
            'unlicensed_users': unlicensed_count,
            'blocked_users': blocked_count,
            'inactive_licensed': inactive_licensed_count,
            'coverage': round(licensed_count / len(processed_users) * 100, 1) if processed_users else 0
        }
    })

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
