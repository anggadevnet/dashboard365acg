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

# Mapping license names to short names
LICENSE_MAP = {
    "Office 365 E1": "E1",
    "Office 365 E3": "E3",
    "Microsoft 365 E3": "ME3",
    "Microsoft 365 E5": "ME5",
    "Office 365 F3": "F3",
    "Microsoft 365 F3": "MF3",
    "Exchange Online (Plan 1)": "EXO",
    "Power BI Pro": "PBI Pro",
    "Power BI Premium Per User": "PBI Premium",
    "Visio Plan 2": "Visio",
    "Planner Plan 1": "Planner",
    "Planner and Project Plan 3": "Project",
    "Power Automate Premium": "PA Premium",
    "Power Apps Premium": "PowerApps",
    "Microsoft Entra ID P1": "Entra",
    "Teams Premium (for Departments)": "Teams Premium",
    "Microsoft 365 Copilot": "Copilot"
}

# HTML Dashboard UI Modern
DASHBOARD_HTML = '''
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>M365 License Monitor | Lintasarta</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;14..32,400;14..32,500;14..32,600;14..32,700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #e9ecef 100%);
            min-height: 100vh;
        }
        
        /* Navbar Modern */
        .navbar {
            background: rgba(255,255,255,0.95);
            backdrop-filter: blur(10px);
            padding: 16px 32px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.05);
            position: sticky;
            top: 0;
            z-index: 100;
            border-bottom: 1px solid rgba(0,0,0,0.05);
        }
        .logo {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .logo-icon {
            width: 40px;
            height: 40px;
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
            box-shadow: 0 4px 10px rgba(0,120,212,0.3);
        }
        .logo-text {
            font-size: 20px;
            font-weight: 700;
            background: linear-gradient(135deg, #1a1a2e, #16213e);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }
        .logo-sub {
            font-size: 12px;
            color: #666;
            font-weight: 400;
        }
        .user-info {
            display: flex;
            align-items: center;
            gap: 20px;
            background: #f8f9fa;
            padding: 8px 20px;
            border-radius: 40px;
        }
        .user-avatar {
            width: 36px;
            height: 36px;
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 600;
        }
        .logout-btn {
            background: none;
            border: none;
            color: #dc3545;
            cursor: pointer;
            font-size: 18px;
            transition: transform 0.2s;
        }
        .logout-btn:hover { transform: scale(1.1); }
        
        /* Container */
        .container { max-width: 1600px; margin: 0 auto; padding: 24px 32px; }
        
        /* Welcome Banner */
        .welcome-banner {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 24px;
            padding: 24px 32px;
            margin-bottom: 32px;
            color: white;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
        }
        .welcome-title { font-size: 24px; font-weight: 700; margin-bottom: 8px; }
        .welcome-subtitle { opacity: 0.9; font-size: 14px; }
        .update-badge {
            background: rgba(255,255,255,0.2);
            padding: 8px 16px;
            border-radius: 40px;
            font-size: 13px;
        }
        
        /* Stats Grid Modern */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 32px;
        }
        .stat-card {
            background: white;
            border-radius: 20px;
            padding: 20px;
            transition: all 0.3s ease;
            cursor: pointer;
            border: 1px solid rgba(0,0,0,0.05);
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        .stat-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 24px rgba(0,0,0,0.1);
        }
        .stat-icon {
            width: 48px;
            height: 48px;
            border-radius: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 24px;
            margin-bottom: 16px;
        }
        .stat-value { font-size: 32px; font-weight: 800; margin-bottom: 4px; }
        .stat-label { color: #666; font-size: 13px; font-weight: 500; }
        .stat-trend { font-size: 12px; margin-top: 8px; color: #28a745; }
        
        /* Charts Grid */
        .charts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 24px;
            margin-bottom: 32px;
        }
        .chart-card {
            background: white;
            border-radius: 24px;
            padding: 24px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.04);
            border: 1px solid rgba(0,0,0,0.05);
        }
        .chart-title {
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        /* Filter Bar Modern */
        .filter-bar {
            background: white;
            border-radius: 20px;
            padding: 16px 24px;
            margin-bottom: 24px;
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        .search-box {
            flex: 1;
            min-width: 250px;
            padding: 12px 20px;
            border: 1px solid #e0e0e0;
            border-radius: 40px;
            font-size: 14px;
            transition: all 0.2s;
            background: #f8f9fa;
        }
        .search-box:focus {
            outline: none;
            border-color: #0078D4;
            background: white;
        }
        .filter-group {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        .filter-btn {
            padding: 10px 20px;
            border: none;
            border-radius: 40px;
            cursor: pointer;
            background: #f0f2f5;
            font-size: 13px;
            font-weight: 500;
            transition: all 0.2s;
            font-family: 'Inter', sans-serif;
        }
        .filter-btn i { margin-right: 6px; }
        .filter-btn:hover { background: #e0e0e0; }
        .filter-btn.active {
            background: #0078D4;
            color: white;
            box-shadow: 0 4px 12px rgba(0,120,212,0.3);
        }
        .filter-btn.warning.active {
            background: #dc3545;
            box-shadow: 0 4px 12px rgba(220,53,69,0.3);
        }
        .export-btn {
            background: linear-gradient(135deg, #28a745, #20c997);
            color: white;
            padding: 10px 24px;
            border: none;
            border-radius: 40px;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.2s;
        }
        .export-btn:hover { transform: scale(1.02); box-shadow: 0 4px 12px rgba(40,167,69,0.3); }
        
        /* Table Modern */
        .table-container {
            background: white;
            border-radius: 24px;
            padding: 0;
            overflow: hidden;
            box-shadow: 0 2px 12px rgba(0,0,0,0.04);
            border: 1px solid rgba(0,0,0,0.05);
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        th {
            text-align: left;
            padding: 16px 20px;
            background: #f8f9fa;
            font-weight: 600;
            color: #333;
            border-bottom: 1px solid #e0e0e0;
        }
        td {
            padding: 14px 20px;
            border-bottom: 1px solid #f0f0f0;
        }
        tr:hover { background: #fafbfc; }
        
        /* Badges */
        .badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 500;
            margin: 2px;
        }
        .badge-primary { background: #e3f2fd; color: #0078D4; }
        .badge-success { background: #d4edda; color: #28a745; }
        .badge-warning { background: #f8d7da; color: #dc3545; }
        .badge-info { background: #d1ecf1; color: #17a2b8; }
        .badge-guest { background: #e9ecef; color: #6c757d; }
        
        /* Row highlight */
        .row-blocked { background: #fff5f5; border-left: 3px solid #dc3545; }
        .row-inactive { background: #fffbf0; border-left: 3px solid #ffc107; }
        .row-guest { background: #f8f9fa; opacity: 0.85; }
        
        /* Loading */
        .loading {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 80px;
        }
        .spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #e0e0e0;
            border-top-color: #0078D4;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            margin-bottom: 20px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        
        @media (max-width: 768px) {
            .container { padding: 16px; }
            .stats-grid { grid-template-columns: repeat(2, 1fr); gap: 12px; }
            .charts-grid { grid-template-columns: 1fr; }
            .filter-bar { flex-direction: column; align-items: stretch; }
            .filter-group { justify-content: center; }
            th, td { padding: 10px 12px; font-size: 12px; }
        }
    </style>
</head>
<body>
    <div class="navbar">
        <div class="logo">
            <div class="logo-icon"><i class="fas fa-chart-line"></i></div>
            <div>
                <div class="logo-text">License Monitor</div>
                <div class="logo-sub">Lintasarta</div>
            </div>
        </div>
        <div class="user-info">
            <div class="user-avatar"><i class="fas fa-user"></i></div>
            <span style="font-weight: 500;">{{ user.name }}</span>
            <a href="/logout" class="logout-btn"><i class="fas fa-sign-out-alt"></i></a>
        </div>
    </div>
    
    <div class="container">
        <div id="loading" class="loading">
            <div class="spinner"></div>
            <p style="color: #666;">Loading data dari Microsoft 365...</p>
        </div>
        
        <div id="content" style="display:none;">
            <!-- Welcome Banner -->
            <div class="welcome-banner">
                <div>
                    <div class="welcome-title"><i class="fas fa-chart-pie"></i> Dashboard Overview</div>
                    <div class="welcome-subtitle">Monitoring lisensi dan aktivitas pengguna Microsoft 365</div>
                </div>
                <div class="update-badge"><i class="fas fa-sync-alt"></i> Auto refresh every 5 min</div>
            </div>
            
            <!-- Stats Cards -->
            <div class="stats-grid" id="statsGrid"></div>
            
            <!-- Charts -->
            <div class="charts-grid">
                <div class="chart-card">
                    <div class="chart-title"><i class="fas fa-chart-bar" style="color: #0078D4;"></i> Top License Distribution</div>
                    <canvas id="licenseChart" style="max-height: 300px;"></canvas>
                </div>
                <div class="chart-card">
                    <div class="chart-title"><i class="fas fa-chart-pie" style="color: #28a745;"></i> User Status Overview</div>
                    <canvas id="statusChart" style="max-height: 300px;"></canvas>
                </div>
            </div>
            
            <!-- Filter Bar -->
            <div class="filter-bar">
                <div class="search-box">
                    <i class="fas fa-search" style="color: #999; margin-right: 8px;"></i>
                    <input type="text" id="searchInput" placeholder="Cari nama, email, atau department..." style="border: none; background: transparent; width: 85%; outline: none;">
                </div>
                <div class="filter-group">
                    <button id="filterAll" class="filter-btn active"><i class="fas fa-users"></i> All</button>
                    <button id="filterInternal" class="filter-btn"><i class="fas fa-building"></i> Internal Lintasarta</button>
                    <button id="filterGuest" class="filter-btn"><i class="fas fa-globe"></i> Guest</button>
                    <button id="filterLicensed" class="filter-btn"><i class="fas fa-check-circle"></i> Licensed</button>
                    <button id="filterUnlicensed" class="filter-btn"><i class="fas fa-times-circle"></i> Unlicensed</button>
                    <button id="filterBlockedE1" class="filter-btn warning"><i class="fas fa-ban"></i> Blocked + E1</button>
                    <button id="filterBlockedE3" class="filter-btn warning"><i class="fas fa-ban"></i> Blocked + E3</button>
                </div>
                <button id="exportBtn" class="export-btn"><i class="fas fa-download"></i> Export CSV</button>
            </div>
            
            <!-- User Table -->
            <div class="table-container">
                <table id="userTable">
                    <thead>
                        <tr>
                            <th><i class="fas fa-user"></i> Name</th>
                            <th><i class="fas fa-envelope"></i> Email</th>
                            <th><i class="fas fa-building"></i> Department</th>
                            <th><i class="fas fa-tag"></i> Type</th>
                            <th><i class="fas fa-shield-alt"></i> Status</th>
                            <th><i class="fas fa-calendar-alt"></i> Last Sign In</th>
                            <th><i class="fas fa-key"></i> Licenses</th>
                            <th><i class="fas fa-hashtag"></i> Count</th>
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
            const colorMap = {
                internal: { bg: '#e3f2fd', icon: '#0078D4', iconName: 'building' },
                guest: { bg: '#f8f9fa', icon: '#6c757d', iconName: 'globe' },
                licensed: { bg: '#d4edda', icon: '#28a745', iconName: 'check-circle' },
                unlicensed: { bg: '#fff3cd', icon: '#ffc107', iconName: 'times-circle' },
                blockedE1: { bg: '#f8d7da', icon: '#dc3545', iconName: 'ban' },
                blockedE3: { bg: '#f8d7da', icon: '#dc3545', iconName: 'ban' }
            };
            
            document.getElementById('statsGrid').innerHTML = `
                <div class="stat-card" onclick="setFilter('internal')">
                    <div class="stat-icon" style="background: ${colorMap.internal.bg}"><i class="fas fa-building" style="color: ${colorMap.internal.icon}"></i></div>
                    <div class="stat-value">${summary.internal_users}</div>
                    <div class="stat-label">Internal Lintasarta</div>
                </div>
                <div class="stat-card" onclick="setFilter('guest')">
                    <div class="stat-icon" style="background: ${colorMap.guest.bg}"><i class="fas fa-globe" style="color: ${colorMap.guest.icon}"></i></div>
                    <div class="stat-value">${summary.guest_users}</div>
                    <div class="stat-label">Guest Users</div>
                </div>
                <div class="stat-card" onclick="setFilter('licensed')">
                    <div class="stat-icon" style="background: ${colorMap.licensed.bg}"><i class="fas fa-check-circle" style="color: ${colorMap.licensed.icon}"></i></div>
                    <div class="stat-value">${summary.licensed_users}</div>
                    <div class="stat-label">Licensed</div>
                </div>
                <div class="stat-card" onclick="setFilter('unlicensed')">
                    <div class="stat-icon" style="background: ${colorMap.unlicensed.bg}"><i class="fas fa-times-circle" style="color: ${colorMap.unlicensed.icon}"></i></div>
                    <div class="stat-value">${summary.unlicensed_users}</div>
                    <div class="stat-label">Unlicensed</div>
                </div>
                <div class="stat-card" onclick="setFilter('blocked_e1')">
                    <div class="stat-icon" style="background: ${colorMap.blockedE1.bg}"><i class="fas fa-ban" style="color: ${colorMap.blockedE1.icon}"></i></div>
                    <div class="stat-value">${summary.blocked_e1}</div>
                    <div class="stat-label">Blocked + E1</div>
                </div>
                <div class="stat-card" onclick="setFilter('blocked_e3')">
                    <div class="stat-icon" style="background: ${colorMap.blockedE3.bg}"><i class="fas fa-ban" style="color: ${colorMap.blockedE3.icon}"></i></div>
                    <div class="stat-value">${summary.blocked_e3}</div>
                    <div class="stat-label">Blocked + E3</div>
                </div>
            `;
        }
        
        function setFilter(filter) {
            currentFilter = filter;
            const btns = ['filterAll', 'filterInternal', 'filterGuest', 'filterLicensed', 'filterUnlicensed', 'filterBlockedE1', 'filterBlockedE3'];
            const mapping = {
                'filterAll':'all', 'filterInternal':'internal', 'filterGuest':'guest',
                'filterLicensed':'licensed', 'filterUnlicensed':'unlicensed',
                'filterBlockedE1':'blocked_e1', 'filterBlockedE3':'blocked_e3'
            };
            btns.forEach(btnId => {
                const btn = document.getElementById(btnId);
                if (mapping[btnId] === filter) btn.classList.add('active');
                else btn.classList.remove('active');
            });
            renderTable();
        }
        
        function renderCharts() {
            const labels = Object.keys(licenseStats).slice(0, 8);
            const values = Object.values(licenseStats).slice(0, 8);
            new Chart(document.getElementById('licenseChart'), {
                type: 'bar',
                data: { labels, datasets: [{ label: 'Users', data: values, backgroundColor: '#0078D4', borderRadius: 8 }] },
                options: { responsive: true, maintainAspectRatio: true, plugins: { legend: { display: false } } }
            });
            
            const internal = allUsers.filter(u => !u.is_guest && !u.sign_blocked).length;
            const guest = allUsers.filter(u => u.is_guest && !u.sign_blocked).length;
            const blocked = allUsers.filter(u => u.sign_blocked).length;
            new Chart(document.getElementById('statusChart'), {
                type: 'doughnut',
                data: { labels: ['Internal', 'Guest', 'Blocked'], datasets: [{ data: [internal, guest, blocked], backgroundColor: ['#0078D4', '#6c757d', '#dc3545'] }] },
                options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
            });
        }
        
        function getFilteredUsers() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            let filtered = allUsers;
            
            if (currentFilter === 'internal') filtered = allUsers.filter(u => !u.is_guest);
            else if (currentFilter === 'guest') filtered = allUsers.filter(u => u.is_guest);
            else if (currentFilter === 'licensed') filtered = allUsers.filter(u => u.license_count > 0 && !u.sign_blocked && !u.is_guest);
            else if (currentFilter === 'unlicensed') filtered = allUsers.filter(u => u.license_count === 0 && !u.sign_blocked && !u.is_guest);
            else if (currentFilter === 'blocked_e1') filtered = allUsers.filter(u => u.sign_blocked && u.has_e1);
            else if (currentFilter === 'blocked_e3') filtered = allUsers.filter(u => u.sign_blocked && u.has_e3);
            
            if (searchTerm) {
                filtered = filtered.filter(u => 
                    u.name.toLowerCase().includes(searchTerm) || 
                    u.email.toLowerCase().includes(searchTerm) ||
                    (u.department && u.department.toLowerCase().includes(searchTerm))
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
                if (user.sign_blocked) rowClass = 'row-blocked';
                else if (user.inactive_days > 90 && user.license_count > 0) rowClass = 'row-inactive';
                else if (user.is_guest) rowClass = 'row-guest';
                if (rowClass) row.className = rowClass;
                
                row.insertCell(0).innerHTML = `<strong>${user.name}</strong>`;
                row.insertCell(1).innerHTML = `<a href="mailto:${user.email}" style="color: #0078D4; text-decoration: none;">${user.email}</a>`;
                row.insertCell(2).innerHTML = user.department || '—';
                
                let typeHtml = user.is_guest ? '<span class="badge badge-guest"><i class="fas fa-globe"></i> Guest</span>' : '<span class="badge badge-primary"><i class="fas fa-building"></i> Internal</span>';
                row.insertCell(3).innerHTML = typeHtml;
                
                let statusHtml = '';
                if (user.sign_blocked) statusHtml = '<span class="badge badge-warning"><i class="fas fa-ban"></i> Blocked</span>';
                else if (user.inactive_days > 90) statusHtml = `<span class="badge badge-info"><i class="fas fa-clock"></i> Inactive ${user.inactive_days}d</span>`;
                else statusHtml = '<span class="badge badge-success"><i class="fas fa-check-circle"></i> Active</span>';
                row.insertCell(4).innerHTML = statusHtml;
                
                row.insertCell(5).innerHTML = user.last_sign_in || '<span class="badge badge-guest">Never</span>';
                row.insertCell(6).innerHTML = user.licenses.map(l => `<span class="badge badge-primary">${l}</span>`).join(' ') || '<span class="badge badge-warning">No License</span>';
                row.insertCell(7).innerHTML = `<span style="font-weight: 600;">${user.license_count}</span>`;
            });
        }
        
        function exportCSV() {
            const filtered = getFilteredUsers();
            let csv = "Name,Email,Department,User Type,Sign Status,Last Sign In,Inactive Days,Licenses,License Count\\n";
            filtered.forEach(u => {
                csv += `"${u.name}","${u.email}","${u.department || ''}","${u.is_guest ? 'Guest' : 'Internal'}","${u.sign_blocked ? 'Blocked' : (u.inactive_days > 90 ? 'Inactive' : 'Active')}","${u.last_sign_in || 'Never'}","${u.inactive_days || '-'}","${u.licenses.join('; ')}",${u.license_count}\\n`;
            });
            const blob = new Blob([csv], {type:'text/csv'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `m365_report_${new Date().toISOString().split('T')[0]}.csv`;
            a.click();
        }
        
        document.getElementById('searchInput').addEventListener('keyup', renderTable);
        document.getElementById('filterAll').onclick = () => setFilter('all');
        document.getElementById('filterInternal').onclick = () => setFilter('internal');
        document.getElementById('filterGuest').onclick = () => setFilter('guest');
        document.getElementById('filterLicensed').onclick = () => setFilter('licensed');
        document.getElementById('filterUnlicensed').onclick = () => setFilter('unlicensed');
        document.getElementById('filterBlockedE1').onclick = () => setFilter('blocked_e1');
        document.getElementById('filterBlockedE3').onclick = () => setFilter('blocked_e3');
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
    <title>M365 License Monitor | Lintasarta</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .login-card {
            background: white;
            border-radius: 32px;
            padding: 48px;
            max-width: 500px;
            width: 90%;
            text-align: center;
            box-shadow: 0 25px 50px rgba(0,0,0,0.2);
            animation: fadeInUp 0.5s ease;
        }
        @keyframes fadeInUp {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .logo-icon {
            width: 70px;
            height: 70px;
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 20px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 20px;
            font-size: 32px;
            box-shadow: 0 10px 20px rgba(0,120,212,0.3);
        }
        h1 { font-size: 28px; font-weight: 700; margin-bottom: 8px; color: #1a1a2e; }
        .subtitle { color: #666; margin-bottom: 32px; font-size: 14px; }
        .features {
            text-align: left;
            margin: 32px 0;
            background: #f8f9fa;
            padding: 20px 24px;
            border-radius: 20px;
        }
        .features li {
            list-style: none;
            margin: 12px 0;
            color: #333;
            font-size: 14px;
        }
        .features i {
            width: 24px;
            color: #0078D4;
            margin-right: 12px;
        }
        .btn-login {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            color: white;
            border: none;
            padding: 14px 32px;
            border-radius: 40px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 10px;
            transition: all 0.3s;
            text-decoration: none;
        }
        .btn-login:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(0,120,212,0.3);
        }
        .footer {
            margin-top: 32px;
            font-size: 11px;
            color: #999;
        }
    </style>
</head>
<body>
    <div class="login-card">
        <div class="logo-icon"><i class="fas fa-chart-line" style="color: white;"></i></div>
        <h1>License Monitor</h1>
        <p class="subtitle">Microsoft 365 License & Activity Monitoring</p>
        <div class="features">
            <li><i class="fas fa-check-circle"></i> Internal & Guest users</li>
            <li><i class="fas fa-tag"></i> License: E1, E3, ME3, F3</li>
            <li><i class="fas fa-ban"></i> Blocked users by license type</li>
            <li><i class="fas fa-clock"></i> Inactive >90 days tracking</li>
            <li><i class="fas fa-download"></i> Export to CSV</li>
        </div>
        <a href="/login" class="btn-login"><i class="fas fa-microsoft"></i> Login with Microsoft 365</a>
        <div class="footer">Powered by Lintasarta</div>
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
    
    # Ambil semua user
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
    
    # Ambil SKU license
    skus_response = requests.get("https://graph.microsoft.com/v1.0/subscribedSkus", headers=headers)
    skus = skus_response.json().get("value", []) if skus_response.status_code == 200 else []
    sku_map = {}
    for sku in skus:
        sku_id = sku.get("skuId")
        sku_name = sku.get("skuPartNumber", "Unknown")
        # Map to short name
        short_name = LICENSE_MAP.get(sku_name, sku_name)
        sku_map[sku_id] = short_name
    
    # Proses data
    processed_users = []
    license_stats = {}
    today = datetime.now()
    
    for user in users:
        email = user.get("userPrincipalName", "")
        is_guest = "#EXT#" in email or user.get("userType") == "Guest"
        
        # License
        licenses = user.get("assignedLicenses", [])
        license_names = []
        has_e1 = False
        has_e3 = False
        
        for lic in licenses:
            sku_id = lic.get("skuId")
            if sku_id and sku_id in sku_map:
                lic_name = sku_map[sku_id]
                license_names.append(lic_name)
                license_stats[lic_name] = license_stats.get(lic_name, 0) + 1
                if lic_name == "E1":
                    has_e1 = True
                elif lic_name == "E3" or lic_name == "ME3":
                    has_e3 = True
        
        sign_blocked = user.get("accountEnabled") == False
        
        # Last sign in
        last_sign_in_str = 'Never'
        inactive_days = None
        sign_in_activity = user.get("signInActivity", {})
        last_sign_in_date = sign_in_activity.get("lastSignInDateTime")
        
        if last_sign_in_date:
            try:
                last_sign_in = datetime.fromisoformat(last_sign_in_date.replace('Z', '+00:00'))
                inactive_days = (today - last_sign_in).days
                last_sign_in_str = last_sign_in.strftime('%Y-%m-%d')
            except:
                last_sign_in_str = last_sign_in_date[:10] if last_sign_in_date else 'Never'
                inactive_days = None
        
        processed_users.append({
            "name": user.get("displayName", "N/A"),
            "email": email,
            "department": user.get("department", "N/A"),
            "is_guest": is_guest,
            "license_count": len(license_names),
            "licenses": license_names,
            "sign_blocked": sign_blocked,
            "has_e1": has_e1,
            "has_e3": has_e3,
            "last_sign_in": last_sign_in_str,
            "inactive_days": inactive_days if inactive_days else None
        })
    
    # Pisahkan internal dan guest
    internal_users = [u for u in processed_users if not u['is_guest']]
    guest_users = [u for u in processed_users if u['is_guest']]
    
    # Statistik
    licensed_active = len([u for u in internal_users if u['license_count'] > 0 and not u['sign_blocked']])
    unlicensed_active = len([u for u in internal_users if u['license_count'] == 0 and not u['sign_blocked']])
    blocked_e1 = len([u for u in processed_users if u['sign_blocked'] and u['has_e1']])
    blocked_e3 = len([u for u in processed_users if u['sign_blocked'] and u['has_e3']])
    
    return jsonify({
        'users': processed_users,
        'license_stats': dict(sorted(license_stats.items(), key=lambda x: x[1], reverse=True)),
        'summary': {
            'total_users': len(processed_users),
            'internal_users': len(internal_users),
            'guest_users': len(guest_users),
            'licensed_users': licensed_active,
            'unlicensed_users': unlicensed_active,
            'blocked_e1': blocked_e1,
            'blocked_e3': blocked_e3
        }
    })

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
