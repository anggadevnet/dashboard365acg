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

# Mapping license names
LICENSE_MAP = {
    "STANDARDPACK": "E1",
    "ENTERPRISEPACK": "E3",
    "SPE_E3": "ME3",
    "DESKLESSPACK": "F3",
    "SPE_F1": "MF3",
    "POWER_BI_PRO": "PBI Pro",
    "POWER_BI_STANDARD": "PBI Free",
    "PBI_PREMIUM_PER_USER": "PBI Premium",
    "VISIOCLIENT": "Visio",
    "PROJECTPROFESSIONAL": "Project",
    "PROJECT_P1": "Project Plan 1",
    "AAD_PREMIUM": "Entra P1",
    "FLOW_FREE": "Power Automate",
    "POWERAPPS_VIRAL": "Power Apps",
    "POWERAPPS_PER_USER": "Power Apps Premium",
    "Microsoft_365_Copilot": "Copilot",
    "Teams_Premium_(for_Departments)": "Teams Premium",
    "STREAM": "Stream",
    "EXCHANGESTANDARD": "Exchange"
}

# HTML Dashboard dengan Toggle
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
            background: #f0f2f5;
            min-height: 100vh;
        }
        
        /* Layout dengan Toggle */
        .app-wrapper {
            display: flex;
            min-height: 100vh;
        }
        
        /* SIDEBAR - Bisa toggle */
        .sidebar {
            width: 280px;
            background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%);
            color: white;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            position: fixed;
            height: 100vh;
            z-index: 1000;
            overflow-y: auto;
            box-shadow: 4px 0 20px rgba(0,0,0,0.1);
        }
        .sidebar.collapsed {
            width: 70px;
        }
        .sidebar.collapsed .sidebar-text,
        .sidebar.collapsed .sidebar-label,
        .sidebar.collapsed .stat-small {
            display: none;
        }
        .sidebar.collapsed .sidebar-icon {
            margin-right: 0;
        }
        .sidebar.collapsed .sidebar-item {
            justify-content: center;
            padding: 12px;
        }
        
        /* Main Content */
        .main-content {
            flex: 1;
            margin-left: 280px;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            min-height: 100vh;
        }
        .main-content.expanded {
            margin-left: 70px;
        }
        
        /* Toggle Button */
        .toggle-btn {
            position: fixed;
            left: 280px;
            top: 20px;
            z-index: 1001;
            background: #0078D4;
            color: white;
            border: none;
            width: 32px;
            height: 32px;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
        }
        .toggle-btn.collapsed {
            left: 70px;
        }
        .toggle-btn:hover {
            background: #005a9e;
            transform: scale(1.05);
        }
        
        /* Sidebar Content */
        .sidebar-header {
            padding: 24px 20px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            margin-bottom: 20px;
        }
        .sidebar-logo {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sidebar-logo-icon {
            width: 40px;
            height: 40px;
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
        }
        .sidebar-logo-text {
            font-weight: 700;
            font-size: 18px;
        }
        .sidebar-logo-sub {
            font-size: 11px;
            opacity: 0.7;
        }
        
        .sidebar-nav {
            padding: 0 16px;
        }
        .sidebar-item {
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 12px 16px;
            margin: 4px 0;
            border-radius: 12px;
            cursor: pointer;
            transition: all 0.2s;
            color: rgba(255,255,255,0.8);
        }
        .sidebar-item:hover {
            background: rgba(255,255,255,0.1);
            color: white;
        }
        .sidebar-item.active {
            background: #0078D4;
            color: white;
        }
        .sidebar-icon {
            width: 24px;
            font-size: 18px;
            text-align: center;
        }
        .sidebar-text {
            font-size: 14px;
            font-weight: 500;
        }
        
        .sidebar-divider {
            height: 1px;
            background: rgba(255,255,255,0.1);
            margin: 20px 16px;
        }
        
        .sidebar-stats {
            padding: 0 16px;
            margin-top: 20px;
        }
        .stat-item {
            background: rgba(255,255,255,0.05);
            border-radius: 12px;
            padding: 12px;
            margin-bottom: 8px;
        }
        .stat-label-small {
            font-size: 11px;
            opacity: 0.7;
            margin-bottom: 4px;
        }
        .stat-value-small {
            font-size: 20px;
            font-weight: 700;
        }
        
        /* Navbar Main */
        .main-navbar {
            background: white;
            padding: 16px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            position: sticky;
            top: 0;
            z-index: 99;
        }
        .page-title {
            font-size: 20px;
            font-weight: 700;
            color: #1a1a2e;
        }
        .user-info {
            display: flex;
            align-items: center;
            gap: 16px;
            background: #f8f9fa;
            padding: 8px 20px;
            border-radius: 40px;
        }
        .logout-btn {
            background: none;
            border: none;
            color: #dc3545;
            cursor: pointer;
            font-size: 16px;
        }
        
        /* Container */
        .container {
            padding: 24px;
            max-width: 100%;
            overflow-x: auto;
        }
        
        /* Stats Grid */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
        }
        .stat-card {
            background: white;
            border-radius: 20px;
            padding: 20px;
            cursor: pointer;
            transition: all 0.2s;
            border: 1px solid rgba(0,0,0,0.05);
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        .stat-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.1);
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
        
        /* Charts */
        .charts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
        }
        .chart-card {
            background: white;
            border-radius: 20px;
            padding: 20px;
        }
        .chart-title {
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 16px;
        }
        
        /* Filter Bar */
        .filter-bar {
            background: white;
            border-radius: 20px;
            padding: 16px 20px;
            margin-bottom: 20px;
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
            align-items: center;
        }
        .search-box {
            flex: 1;
            min-width: 250px;
            padding: 10px 16px;
            border: 1px solid #e0e0e0;
            border-radius: 40px;
            background: #f8f9fa;
        }
        .search-box input {
            border: none;
            background: transparent;
            width: 85%;
            outline: none;
        }
        .filter-group {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        .filter-btn {
            padding: 8px 16px;
            border: none;
            border-radius: 40px;
            cursor: pointer;
            background: #f0f2f5;
            font-size: 13px;
            font-weight: 500;
        }
        .filter-btn.active { background: #0078D4; color: white; }
        .filter-btn.warning.active { background: #dc3545; color: white; }
        .export-btn {
            background: linear-gradient(135deg, #28a745, #20c997);
            color: white;
            padding: 8px 20px;
            border: none;
            border-radius: 40px;
            cursor: pointer;
            font-weight: 600;
        }
        
        /* Table */
        .table-container {
            background: white;
            border-radius: 20px;
            overflow-x: auto;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            min-width: 800px;
        }
        th {
            padding: 14px 16px;
            text-align: left;
            background: #f8f9fa;
            font-weight: 600;
            border-bottom: 1px solid #e0e0e0;
        }
        td {
            padding: 12px 16px;
            border-bottom: 1px solid #f0f0f0;
        }
        tr:hover { background: #fafbfc; }
        
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
        
        /* Billing Table */
        .billing-table-container {
            background: white;
            border-radius: 20px;
            overflow-x: auto;
        }
        .billing-header {
            padding: 20px;
            border-bottom: 1px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            flex-wrap: wrap;
        }
        .progress-bar {
            width: 100px;
            height: 6px;
            background: #e0e0e0;
            border-radius: 3px;
            overflow: hidden;
        }
        .progress-fill {
            height: 100%;
            border-radius: 3px;
        }
        .progress-fill.green { background: #28a745; }
        .progress-fill.yellow { background: #ffc107; }
        .progress-fill.red { background: #dc3545; }
        
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
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        
        .hidden { display: none; }
        
        @media (max-width: 768px) {
            .sidebar { transform: translateX(-100%); }
            .sidebar.mobile-open { transform: translateX(0); }
            .main-content { margin-left: 0 !important; }
            .toggle-btn { left: 20px; }
            .toggle-btn.collapsed { left: 20px; }
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .charts-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
<div class="app-wrapper">
    <!-- Toggle Button -->
    <button class="toggle-btn" id="toggleBtn" onclick="toggleSidebar()">
        <i class="fas fa-bars"></i>
    </button>
    
    <!-- Sidebar -->
    <div class="sidebar" id="sidebar">
        <div class="sidebar-header">
            <div class="sidebar-logo">
                <div class="sidebar-logo-icon"><i class="fas fa-chart-line"></i></div>
                <div>
                    <div class="sidebar-logo-text">License Monitor</div>
                    <div class="sidebar-logo-sub">Lintasarta</div>
                </div>
            </div>
        </div>
        
        <div class="sidebar-nav">
            <div class="sidebar-item active" id="navDashboardBtn" onclick="showPage('dashboard')">
                <div class="sidebar-icon"><i class="fas fa-tachometer-alt"></i></div>
                <div class="sidebar-text">Dashboard</div>
            </div>
            <div class="sidebar-item" id="navBillingBtn" onclick="showPage('billing')">
                <div class="sidebar-icon"><i class="fas fa-dollar-sign"></i></div>
                <div class="sidebar-text">Billing & Subscriptions</div>
            </div>
        </div>
        
        <div class="sidebar-divider"></div>
        
        <div class="sidebar-stats" id="sidebarStats">
            <div class="stat-item">
                <div class="stat-label-small">Total Users</div>
                <div class="stat-value-small" id="sidebarTotal">-</div>
            </div>
            <div class="stat-item">
                <div class="stat-label-small">Licensed</div>
                <div class="stat-value-small" id="sidebarLicensed">-</div>
            </div>
            <div class="stat-item">
                <div class="stat-label-small">Blocked + E1</div>
                <div class="stat-value-small" id="sidebarBlockedE1">-</div>
            </div>
            <div class="stat-item">
                <div class="stat-label-small">Blocked + E3</div>
                <div class="stat-value-small" id="sidebarBlockedE3">-</div>
            </div>
        </div>
    </div>
    
    <!-- Main Content -->
    <div class="main-content" id="mainContent">
        <div class="main-navbar">
            <div class="page-title" id="pageTitle"><i class="fas fa-chart-pie"></i> Dashboard</div>
            <div class="user-info">
                <i class="fas fa-user-circle" style="font-size: 20px; color: #0078D4;"></i>
                <span>{{ user.name }}</span>
                <a href="/logout" class="logout-btn"><i class="fas fa-sign-out-alt"></i></a>
            </div>
        </div>
        
        <div class="container">
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <p>Loading data dari Microsoft 365...</p>
            </div>
            
            <!-- Dashboard Page -->
            <div id="dashboardPage">
                <div class="stats-grid" id="statsGrid"></div>
                
                <div class="charts-grid">
                    <div class="chart-card">
                        <div class="chart-title"><i class="fas fa-chart-bar"></i> Top License Distribution</div>
                        <canvas id="licenseChart" style="max-height: 280px;"></canvas>
                    </div>
                    <div class="chart-card">
                        <div class="chart-title"><i class="fas fa-chart-pie"></i> User Status Overview</div>
                        <canvas id="statusChart" style="max-height: 280px;"></canvas>
                    </div>
                </div>
                
                <div class="filter-bar">
                    <div class="search-box">
                        <i class="fas fa-search"></i>
                        <input type="text" id="searchInput" placeholder="Cari nama, email, atau department...">
                    </div>
                    <div class="filter-group">
                        <button id="filterAll" class="filter-btn active">All</button>
                        <button id="filterInternal" class="filter-btn">Internal</button>
                        <button id="filterGuest" class="filter-btn">Guest</button>
                        <button id="filterLicensed" class="filter-btn">Licensed</button>
                        <button id="filterUnlicensed" class="filter-btn">Unlicensed</button>
                        <button id="filterBlockedE1" class="filter-btn warning">Blocked + E1</button>
                        <button id="filterBlockedE3" class="filter-btn warning">Blocked + E3</button>
                    </div>
                    <button id="exportBtn" class="export-btn"><i class="fas fa-download"></i> Export</button>
                </div>
                
                <div class="table-container">
                    <table id="userTable">
                        <thead>
                            <tr><th>Name</th><th>Email</th><th>Department</th><th>Type</th><th>Status</th><th>Last Sign In</th><th>Licenses</th><th>Count</th></tr>
                        </thead>
                        <tbody id="tableBody"></tbody>
                    </table>
                </div>
            </div>
            
            <!-- Billing Page -->
            <div id="billingPage" class="hidden">
                <div class="billing-table-container">
                    <div class="billing-header">
                        <div><i class="fas fa-tags"></i> <strong>Active Subscriptions</strong></div>
                        <div id="totalSKU"></div>
                    </div>
                    <div style="overflow-x: auto;">
                        <table id="billingTable">
                            <thead><tr><th>License Name</th><th>SKU Code</th><th>Total</th><th>Used</th><th>Available</th><th>Usage</th></tr></thead>
                            <tbody id="billingBody"></tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<script>
    let allUsers = [];
    let licenseStats = {};
    let subscriptions = [];
    let currentFilter = 'all';
    
    // Toggle Sidebar
    function toggleSidebar() {
        const sidebar = document.getElementById('sidebar');
        const mainContent = document.getElementById('mainContent');
        const toggleBtn = document.getElementById('toggleBtn');
        
        sidebar.classList.toggle('collapsed');
        mainContent.classList.toggle('expanded');
        toggleBtn.classList.toggle('collapsed');
        
        // Simpan state ke localStorage
        const isCollapsed = sidebar.classList.contains('collapsed');
        localStorage.setItem('sidebarCollapsed', isCollapsed);
    }
    
    // Load saved sidebar state
    function loadSidebarState() {
        const isCollapsed = localStorage.getItem('sidebarCollapsed') === 'true';
        if (isCollapsed) {
            document.getElementById('sidebar').classList.add('collapsed');
            document.getElementById('mainContent').classList.add('expanded');
            document.getElementById('toggleBtn').classList.add('collapsed');
        }
    }
    
    async function loadData() {
        const res = await fetch('/api/license-data');
        const data = await res.json();
        if(data.error){ alert('Session expired'); window.location='/logout'; return; }
        allUsers = data.users;
        licenseStats = data.license_stats;
        subscriptions = data.subscriptions;
        updateStats(data.summary);
        updateSidebarStats(data.summary);
        renderBillingTable();
        renderCharts();
        renderTable();
        document.getElementById('loading').style.display = 'none';
        document.getElementById('dashboardPage').classList.remove('hidden');
    }
    
    function updateSidebarStats(summary) {
        document.getElementById('sidebarTotal').innerHTML = summary.total_users;
        document.getElementById('sidebarLicensed').innerHTML = summary.licensed_users;
        document.getElementById('sidebarBlockedE1').innerHTML = summary.blocked_e1;
        document.getElementById('sidebarBlockedE3').innerHTML = summary.blocked_e3;
    }
    
    function updateStats(summary) {
        document.getElementById('statsGrid').innerHTML = `
            <div class="stat-card" onclick="setFilter('internal')">
                <div class="stat-icon" style="background:#e3f2fd"><i class="fas fa-building" style="color:#0078D4"></i></div>
                <div class="stat-value">${summary.internal_users}</div><div class="stat-label">Internal Users</div>
            </div>
            <div class="stat-card" onclick="setFilter('guest')">
                <div class="stat-icon" style="background:#f8f9fa"><i class="fas fa-globe" style="color:#6c757d"></i></div>
                <div class="stat-value">${summary.guest_users}</div><div class="stat-label">Guest Users</div>
            </div>
            <div class="stat-card" onclick="setFilter('licensed')">
                <div class="stat-icon" style="background:#d4edda"><i class="fas fa-check-circle" style="color:#28a745"></i></div>
                <div class="stat-value">${summary.licensed_users}</div><div class="stat-label">Licensed</div>
            </div>
            <div class="stat-card" onclick="setFilter('unlicensed')">
                <div class="stat-icon" style="background:#fff3cd"><i class="fas fa-times-circle" style="color:#ffc107"></i></div>
                <div class="stat-value">${summary.unlicensed_users}</div><div class="stat-label">Unlicensed</div>
            </div>
            <div class="stat-card" onclick="setFilter('blocked_e1')">
                <div class="stat-icon" style="background:#f8d7da"><i class="fas fa-ban" style="color:#dc3545"></i></div>
                <div class="stat-value">${summary.blocked_e1}</div><div class="stat-label">Blocked + E1</div>
            </div>
            <div class="stat-card" onclick="setFilter('blocked_e3')">
                <div class="stat-icon" style="background:#f8d7da"><i class="fas fa-ban" style="color:#dc3545"></i></div>
                <div class="stat-value">${summary.blocked_e3}</div><div class="stat-label">Blocked + E3</div>
            </div>
        `;
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
    
    function renderBillingTable() {
        const tbody = document.getElementById('billingBody');
        tbody.innerHTML = '';
        subscriptions.forEach(sub => {
            const row = tbody.insertRow();
            const usagePercent = (sub.consumed / sub.enabled) * 100;
            let progressClass = 'green';
            if (usagePercent > 90) progressClass = 'red';
            else if (usagePercent > 70) progressClass = 'yellow';
            row.insertCell(0).innerHTML = `<strong>${sub.displayName}</strong>`;
            row.insertCell(1).innerHTML = `<span class="badge badge-info">${sub.skuId}</span>`;
            row.insertCell(2).innerHTML = sub.enabled.toLocaleString();
            row.insertCell(3).innerHTML = sub.consumed.toLocaleString();
            row.insertCell(4).innerHTML = `<strong style="color:#28a745">${sub.available.toLocaleString()}</strong>`;
            row.insertCell(5).innerHTML = `<div class="progress-bar"><div class="progress-fill ${progressClass}" style="width: ${usagePercent}%"></div></div><span style="font-size:10px;">${usagePercent.toFixed(1)}%</span>`;
        });
        document.getElementById('totalSKU').innerHTML = `<i class="fas fa-chart-line"></i> Total: ${subscriptions.length} SKU Active`;
    }
    
    function renderTable() {
        const filtered = getFilteredUsers();
        const tbody = document.getElementById('tableBody');
        tbody.innerHTML = '';
        filtered.forEach(user => {
            const row = tbody.insertRow();
            row.insertCell(0).innerHTML = `<strong>${user.name}</strong>`;
            row.insertCell(1).innerHTML = `<a href="mailto:${user.email}" style="color:#0078D4;">${user.email}</a>`;
            row.insertCell(2).innerHTML = user.department || '—';
            row.insertCell(3).innerHTML = user.is_guest ? '<span class="badge badge-info">Guest</span>' : '<span class="badge badge-primary">Internal</span>';
            row.insertCell(4).innerHTML = user.sign_blocked ? '<span class="badge badge-warning">Blocked</span>' : '<span class="badge badge-success">Active</span>';
            row.insertCell(5).innerHTML = user.last_sign_in || 'Never';
            row.insertCell(6).innerHTML = user.licenses.map(l => `<span class="badge badge-primary">${l}</span>`).join(' ') || '<span class="badge badge-warning">No License</span>';
            row.insertCell(7).innerHTML = `<strong>${user.license_count}</strong>`;
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
        if (searchTerm) filtered = filtered.filter(u => u.name.toLowerCase().includes(searchTerm) || u.email.toLowerCase().includes(searchTerm));
        return filtered;
    }
    
    function setFilter(filter) {
        currentFilter = filter;
        const btns = ['filterAll', 'filterInternal', 'filterGuest', 'filterLicensed', 'filterUnlicensed', 'filterBlockedE1', 'filterBlockedE3'];
        const mapping = {'filterAll':'all','filterInternal':'internal','filterGuest':'guest','filterLicensed':'licensed','filterUnlicensed':'unlicensed','filterBlockedE1':'blocked_e1','filterBlockedE3':'blocked_e3'};
        btns.forEach(btnId => {
            const btn = document.getElementById(btnId);
            if (mapping[btnId] === filter) btn.classList.add('active');
            else btn.classList.remove('active');
        });
        renderTable();
    }
    
    function showPage(page) {
        if (page === 'dashboard') {
            document.getElementById('dashboardPage').classList.remove('hidden');
            document.getElementById('billingPage').classList.add('hidden');
            document.getElementById('pageTitle').innerHTML = '<i class="fas fa-chart-pie"></i> Dashboard';
            document.getElementById('navDashboardBtn').classList.add('active');
            document.getElementById('navBillingBtn').classList.remove('active');
        } else {
            document.getElementById('dashboardPage').classList.add('hidden');
            document.getElementById('billingPage').classList.remove('hidden');
            document.getElementById('pageTitle').innerHTML = '<i class="fas fa-dollar-sign"></i> Billing & Subscriptions';
            document.getElementById('navDashboardBtn').classList.remove('active');
            document.getElementById('navBillingBtn').classList.add('active');
        }
    }
    
    function exportCSV() {
        const filtered = getFilteredUsers();
        let csv = "Name,Email,Department,User Type,Status,Last Sign In,Licenses,Count\\n";
        filtered.forEach(u => {
            csv += `"${u.name}","${u.email}","${u.department || ''}","${u.is_guest ? 'Guest' : 'Internal'}","${u.sign_blocked ? 'Blocked' : 'Active'}","${u.last_sign_in || 'Never'}","${u.licenses.join('; ')}",${u.license_count}\\n`;
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
    
    loadSidebarState();
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
        }
        h1 { font-size: 28px; font-weight: 700; margin-bottom: 8px; }
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
            font-size: 14px;
        }
        .features i { width: 24px; color: #0078D4; margin-right: 12px; }
        .btn-login {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            color: white;
            padding: 14px 32px;
            border-radius: 40px;
            font-size: 16px;
            font-weight: 600;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 10px;
        }
        .footer { margin-top: 32px; font-size: 11px; color: #999; }
    </style>
</head>
<body>
    <div class="login-card">
        <div class="logo-icon"><i class="fas fa-chart-line" style="color: white;"></i></div>
        <h1>License Monitor</h1>
        <p class="subtitle">Microsoft 365 License & Billing Monitoring</p>
        <div class="features">
            <li><i class="fas fa-building"></i> Internal & Guest users</li>
            <li><i class="fas fa-tag"></i> License: E1, E3, ME3, F3</li>
            <li><i class="fas fa-dollar-sign"></i> Billing & Subscription tracking</li>
            <li><i class="fas fa-ban"></i> Blocked users by license</li>
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
    
    # Ambil SKU license untuk billing
    skus_response = requests.get("https://graph.microsoft.com/v1.0/subscribedSkus", headers=headers)
    skus = skus_response.json().get("value", []) if skus_response.status_code == 200 else []
    sku_map = {}
    subscriptions = []
    
    for sku in skus:
        sku_id = sku.get("skuId")
        sku_name = sku.get("skuPartNumber", "Unknown")
        short_name = LICENSE_MAP.get(sku_name, sku_name)
        sku_map[sku_id] = short_name
        
        prepaid = sku.get("prepaidUnits", {})
        enabled = prepaid.get("enabled", 0)
        consumed = sku.get("consumedUnits", 0)
        
        if enabled > 0:
            subscriptions.append({
                "skuId": sku_name,
                "displayName": short_name,
                "enabled": enabled,
                "consumed": consumed,
                "available": enabled - consumed
            })
    
    # Proses data user
    processed_users = []
    license_stats = {}
    today = datetime.now()
    
    for user in users:
        email = user.get("userPrincipalName", "")
        is_guest = "#EXT#" in email or user.get("userType") == "Guest"
        
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
        
        last_sign_in_str = 'Never'
        sign_in_activity = user.get("signInActivity", {})
        last_sign_in_date = sign_in_activity.get("lastSignInDateTime")
        if last_sign_in_date:
            try:
                last_sign_in = datetime.fromisoformat(last_sign_in_date.replace('Z', '+00:00'))
                last_sign_in_str = last_sign_in.strftime('%Y-%m-%d')
            except:
                last_sign_in_str = last_sign_in_date[:10] if last_sign_in_date else 'Never'
        
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
            "last_sign_in": last_sign_in_str
        })
    
    internal_users = [u for u in processed_users if not u['is_guest']]
    guest_users = [u for u in processed_users if u['is_guest']]
    licensed_active = len([u for u in internal_users if u['license_count'] > 0 and not u['sign_blocked']])
    unlicensed_active = len([u for u in internal_users if u['license_count'] == 0 and not u['sign_blocked']])
    blocked_e1 = len([u for u in processed_users if u['sign_blocked'] and u['has_e1']])
    blocked_e3 = len([u for u in processed_users if u['sign_blocked'] and u['has_e3']])
    
    return jsonify({
        'users': processed_users,
        'license_stats': dict(sorted(license_stats.items(), key=lambda x: x[1], reverse=True)),
        'subscriptions': subscriptions,
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
