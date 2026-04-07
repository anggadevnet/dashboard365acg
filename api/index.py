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

# HTML Dashboard - UI Premium Modern
DASHBOARD_HTML = '''
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>M365 License Monitor | Lintasarta</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;14..32,400;14..32,500;14..32,600;14..32,700;14..32,800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: 'Inter', sans-serif;
            background: radial-gradient(circle at 20% 50%, #f5f7fc 0%, #eef2f8 100%);
            min-height: 100vh;
        }
        
        /* Animated Background */
        .bg-animation {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: -1;
            overflow: hidden;
        }
        .bg-animation::before {
            content: '';
            position: absolute;
            width: 200%;
            height: 200%;
            background: radial-gradient(circle, rgba(0,120,212,0.03) 0%, transparent 70%);
            animation: rotate 30s linear infinite;
        }
        @keyframes rotate {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }
        
        /* Layout */
        .app {
            display: flex;
            min-height: 100vh;
        }
        
        /* SIDEBAR - Glassmorphism */
        .sidebar {
            width: 280px;
            background: rgba(22, 33, 62, 0.95);
            backdrop-filter: blur(10px);
            color: white;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            position: fixed;
            height: 100vh;
            z-index: 100;
            overflow-y: auto;
            border-right: 1px solid rgba(255,255,255,0.1);
        }
        .sidebar.collapsed {
            width: 80px;
        }
        .sidebar.collapsed .sidebar-text,
        .sidebar.collapsed .sidebar-label,
        .sidebar.collapsed .stat-value-small,
        .sidebar.collapsed .stat-label-small {
            display: none;
        }
        .sidebar.collapsed .sidebar-item {
            justify-content: center;
            padding: 14px;
        }
        .sidebar.collapsed .sidebar-item .sidebar-icon {
            margin-right: 0;
            font-size: 20px;
        }
        .sidebar.collapsed .sidebar-header {
            padding: 20px 0;
            text-align: center;
        }
        .sidebar.collapsed .sidebar-logo {
            justify-content: center;
        }
        .sidebar.collapsed .stat-item {
            text-align: center;
            padding: 10px;
        }
        
        /* Custom Scrollbar */
        .sidebar::-webkit-scrollbar { width: 4px; }
        .sidebar::-webkit-scrollbar-track { background: rgba(255,255,255,0.1); }
        .sidebar::-webkit-scrollbar-thumb { background: #0078D4; border-radius: 4px; }
        
        /* Main Content */
        .main-content {
            flex: 1;
            margin-left: 280px;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .main-content.expanded {
            margin-left: 80px;
        }
        
        /* Toggle Button */
        .toggle-btn {
            position: fixed;
            left: 280px;
            top: 20px;
            z-index: 101;
            width: 36px;
            height: 36px;
            background: white;
            border: none;
            border-radius: 12px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            transition: all 0.3s;
            color: #1a1a2e;
        }
        .toggle-btn:hover {
            transform: scale(1.05);
            background: #0078D4;
            color: white;
        }
        .toggle-btn.collapsed {
            left: 80px;
        }
        
        /* Sidebar Header */
        .sidebar-header {
            padding: 24px 20px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            margin-bottom: 24px;
        }
        .sidebar-logo {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sidebar-logo-icon {
            width: 44px;
            height: 44px;
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
            box-shadow: 0 4px 12px rgba(0,120,212,0.3);
        }
        .sidebar-logo-text { font-weight: 800; font-size: 18px; letter-spacing: -0.5px; }
        .sidebar-logo-sub { font-size: 10px; opacity: 0.6; margin-top: 2px; }
        
        /* Sidebar Navigation */
        .sidebar-nav { padding: 0 16px; }
        .sidebar-item {
            display: flex;
            align-items: center;
            gap: 14px;
            padding: 12px 16px;
            margin: 6px 0;
            border-radius: 14px;
            cursor: pointer;
            transition: all 0.3s;
            color: rgba(255,255,255,0.7);
        }
        .sidebar-item:hover {
            background: rgba(255,255,255,0.1);
            color: white;
            transform: translateX(4px);
        }
        .sidebar-item.active {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            color: white;
            box-shadow: 0 4px 12px rgba(0,120,212,0.3);
        }
        .sidebar-icon { width: 24px; font-size: 18px; }
        .sidebar-text { font-size: 14px; font-weight: 500; }
        
        .sidebar-divider {
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
            margin: 20px 16px;
        }
        
        /* Sidebar Stats */
        .sidebar-stats { padding: 0 16px; margin-top: 20px; }
        .stat-item {
            background: rgba(255,255,255,0.05);
            border-radius: 14px;
            padding: 12px;
            margin-bottom: 10px;
            transition: all 0.3s;
        }
        .stat-item:hover {
            background: rgba(255,255,255,0.1);
            transform: translateX(4px);
        }
        .stat-label-small { font-size: 11px; opacity: 0.6; margin-bottom: 6px; letter-spacing: 0.5px; }
        .stat-value-small { font-size: 22px; font-weight: 800; line-height: 1.2; }
        
        /* Navbar Premium */
        .top-navbar {
            background: rgba(255,255,255,0.8);
            backdrop-filter: blur(20px);
            padding: 14px 28px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
            position: sticky;
            top: 0;
            z-index: 99;
            border-bottom: 1px solid rgba(0,0,0,0.05);
        }
        .page-title {
            font-size: 22px;
            font-weight: 800;
            background: linear-gradient(135deg, #1a1a2e, #0078D4);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
            letter-spacing: -0.5px;
        }
        .user-info {
            display: flex;
            align-items: center;
            gap: 16px;
            background: white;
            padding: 6px 20px 6px 16px;
            border-radius: 60px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
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
        }
        .logout-btn {
            background: none;
            border: none;
            color: #dc3545;
            cursor: pointer;
            font-size: 16px;
            transition: all 0.2s;
        }
        .logout-btn:hover { transform: scale(1.1); }
        
        /* Container */
        .container {
            padding: 28px 32px;
            max-width: 100%;
        }
        
        /* Stats Grid Premium */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 20px;
            margin-bottom: 28px;
        }
        .stat-card {
            background: white;
            border-radius: 24px;
            padding: 20px;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            border: 1px solid rgba(0,0,0,0.04);
            box-shadow: 0 4px 12px rgba(0,0,0,0.02);
            position: relative;
            overflow: hidden;
        }
        .stat-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 4px;
            background: linear-gradient(90deg, #0078D4, #00A4EF);
            transform: scaleX(0);
            transition: transform 0.3s;
        }
        .stat-card:hover::before {
            transform: scaleX(1);
        }
        .stat-card:hover {
            transform: translateY(-6px);
            box-shadow: 0 20px 35px -12px rgba(0,0,0,0.15);
            border-color: transparent;
        }
        .stat-icon {
            width: 48px;
            height: 48px;
            border-radius: 18px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 24px;
            margin-bottom: 16px;
            transition: all 0.3s;
        }
        .stat-card:hover .stat-icon {
            transform: scale(1.05);
        }
        .stat-value {
            font-size: 32px;
            font-weight: 800;
            margin-bottom: 6px;
            letter-spacing: -1px;
        }
        .stat-label {
            color: #666;
            font-size: 12px;
            font-weight: 500;
            letter-spacing: 0.3px;
        }
        
        /* Filter Bar Premium */
        .filter-bar {
            background: white;
            border-radius: 20px;
            padding: 12px 20px;
            margin-bottom: 24px;
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            border: 1px solid rgba(0,0,0,0.04);
        }
        .search-box {
            flex: 1;
            min-width: 260px;
            padding: 10px 18px;
            border: 1px solid #e0e0e0;
            border-radius: 40px;
            background: #fafbfc;
            font-size: 13px;
            transition: all 0.3s;
        }
        .search-box:focus-within {
            border-color: #0078D4;
            background: white;
            box-shadow: 0 0 0 3px rgba(0,120,212,0.1);
        }
        .search-box input {
            border: none;
            background: transparent;
            width: 90%;
            outline: none;
            font-size: 13px;
        }
        .filter-group {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        .filter-btn {
            padding: 8px 18px;
            border: none;
            border-radius: 40px;
            cursor: pointer;
            background: #f0f2f5;
            font-size: 12px;
            font-weight: 600;
            transition: all 0.2s;
            letter-spacing: 0.3px;
        }
        .filter-btn:hover {
            transform: translateY(-1px);
            background: #e0e0e0;
        }
        .filter-btn.active {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            color: white;
            box-shadow: 0 4px 12px rgba(0,120,212,0.3);
        }
        .filter-btn.warning.active {
            background: linear-gradient(135deg, #dc3545, #e4606d);
            box-shadow: 0 4px 12px rgba(220,53,69,0.3);
        }
        .export-btn {
            background: linear-gradient(135deg, #28a745, #20c997);
            color: white;
            padding: 8px 22px;
            border: none;
            border-radius: 40px;
            cursor: pointer;
            font-weight: 700;
            font-size: 12px;
            transition: all 0.3s;
        }
        .export-btn:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 14px rgba(40,167,69,0.3);
        }
        
        /* Billing Section */
        .billing-section {
            background: white;
            border-radius: 24px;
            margin-bottom: 24px;
            overflow: hidden;
            box-shadow: 0 2px 12px rgba(0,0,0,0.04);
            border: 1px solid rgba(0,0,0,0.04);
        }
        .billing-header {
            padding: 16px 24px;
            background: #fafbfc;
            border-bottom: 1px solid #eef2f6;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
        }
        .billing-title {
            font-size: 14px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 10px;
            color: #1a1a2e;
        }
        .billing-table-container {
            overflow-x: auto;
            max-height: 320px;
            overflow-y: auto;
        }
        
        /* User Section */
        .user-section {
            background: white;
            border-radius: 24px;
            overflow: hidden;
            box-shadow: 0 2px 12px rgba(0,0,0,0.04);
            border: 1px solid rgba(0,0,0,0.04);
        }
        .user-header {
            padding: 16px 24px;
            background: #fafbfc;
            border-bottom: 1px solid #eef2f6;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        /* Tables */
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
        }
        th {
            padding: 14px 16px;
            text-align: left;
            background: #f8f9fa;
            font-weight: 600;
            color: #333;
            border-bottom: 1px solid #eef2f6;
            position: sticky;
            top: 0;
        }
        td {
            padding: 12px 16px;
            border-bottom: 1px solid #f0f0f0;
        }
        tr {
            transition: all 0.2s;
        }
        tr:hover {
            background: linear-gradient(90deg, #f8f9ff, transparent);
        }
        
        /* Badges */
        .badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 30px;
            font-size: 10px;
            font-weight: 600;
            margin: 2px;
            letter-spacing: 0.3px;
        }
        .badge-primary { background: #e3f2fd; color: #0078D4; }
        .badge-success { background: #d4edda; color: #28a745; }
        .badge-warning { background: #f8d7da; color: #dc3545; }
        .badge-info { background: #d1ecf1; color: #17a2b8; }
        .badge-dark { background: #e9ecef; color: #6c757d; }
        
        .progress-bar {
            width: 80px;
            height: 4px;
            background: #e0e0e0;
            border-radius: 4px;
            overflow: hidden;
        }
        .progress-fill { height: 100%; border-radius: 4px; transition: width 0.3s; }
        .progress-fill.green { background: linear-gradient(90deg, #28a745, #20c997); }
        .progress-fill.yellow { background: linear-gradient(90deg, #ffc107, #ffda6a); }
        .progress-fill.red { background: linear-gradient(90deg, #dc3545, #e4606d); }
        
        /* Loading */
        .loading {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 80px;
        }
        .spinner {
            width: 48px;
            height: 48px;
            border: 3px solid #e0e0e0;
            border-top-color: #0078D4;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        
        .hidden { display: none; }
        
        /* Responsive */
        @media (max-width: 1300px) {
            .stats-grid { grid-template-columns: repeat(3, 1fr); gap: 16px; }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .sidebar { transform: translateX(-100%); }
            .sidebar.mobile-open { transform: translateX(0); }
            .main-content { margin-left: 0 !important; }
            .toggle-btn { left: 16px; }
            .container { padding: 16px; }
            .top-navbar { padding: 12px 16px; }
        }
    </style>
</head>
<body>
<div class="bg-animation"></div>
<div class="app">
    <button class="toggle-btn" id="toggleBtn" onclick="toggleSidebar()">
        <i class="fas fa-bars"></i>
    </button>
    
    <div class="sidebar" id="sidebar">
        <div class="sidebar-header">
            <div class="sidebar-logo">
                <div class="sidebar-logo-icon"><i class="fas fa-chart-line"></i></div>
                <div>
                    <div class="sidebar-logo-text">LicMonitor</div>
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
                <div class="sidebar-text">Billing</div>
            </div>
        </div>
        
        <div class="sidebar-divider"></div>
        
        <div class="sidebar-stats" id="sidebarStats">
            <div class="stat-item">
                <div class="stat-label-small">TOTAL USERS</div>
                <div class="stat-value-small" id="sidebarTotal">-</div>
            </div>
            <div class="stat-item">
                <div class="stat-label-small">LICENSED</div>
                <div class="stat-value-small" id="sidebarLicensed">-</div>
            </div>
            <div class="stat-item">
                <div class="stat-label-small">BLOCKED + E1</div>
                <div class="stat-value-small" id="sidebarBlockedE1">-</div>
            </div>
            <div class="stat-item">
                <div class="stat-label-small">BLOCKED + E3</div>
                <div class="stat-value-small" id="sidebarBlockedE3">-</div>
            </div>
        </div>
    </div>
    
    <div class="main-content" id="mainContent">
        <div class="top-navbar">
            <div class="page-title" id="pageTitle"><i class="fas fa-chart-pie"></i> Dashboard</div>
            <div class="user-info">
                <div class="user-avatar"><i class="fas fa-user"></i></div>
                <span style="font-weight: 500; font-size: 13px;">{{ user.name }}</span>
                <a href="/logout" class="logout-btn"><i class="fas fa-sign-out-alt"></i></a>
            </div>
        </div>
        
        <div class="container">
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <p style="margin-top: 20px; color: #666;">Loading data from Microsoft 365...</p>
            </div>
            
            <div id="dashboardPage">
                <div class="stats-grid" id="statsGrid"></div>
                
                <div class="filter-bar">
                    <div class="search-box">
                        <i class="fas fa-search" style="color: #999;"></i>
                        <input type="text" id="searchInput" placeholder="Search by name or email...">
                    </div>
                    <div class="filter-group">
                        <button id="filterAll" class="filter-btn active">All</button>
                        <button id="filterInternal" class="filter-btn">Internal</button>
                        <button id="filterGuest" class="filter-btn">Guest</button>
                        <button id="filterLicensed" class="filter-btn">Licensed</button>
                        <button id="filterUnlicensed" class="filter-btn">Unlicensed</button>
                        <button id="filterBlockedE1" class="filter-btn warning">Blocked+E1</button>
                        <button id="filterBlockedE3" class="filter-btn warning">Blocked+E3</button>
                    </div>
                    <button id="exportBtn" class="export-btn"><i class="fas fa-download"></i> Export CSV</button>
                </div>
                
                <div class="user-section">
                    <div class="user-header">
                        <div><i class="fas fa-users"></i> <strong>User Directory</strong></div>
                        <div><span id="rowCount">0</span> users</div>
                    </div>
                    <div style="overflow-x: auto; max-height: 420px; overflow-y: auto;">
                        <table id="userTable">
                            <thead>
                                <tr><th>Name</th><th>Email</th><th>Dept</th><th>Type</th><th>Status</th><th>Last Sign In</th><th>Licenses</th><th>Count</th></tr>
                            </thead>
                            <tbody id="tableBody"></tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div id="billingPage" class="hidden">
                <div class="billing-section">
                    <div class="billing-header">
                        <div class="billing-title"><i class="fas fa-tags"></i> Active Subscriptions</div>
                        <div id="totalSKU"></div>
                    </div>
                    <div class="billing-table-container" style="max-height: 520px;">
                        <table id="billingTable">
                            <thead><tr><th>License</th><th>SKU</th><th>Total</th><th>Used</th><th>Available</th><th>Usage</th></tr></thead>
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
    
    function toggleSidebar() {
        const sidebar = document.getElementById('sidebar');
        const mainContent = document.getElementById('mainContent');
        const toggleBtn = document.getElementById('toggleBtn');
        sidebar.classList.toggle('collapsed');
        mainContent.classList.toggle('expanded');
        toggleBtn.classList.toggle('collapsed');
        localStorage.setItem('sidebarCollapsed', sidebar.classList.contains('collapsed'));
    }
    
    function loadSidebarState() {
        if (localStorage.getItem('sidebarCollapsed') === 'true') {
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
                <div class="stat-icon" style="background:linear-gradient(135deg,#e3f2fd,#e8f4fd)"><i class="fas fa-building" style="color:#0078D4"></i></div>
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
            row.insertCell(1).innerHTML = `<span class="badge badge-dark">${sub.skuId.substring(0, 20)}</span>`;
            row.insertCell(2).innerHTML = `<strong>${sub.enabled.toLocaleString()}</strong>`;
            row.insertCell(3).innerHTML = sub.consumed.toLocaleString();
            row.insertCell(4).innerHTML = `<strong style="color:#28a745">${sub.available.toLocaleString()}</strong>`;
            row.insertCell(5).innerHTML = `<div class="progress-bar"><div class="progress-fill ${progressClass}" style="width: ${usagePercent}%"></div></div><span style="font-size: 9px; margin-top: 4px; display: block;">${usagePercent.toFixed(0)}% used</span>`;
        });
        document.getElementById('totalSKU').innerHTML = `<i class="fas fa-chart-line"></i> ${subscriptions.length} Active SKUs`;
    }
    
    function renderTable() {
        const filtered = getFilteredUsers();
        const tbody = document.getElementById('tableBody');
        tbody.innerHTML = '';
        document.getElementById('rowCount').innerHTML = filtered.length;
        filtered.forEach(user => {
            const row = tbody.insertRow();
            row.insertCell(0).innerHTML = `<strong>${user.name}</strong>`;
            row.insertCell(1).innerHTML = `<a href="mailto:${user.email}" style="color:#0078D4; text-decoration:none;">${user.email}</a>`;
            row.insertCell(2).innerHTML = user.department || '—';
            row.insertCell(3).innerHTML = user.is_guest ? '<span class="badge badge-info">Guest</span>' : '<span class="badge badge-primary">Internal</span>';
            row.insertCell(4).innerHTML = user.sign_blocked ? '<span class="badge badge-warning">Blocked</span>' : '<span class="badge badge-success">Active</span>';
            row.insertCell(5).innerHTML = user.last_sign_in || '<span class="badge badge-dark">Never</span>';
            row.insertCell(6).innerHTML = user.licenses.map(l => `<span class="badge badge-primary">${l}</span>`).join(' ') || '<span class="badge badge-warning">—</span>';
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
            document.getElementById('pageTitle').innerHTML = '<i class="fas fa-dollar-sign"></i> Billing';
            document.getElementById('navDashboardBtn').classList.remove('active');
            document.getElementById('navBillingBtn').classList.add('active');
        }
    }
    
    function exportCSV() {
        const filtered = getFilteredUsers();
        let csv = "Name,Email,Department,Type,Status,Last Sign In,Licenses,Count\\n";
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
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif;
            background: radial-gradient(circle at 30% 20%, #1a1a2e, #0f0f1a);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .login-card {
            background: rgba(255,255,255,0.98);
            border-radius: 40px;
            padding: 48px;
            max-width: 520px;
            width: 90%;
            text-align: center;
            box-shadow: 0 40px 60px rgba(0,0,0,0.3);
            backdrop-filter: blur(10px);
            animation: fadeInUp 0.6s ease;
        }
        @keyframes fadeInUp {
            from { opacity: 0; transform: translateY(30px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .logo-icon {
            width: 80px;
            height: 80px;
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 24px;
            font-size: 36px;
            box-shadow: 0 12px 24px rgba(0,120,212,0.3);
        }
        h1 { font-size: 32px; font-weight: 800; margin-bottom: 8px; background: linear-gradient(135deg, #1a1a2e, #0078D4); -webkit-background-clip: text; background-clip: text; color: transparent; }
        .subtitle { color: #666; margin-bottom: 32px; font-size: 14px; }
        .features {
            text-align: left;
            margin: 32px 0;
            background: #f8f9fa;
            padding: 24px 28px;
            border-radius: 28px;
        }
        .features li {
            list-style: none;
            margin: 14px 0;
            font-size: 14px;
            font-weight: 500;
            display: flex;
            align-items: center;
        }
        .features i { width: 28px; color: #0078D4; font-size: 18px; margin-right: 12px; }
        .btn-login {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            color: white;
            padding: 14px 36px;
            border-radius: 60px;
            font-size: 16px;
            font-weight: 700;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 12px;
            transition: all 0.3s;
            box-shadow: 0 8px 20px rgba(0,120,212,0.3);
        }
        .btn-login:hover {
            transform: translateY(-3px);
            box-shadow: 0 15px 30px rgba(0,120,212,0.4);
        }
        .footer { margin-top: 32px; font-size: 11px; color: #999; }
    </style>
</head>
<body>
    <div class="login-card">
        <div class="logo-icon"><i class="fas fa-chart-line" style="color: white;"></i></div>
        <h1>LicMonitor</h1>
        <p class="subtitle">Microsoft 365 License & Billing Intelligence</p>
        <div class="features">
            <li><i class="fas fa-building"></i> Internal & Guest Users</li>
            <li><i class="fas fa-tag"></i> E1, E3, ME3 License Tracking</li>
            <li><i class="fas fa-dollar-sign"></i> Real-time Billing Overview</li>
            <li><i class="fas fa-ban"></i> Blocked Users by License Type</li>
            <li><i class="fas fa-chart-line"></i> Export Reports to CSV</li>
        </div>
        <a href="/login" class="btn-login"><i class="fab fa-microsoft"></i> Login with Microsoft 365</a>
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
