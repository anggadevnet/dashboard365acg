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
SCOPES = ["User.Read", "User.Read.All", "Organization.Read.All", "AuditLog.Read.All", "Group.Read.All", "GroupMember.Read.All"]

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

# HTML Dashboard dengan Groups Page (sudah diperbaiki)
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
            background: #f0f4f8;
            min-height: 100vh;
        }
        
        /* Top Navbar */
        .navbar {
            background: white;
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
        }
        .logo {
            display: flex;
            align-items: center;
            gap: 12px;
            cursor: pointer;
        }
        .logo-icon {
            width: 42px;
            height: 42px;
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
            box-shadow: 0 6px 14px rgba(0,120,212,0.25);
        }
        .logo-text {
            font-size: 20px;
            font-weight: 800;
            background: linear-gradient(135deg, #1a1a2e, #0078D4);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
            letter-spacing: -0.5px;
        }
        .logo-sub {
            font-size: 10px;
            color: #666;
            margin-top: 2px;
        }
        
        /* Tab Navigation */
        .tab-nav {
            display: flex;
            gap: 8px;
            background: #f1f3f7;
            padding: 6px;
            border-radius: 60px;
        }
        .tab-btn {
            padding: 10px 28px;
            border: none;
            border-radius: 40px;
            cursor: pointer;
            background: transparent;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
            color: #555;
        }
        .tab-btn.active {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            color: white;
            box-shadow: 0 4px 12px rgba(0,120,212,0.25);
        }
        
        .user-info {
            display: flex;
            align-items: center;
            gap: 16px;
            background: #f1f3f7;
            padding: 6px 20px 6px 16px;
            border-radius: 60px;
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
            max-width: 1600px;
            margin: 0 auto;
            padding: 28px 32px;
        }
        
        /* Welcome Banner */
        .welcome-banner {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            border-radius: 28px;
            padding: 28px 32px;
            margin-bottom: 32px;
            color: white;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            box-shadow: 0 12px 28px rgba(0,120,212,0.2);
        }
        .welcome-title {
            font-size: 24px;
            font-weight: 800;
            margin-bottom: 6px;
            letter-spacing: -0.5px;
        }
        .update-badge {
            background: rgba(255,255,255,0.2);
            padding: 8px 18px;
            border-radius: 40px;
            font-size: 13px;
            font-weight: 500;
        }
        
        /* Stats Grid */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 20px;
            margin-bottom: 32px;
        }
        .stat-card {
            background: white;
            border-radius: 24px;
            padding: 20px;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            border: 1px solid rgba(0,0,0,0.04);
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
            position: relative;
            overflow: hidden;
        }
        .stat-card::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 0;
            width: 100%;
            height: 3px;
            background: linear-gradient(90deg, #0078D4, #00A4EF);
            transform: scaleX(0);
            transition: transform 0.3s;
        }
        .stat-card:hover::after {
            transform: scaleX(1);
        }
        .stat-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 20px 30px -12px rgba(0,0,0,0.12);
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
        }
        .stat-value {
            font-size: 32px;
            font-weight: 800;
            margin-bottom: 4px;
            letter-spacing: -1px;
        }
        .stat-label {
            color: #666;
            font-size: 12px;
            font-weight: 600;
            letter-spacing: 0.3px;
        }
        
        /* Filter Bar */
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
        }
        .filter-btn.active {
            background: linear-gradient(135deg, #0078D4, #00A4EF);
            color: white;
            box-shadow: 0 4px 12px rgba(0,120,212,0.25);
        }
        .filter-btn.warning.active {
            background: linear-gradient(135deg, #dc3545, #e4606d);
            box-shadow: 0 4px 12px rgba(220,53,69,0.25);
        }
        .export-btn {
            background: linear-gradient(135deg, #28a745, #20c997);
            color: white;
            padding: 8px 24px;
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
        
        /* Groups Layout */
        .groups-layout {
            display: grid;
            grid-template-columns: 350px 1fr;
            gap: 24px;
            margin-bottom: 24px;
        }
        .groups-list {
            background: white;
            border-radius: 24px;
            overflow: hidden;
            border: 1px solid rgba(0,0,0,0.04);
            max-height: 600px;
            overflow-y: auto;
        }
        .groups-header {
            padding: 16px 20px;
            background: #fafbfc;
            border-bottom: 1px solid #eef2f6;
            font-weight: 700;
            position: sticky;
            top: 0;
        }
        .group-item {
            padding: 14px 20px;
            border-bottom: 1px solid #f0f0f0;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .group-item:hover {
            background: #f8f9ff;
        }
        .group-item.active {
            background: linear-gradient(90deg, #e3f2fd, transparent);
            border-left: 3px solid #0078D4;
        }
        .group-name {
            font-weight: 600;
            font-size: 14px;
        }
        .group-member-count {
            font-size: 12px;
            color: #666;
            background: #f0f2f5;
            padding: 2px 10px;
            border-radius: 20px;
        }
        .group-search {
            padding: 12px 16px;
            border-bottom: 1px solid #eef2f6;
        }
        .group-search input {
            width: 100%;
            padding: 8px 12px;
            border: 1px solid #e0e0e0;
            border-radius: 30px;
            font-size: 12px;
        }
        
        .members-panel {
            background: white;
            border-radius: 24px;
            border: 1px solid rgba(0,0,0,0.04);
            overflow: hidden;
            max-height: 600px;
            display: flex;
            flex-direction: column;
        }
        .members-header {
            padding: 16px 20px;
            background: #fafbfc;
            border-bottom: 1px solid #eef2f6;
            font-weight: 700;
        }
        .members-list {
            flex: 1;
            overflow-y: auto;
            padding: 0;
        }
        .member-item {
            padding: 12px 20px;
            border-bottom: 1px solid #f0f0f0;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .member-avatar {
            width: 32px;
            height: 32px;
            background: #e3f2fd;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #0078D4;
        }
        .member-info {
            flex: 1;
        }
        .member-name {
            font-weight: 600;
            font-size: 13px;
        }
        .member-email {
            font-size: 11px;
            color: #666;
        }
        .loading-members {
            padding: 40px;
            text-align: center;
            color: #666;
        }
        
        /* Billing Section */
        .billing-section {
            background: white;
            border-radius: 24px;
            overflow: hidden;
            box-shadow: 0 2px 12px rgba(0,0,0,0.04);
            border: 1px solid rgba(0,0,0,0.04);
        }
        .billing-header {
            padding: 18px 24px;
            background: #fafbfc;
            border-bottom: 1px solid #eef2f6;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
        }
        .billing-title {
            font-size: 15px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .billing-table-container {
            overflow-x: auto;
            max-height: 360px;
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
            padding: 18px 24px;
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
            background: #f8f9ff;
        }
        
        /* Badges */
        .badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 30px;
            font-size: 10px;
            font-weight: 600;
            margin: 2px;
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
        .progress-fill { height: 100%; border-radius: 4px; }
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
        @media (max-width: 1000px) {
            .groups-layout { grid-template-columns: 1fr; }
            .stats-grid { grid-template-columns: repeat(3, 1fr); gap: 16px; }
        }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: repeat(2, 1fr); }
            .container { padding: 16px; }
            .navbar { padding: 12px 16px; }
            .tab-nav { order: 3; width: 100%; justify-content: center; }
            .welcome-banner { flex-direction: column; text-align: center; gap: 12px; }
        }
    </style>
</head>
<body>
    <!-- Top Navbar -->
    <div class="navbar">
        <div class="logo" onclick="showPage('dashboard')">
            <div class="logo-icon"><i class="fas fa-chart-line"></i></div>
            <div>
                <div class="logo-text">LicMonitor</div>
                <div class="logo-sub">Lintasarta</div>
            </div>
        </div>
        
        <div class="tab-nav">
            <button class="tab-btn active" id="tabDashboard" onclick="showPage('dashboard')"><i class="fas fa-tachometer-alt"></i> Dashboard</button>
            <button class="tab-btn" id="tabGroups" onclick="showPage('groups')"><i class="fas fa-users"></i> Groups</button>
            <button class="tab-btn" id="tabBilling" onclick="showPage('billing')"><i class="fas fa-dollar-sign"></i> Billing</button>
        </div>
        
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
        
        <!-- Dashboard Page -->
        <div id="dashboardPage">
            <div class="welcome-banner">
                <div>
                    <div class="welcome-title"><i class="fas fa-chart-pie"></i> License Overview</div>
                    <div style="opacity: 0.9; font-size: 14px;">Real-time monitoring user licenses & activity</div>
                </div>
                <div class="update-badge"><i class="fas fa-sync-alt"></i> Auto refresh every 5 minutes</div>
            </div>
            
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
                <div style="overflow-x: auto; max-height: 450px; overflow-y: auto;">
                    <table id="userTable">
                        <thead>
                            <tr><th>Name</th><th>Email</th><th>Dept</th><th>Type</th><th>Status</th><th>Last Sign In</th><th>Licenses</th><th>Count</th></tr>
                        </thead>
                        <tbody id="tableBody"></tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- Groups Page -->
        <div id="groupsPage" class="hidden">
            <div class="welcome-banner">
                <div>
                    <div class="welcome-title"><i class="fas fa-users"></i> Distribution Groups</div>
                    <div style="opacity: 0.9; font-size: 14px;">View group members and manage distribution lists</div>
                </div>
                <div class="update-badge"><i class="fas fa-sync-alt"></i> Click group to view members</div>
            </div>
            
            <div class="groups-layout">
                <!-- Groups List -->
                <div class="groups-list">
                    <div class="groups-header"><i class="fas fa-list"></i> All Groups</div>
                    <div class="group-search">
                        <input type="text" id="groupSearchInput" placeholder="🔍 Search group...">
                    </div>
                    <div id="groupsListContainer">
                        <div style="padding: 20px; text-align: center; color: #666;">Loading groups...</div>
                    </div>
                </div>
                
                <!-- Members Panel -->
                <div class="members-panel">
                    <div class="members-header" id="selectedGroupTitle">
                        <i class="fas fa-users"></i> Select a group to view members
                    </div>
                    <div class="members-list" id="membersListContainer">
                        <div class="loading-members">👈 Click on a group to see its members</div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Billing Page -->
        <div id="billingPage" class="hidden">
            <div class="welcome-banner">
                <div>
                    <div class="welcome-title"><i class="fas fa-dollar-sign"></i> Billing & Subscriptions</div>
                    <div style="opacity: 0.9; font-size: 14px;">License usage and availability tracking</div>
                </div>
                <div class="update-badge"><i class="fas fa-chart-line"></i> Real-time data</div>
            </div>
            
            <div class="billing-section">
                <div class="billing-header">
                    <div class="billing-title"><i class="fas fa-tags"></i> Active Subscriptions</div>
                    <div id="totalSKU"></div>
                </div>
                <div class="billing-table-container" style="max-height: 550px;">
                    <table id="billingTable">
                        <thead>
                            <tr><th>License Name</th><th>SKU Code</th><th>Total</th><th>Used</th><th>Available</th><th>Usage</th></tr>
                        </thead>
                        <tbody id="billingBody"></tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <script>
        let allUsers = [];
        let licenseStats = {};
        let subscriptions = [];
        let currentFilter = 'all';
        let allGroups = [];
        let selectedGroupId = null;
        
        async function loadData() {
            const res = await fetch('/api/license-data');
            const data = await res.json();
            if(data.error){ alert('Session expired'); window.location='/logout'; return; }
            allUsers = data.users;
            licenseStats = data.license_stats;
            subscriptions = data.subscriptions;
            updateStats(data.summary);
            renderBillingTable();
            renderTable();
            document.getElementById('loading').style.display = 'none';
            document.getElementById('dashboardPage').classList.remove('hidden');
        }
        
        // ============ PERBAIKAN: FUNGSI GROUPS YANG SUDAH DIPERBAIKI ============
        async function loadGroups() {
            const container = document.getElementById('groupsListContainer');
            container.innerHTML = '<div style="padding: 20px; text-align: center; color: #666;"><div class="spinner" style="width:30px;height:30px;margin:0 auto 12px;"></div><p>Loading groups...</p></div>';
            
            try {
                const res = await fetch('/api/groups');
                const data = await res.json();
                if(data.error) {
                    container.innerHTML = '<div style="padding: 20px; text-align: center; color: red;">❌ Failed to load groups</div>';
                    return;
                }
                allGroups = data.groups || [];
                renderGroupsList();
                
                // LANGSUNG PILIH GRUP PERTAMA OTOMATIS
                if (allGroups.length > 0 && !selectedGroupId) {
                    const firstGroup = allGroups[0];
                    selectGroup(firstGroup.id, firstGroup.displayName);
                }
            } catch(e) {
                console.error(e);
                container.innerHTML = '<div style="padding: 20px; text-align: center; color: red;">❌ Network error</div>';
            }
        }
        
        function renderGroupsList() {
            const container = document.getElementById('groupsListContainer');
            const searchTerm = document.getElementById('groupSearchInput').value.toLowerCase();
            
            let filtered = allGroups;
            if (searchTerm) {
                filtered = allGroups.filter(g => g.displayName.toLowerCase().includes(searchTerm));
            }
            
            if (filtered.length === 0) {
                container.innerHTML = '<div style="padding: 40px; text-align: center; color: #999;"><i class="fas fa-folder-open"></i><p style="margin-top: 8px;">No groups found</p></div>';
                return;
            }
            
            container.innerHTML = filtered.map(group => `
                <div class="group-item ${selectedGroupId === group.id ? 'active' : ''}" onclick="selectGroup('${group.id}', '${escapeHtml(group.displayName).replace(/'/g, "\\'")}')">
                    <div class="group-name"><i class="fas fa-envelope"></i> ${escapeHtml(group.displayName)}</div>
                    <div class="group-member-count">👥 ${group.memberCount || 0}</div>
                </div>
            `).join('');
        }
        
        async function selectGroup(groupId, groupName) {
            selectedGroupId = groupId;
            renderGroupsList();
            
            document.getElementById('selectedGroupTitle').innerHTML = `<i class="fas fa-users"></i> ${escapeHtml(groupName)} - Members`;
            document.getElementById('membersListContainer').innerHTML = '<div class="loading-members"><div class="spinner" style="width:30px;height:30px;margin:0 auto 12px;"></div><p>Loading members...</p></div>';
            
            try {
                const res = await fetch(`/api/groups/${groupId}/members`);
                const data = await res.json();
                if(data.error) {
                    document.getElementById('membersListContainer').innerHTML = `<div class="loading-members">❌ Error loading members</div>`;
                    return;
                }
                renderMembersList(data.members);
            } catch(e) {
                document.getElementById('membersListContainer').innerHTML = `<div class="loading-members">❌ Network error</div>`;
            }
        }
        
        function renderMembersList(members) {
            if (!members || members.length === 0) {
                document.getElementById('membersListContainer').innerHTML = '<div class="loading-members">📭 No members in this group</div>';
                return;
            }
            
            document.getElementById('membersListContainer').innerHTML = members.map(member => `
                <div class="member-item">
                    <div class="member-avatar"><i class="fas fa-user"></i></div>
                    <div class="member-info">
                        <div class="member-name">${escapeHtml(member.displayName || member.userPrincipalName || 'Unknown')}</div>
                        <div class="member-email">${escapeHtml(member.userPrincipalName || member.mail || 'No email')}</div>
                    </div>
                    <div><span class="badge badge-primary">${member.userType === 'Guest' ? 'Guest' : 'Member'}</span></div>
                </div>
            `).join('');
        }
        
        function escapeHtml(text) {
            if (!text) return '';
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }
        
        function updateStats(summary) {
            document.getElementById('statsGrid').innerHTML = `
                <div class="stat-card" onclick="setFilter('internal')">
                    <div class="stat-icon" style="background:linear-gradient(135deg,#e3f2fd,#e8f4fd)"><i class="fas fa-building" style="color:#0078D4"></i></div>
                    <div class="stat-value">${summary.internal_users}</div>
                    <div class="stat-label">Internal Users</div>
                </div>
                <div class="stat-card" onclick="setFilter('guest')">
                    <div class="stat-icon" style="background:#f8f9fa"><i class="fas fa-globe" style="color:#6c757d"></i></div>
                    <div class="stat-value">${summary.guest_users}</div>
                    <div class="stat-label">Guest Users</div>
                </div>
                <div class="stat-card" onclick="setFilter('licensed')">
                    <div class="stat-icon" style="background:#d4edda"><i class="fas fa-check-circle" style="color:#28a745"></i></div>
                    <div class="stat-value">${summary.licensed_users}</div>
                    <div class="stat-label">Licensed</div>
                </div>
                <div class="stat-card" onclick="setFilter('unlicensed')">
                    <div class="stat-icon" style="background:#fff3cd"><i class="fas fa-times-circle" style="color:#ffc107"></i></div>
                    <div class="stat-value">${summary.unlicensed_users}</div>
                    <div class="stat-label">Unlicensed</div>
                </div>
                <div class="stat-card" onclick="setFilter('blocked_e1')">
                    <div class="stat-icon" style="background:#f8d7da"><i class="fas fa-ban" style="color:#dc3545"></i></div>
                    <div class="stat-value">${summary.blocked_e1}</div>
                    <div class="stat-label">Blocked + E1</div>
                </div>
                <div class="stat-card" onclick="setFilter('blocked_e3')">
                    <div class="stat-icon" style="background:#f8d7da"><i class="fas fa-ban" style="color:#dc3545"></i></div>
                    <div class="stat-value">${summary.blocked_e3}</div>
                    <div class="stat-label">Blocked + E3</div>
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
        
        // ============ PERBAIKAN: FUNGSI SHOW PAGE ============
        function showPage(page) {
            if (page === 'dashboard') {
                document.getElementById('dashboardPage').classList.remove('hidden');
                document.getElementById('groupsPage').classList.add('hidden');
                document.getElementById('billingPage').classList.add('hidden');
                document.getElementById('tabDashboard').classList.add('active');
                document.getElementById('tabGroups').classList.remove('active');
                document.getElementById('tabBilling').classList.remove('active');
            } else if (page === 'groups') {
                document.getElementById('dashboardPage').classList.add('hidden');
                document.getElementById('groupsPage').classList.remove('hidden');
                document.getElementById('billingPage').classList.add('hidden');
                document.getElementById('tabDashboard').classList.remove('active');
                document.getElementById('tabGroups').classList.add('active');
                document.getElementById('tabBilling').classList.remove('active');
                
                // LOAD GROUPS SETIAP BUKA HALAMAN GROUPS
                if (allGroups.length === 0) {
                    loadGroups();
                } else {
                    renderGroupsList();
                    if (allGroups.length > 0 && !selectedGroupId) {
                        selectGroup(allGroups[0].id, allGroups[0].displayName);
                    }
                }
            } else {
                document.getElementById('dashboardPage').classList.add('hidden');
                document.getElementById('groupsPage').classList.add('hidden');
                document.getElementById('billingPage').classList.remove('hidden');
                document.getElementById('tabDashboard').classList.remove('active');
                document.getElementById('tabGroups').classList.remove('active');
                document.getElementById('tabBilling').classList.add('active');
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
        document.getElementById('groupSearchInput').addEventListener('keyup', () => renderGroupsList());
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
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
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
            border-radius: 40px;
            padding: 48px;
            max-width: 520px;
            width: 90%;
            text-align: center;
            box-shadow: 0 40px 60px rgba(0,0,0,0.2);
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
            <li><i class="fas fa-users"></i> Distribution Group Management</li>
            <li><i class="fas fa-dollar-sign"></i> Real-time Billing Overview</li>
            <li><i class="fas fa-ban"></i> Blocked Users by License Type</li>
            <li><i class="fas fa-download"></i> Export Reports to CSV</li>
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

@app.route('/api/groups')
def api_groups():
    token = session.get('access_token')
    if not token:
        return jsonify({'error': 'Unauthorized'}), 401
    
    headers = {"Authorization": f"Bearer {token}"}
    
    groups = []
    url = "https://graph.microsoft.com/v1.0/groups?$filter=mailEnabled eq true or groupTypes/any(c:c eq 'Unified')&$select=id,displayName,mail,groupTypes&$top=999"
    
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            for group in data.get("value", []):
                # Hitung jumlah member
                member_count = 0
                member_url = f"https://graph.microsoft.com/v1.0/groups/{group.get('id')}/members/$count"
                count_resp = requests.get(member_url, headers={**headers, "ConsistencyLevel": "eventual"})
                if count_resp.status_code == 200:
                    try:
                        member_count = int(count_resp.text)
                    except:
                        member_count = 0
                
                groups.append({
                    "id": group.get("id"),
                    "displayName": group.get("displayName", "N/A"),
                    "mail": group.get("mail", ""),
                    "groupType": "Unified" if "Unified" in group.get("groupTypes", []) else "Distribution",
                    "memberCount": member_count
                })
            url = data.get("@odata.nextLink")
        else:
            break
    
    # Urutkan berdasarkan nama
    groups.sort(key=lambda x: x['displayName'].lower())
    return jsonify({'groups': groups, 'total': len(groups)})

@app.route('/api/groups/<group_id>/members')
def api_group_members(group_id):
    token = session.get('access_token')
    if not token:
        return jsonify({'error': 'Unauthorized'}), 401
    
    headers = {"Authorization": f"Bearer {token}"}
    
    members = []
    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members?$select=id,displayName,userPrincipalName,mail,userType&$top=999"
    
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            for member in data.get("value", []):
                members.append({
                    "id": member.get("id"),
                    "displayName": member.get("displayName", "N/A"),
                    "userPrincipalName": member.get("userPrincipalName", ""),
                    "mail": member.get("mail", ""),
                    "userType": member.get("userType", "Member")
                })
            url = data.get("@odata.nextLink")
        else:
            break
    
    return jsonify({'members': members, 'count': len(members)})

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
