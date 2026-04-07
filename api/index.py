import os
import requests
import msal
from flask import Flask, request, redirect, session, jsonify, render_template_string
from datetime import datetime

app = Flask(__name__)
app.secret_key = "rahasia-banget-2024"

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "User.Read.All", "Organization.Read.All"]

# HTML sederhana untuk testing
LOGIN_HTML = '''
<!DOCTYPE html>
<html>
<head><title>M365 License Monitor</title></head>
<body style="font-family: Arial; text-align: center; padding: 50px;">
    <h1>📊 M365 License Monitor</h1>
    <p>Monitoring lisensi pengguna Microsoft 365</p>
    <a href="/login" style="background: #0078D4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px;">🔐 Login dengan Microsoft 365</a>
</body>
</html>
'''

DASHBOARD_HTML = '''
<!DOCTYPE html>
<html>
<head>
    <title>Dashboard</title>
    <style>
        body { font-family: Arial; padding: 20px; }
        .stats { display: flex; gap: 20px; margin-bottom: 30px; }
        .card { background: #f0f2f5; padding: 20px; border-radius: 10px; text-align: center; flex: 1; }
        .value { font-size: 36px; font-weight: bold; color: #0078D4; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #f0f2f5; }
        .logout { background: #dc3545; color: white; padding: 8px 16px; text-decoration: none; border-radius: 6px; float: right; }
        .badge { background: #0078D4; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; margin: 2px; display: inline-block; }
    </style>
</head>
<body>
    <a href="/logout" class="logout">🚪 Logout</a>
    <h1>📊 M365 License Monitor</h1>
    <p>Welcome, {{ name }} ({{ email }})</p>
    <div class="stats">
        <div class="card"><div>Total Users</div><div class="value">{{ total_users }}</div></div>
        <div class="card"><div>Licensed Users</div><div class="value">{{ licensed_users }}</div></div>
        <div class="card"><div>Unlicensed Users</div><div class="value">{{ unlicensed_users }}</div></div>
        <div class="card"><div>Coverage</div><div class="value">{{ coverage }}%</div></div>
    </div>
    <h3>📋 Licensed Users</h3>
    <table>
        <thead><tr><th>Name</th><th>Email</th><th>Department</th><th>Licenses</th></tr></thead>
        <tbody>
        {% for user in licensed_list %}
        <tr>
            <td>{{ user.name }}</td>
            <td>{{ user.email }}</td>
            <td>{{ user.department }}</td>
            <td>{% for lic in user.licenses %}<span class="badge">{{ lic }}</span> {% endfor %}</td>
        </tr>
        {% endfor %}
        </tbody>
    </table>
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
        # Simpan token di session
        session['access_token'] = result['access_token']
        
        # Ambil info user
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
    
    token = session.get('access_token')
    headers = {"Authorization": f"Bearer {token}"}
    
    # Ambil semua user
    users = []
    url = "https://graph.microsoft.com/v1.0/users?$select=displayName,userPrincipalName,assignedLicenses,department&$top=999"
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
    
    # Buat mapping SKU ID ke nama
    sku_map = {sku.get("skuId"): sku.get("skuPartNumber", "Unknown") for sku in skus}
    
    # Proses data
    licensed_users = []
    unlicensed_count = 0
    
    for user in users:
        licenses = user.get("assignedLicenses", [])
        license_names = []
        for lic in licenses:
            sku_id = lic.get("skuId")
            if sku_id and sku_id in sku_map:
                license_names.append(sku_map[sku_id])
        
        user_data = {
            "name": user.get("displayName", "N/A"),
            "email": user.get("userPrincipalName", "N/A"),
            "department": user.get("department", "N/A"),
            "licenses": license_names
        }
        
        if license_names:
            licensed_users.append(user_data)
        else:
            unlicensed_count += 1
    
    return render_template_string(DASHBOARD_HTML,
        name=session['user']['name'],
        email=session['user']['email'],
        total_users=len(users),
        licensed_users=len(licensed_users),
        unlicensed_users=unlicensed_count,
        coverage=round(len(licensed_users)/len(users)*100, 1) if users else 0,
        licensed_list=licensed_users[:50]
    )

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
