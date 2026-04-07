import os
import requests
import msal
from flask import Flask, request, redirect, session, render_template, jsonify
from functools import wraps
from datetime import datetime

app = Flask(__name__)
app.secret_key = "ganti-ini-dengan-string-random-2024"

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI", "http://localhost:5000/callback")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "User.Read.All", "Organization.Read.All"]

def get_msal_app():
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY
    )

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            return redirect('/login')
        return f(*args, **kwargs)
    return decorated_function

def get_token_from_session():
    token_data = session.get('token_data', {})
    if not token_data:
        return None
    expires_at = token_data.get('expires_at', 0)
    if datetime.now().timestamp() > expires_at - 300:
        app_msal = get_msal_app()
        result = app_msal.acquire_token_by_refresh_token(
            token_data['refresh_token'],
            scopes=SCOPES
        )
        if 'access_token' in result:
            session['token_data'] = {
                'access_token': result['access_token'],
                'refresh_token': result.get('refresh_token', token_data['refresh_token']),
                'expires_at': datetime.now().timestamp() + result['expires_in']
            }
            return result['access_token']
        return None
    return token_data.get('access_token')

def get_all_users(token):
    headers = {"Authorization": f"Bearer {token}"}
    users = []
    url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,assignedLicenses,userType,department&$top=999"
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            users.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        else:
            break
    return users

def get_subscribed_skus(token):
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get("https://graph.microsoft.com/v1.0/subscribedSkus", headers=headers)
    if response.status_code == 200:
        return response.json().get("value", [])
    return []

def get_license_details(sku_id, skus):
    for sku in skus:
        if sku.get("skuId") == sku_id:
            return sku.get("skuPartNumber", "Unknown")
    return "Unknown"

def analyze_licenses(users, skus):
    license_stats = {}
    licensed_users = []
    unlicensed_users = []
    for user in users:
        user_licenses = user.get("assignedLicenses", [])
        license_names = []
        for lic in user_licenses:
            sku_id = lic.get("skuId")
            license_name = get_license_details(sku_id, skus)
            license_names.append(license_name)
            if license_name not in license_stats:
                license_stats[license_name] = 0
            license_stats[license_name] += 1
        user_data = {
            "name": user.get("displayName", "N/A"),
            "email": user.get("userPrincipalName", "N/A"),
            "department": user.get("department", "N/A"),
            "user_type": user.get("userType", "Member"),
            "license_count": len(user_licenses),
            "licenses": license_names
        }
        if user_licenses:
            licensed_users.append(user_data)
        else:
            unlicensed_users.append(user_data)
    return licensed_users, unlicensed_users, license_stats

@app.route('/')
def home():
    if 'user' in session:
        return redirect('/dashboard')
    return render_template('login.html')

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
        headers = {"Authorization": f"Bearer {result['access_token']}"}
        me_response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
        user_info = me_response.json() if me_response.status_code == 200 else {}
        session['user'] = {
            'name': user_info.get('displayName', 'Admin'),
            'email': user_info.get('userPrincipalName', ''),
            'id': user_info.get('id', '')
        }
        session['token_data'] = {
            'access_token': result['access_token'],
            'refresh_token': result.get('refresh_token', ''),
            'expires_at': datetime.now().timestamp() + result['expires_in']
        }
        return redirect('/dashboard')
    else:
        return f"Login failed: {result.get('error_description', 'Unknown error')}", 400

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', user=session['user'])

@app.route('/api/license-data')
@login_required
def api_license_data():
    token = get_token_from_session()
    if not token:
        return jsonify({'error': 'Unauthorized'}), 401
    skus = get_subscribed_skus(token)
    users = get_all_users(token)
    licensed_users, unlicensed_users, license_stats = analyze_licenses(users, skus)
    return jsonify({
        'total_users': len(users),
        'licensed_users': len(licensed_users),
        'unlicensed_users': len(unlicensed_users),
        'license_coverage': round(len(licensed_users) / len(users) * 100, 1) if users else 0,
        'license_stats': license_stats,
        'licensed_users_list': licensed_users,
        'unlicensed_users_list': unlicensed_users[:100]
    })

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True, port=5000)