"""
app.py  —  Unified application entry point
==========================================
Mounts two sub-applications under a single Waitress server:

  /cbsl  →  cbsl.py   (CBSL report automation)
  /ifrs  →  ifrs.py   (IFRS report automation)

The root path ( / ) serves a landing dashboard with navigation buttons
for both modules.  All routes are protected by session-based auth.
"""

import os
import time
from pathlib import Path

from flask import (Flask, flash, jsonify, redirect, render_template, render_template_string,
                   request, session, url_for)
from flask.sessions import SecureCookieSessionInterface  # noqa: used in PermissionGate
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.serving import run_simple
from werkzeug.wrappers import Request as WerkzeugRequest  # noqa: used in PermissionGate

# ── Load .env into the process environment (before importing sub-apps, so any
# env-driven config they read is available) ─────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv(dotenv_path=Path(__file__).resolve().parent / ".env", override=True)
except ImportError:
    pass

# ── Auth helpers ────────────────────────────────────────────────────────────
import auth

# ── Import sub-applications ─────────────────────────────────────────────────
import cbsl as cbsl_module
import ifrs as ifrs_module

# ── Landing application ──────────────────────────────────────────────────────
landing = Flask(
    __name__,
    static_folder=str(Path(__file__).resolve().parent / "static"),
    template_folder=str(Path(__file__).resolve().parent / "templates"),
)
landing.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me-in-prod")

# ── Landing page HTML ────────────────────────────────────────────────────────
_LANDING_HTML = """\
<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Report Automation Portal</title>
  <style>
    :root {
      --bg-page: #0f172a;
      --bg-card: rgba(255,255,255,.06);
      --border: rgba(255,255,255,.12);
      --text-primary: #f1f5f9;
      --text-secondary: #94a3b8;
      --text-footer: #475569;
      --divider: rgba(255,255,255,.1);
      --btn-toggle-bg: rgba(255,255,255,.08);
      --btn-toggle-border: rgba(255,255,255,.15);
      --user-bar-bg: rgba(255,255,255,.04);
      --user-bar-border: rgba(255,255,255,.10);
    }
    [data-theme="light"] {
      --bg-page: #f0f4f8;
      --bg-card: #ffffff;
      --border: #e2e8f0;
      --text-primary: #0f172a;
      --text-secondary: #475569;
      --text-footer: #94a3b8;
      --divider: #e2e8f0;
      --btn-toggle-bg: #f1f5f9;
      --btn-toggle-border: #cbd5e1;
      --user-bar-bg: #f8fafc;
      --user-bar-border: #e2e8f0;
    }
    html { transition: background .25s; }
    html, body { height: 100%; margin: 0; }
    body {
      min-height: 100vh;
      background: var(--bg-page);
      display: flex; flex-direction: column;
      align-items: center; justify-content: center;
      font-family: 'Segoe UI', system-ui, sans-serif;
      color: var(--text-primary);
    }
    [data-theme="dark"] body {
      background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 60%, #0f172a 100%);
    }
    .portal-card {
      background: var(--bg-card);
      border: 1px solid var(--border);
      border-radius: 1.25rem;
      padding: 2.75rem 3.25rem 2.25rem;
      text-align: center;
      max-width: 540px; width: 90%;
      box-shadow: 0 20px 60px rgba(0,0,0,.35);
    }
    .portal-logo { height: 64px; width: auto; margin-bottom: 1rem; }
    .portal-title { font-size: 1.75rem; font-weight: 700; letter-spacing: .4px; margin-bottom: .3rem; }
    .portal-subtitle { color: var(--text-secondary); font-size: .92rem; margin-bottom: .5rem; }

    /* User bar */
    .user-bar {
      background: var(--user-bar-bg);
      border: 1px solid var(--user-bar-border);
      border-radius: .65rem;
      padding: .55rem 1rem;
      display: flex; align-items: center; justify-content: space-between;
      gap: .75rem; margin-bottom: 1.75rem; text-align: left;
    }
    .user-bar-name { font-size: .84rem; font-weight: 600; color: var(--text-primary); }
    .user-bar-perms { font-size: .75rem; color: var(--text-secondary); margin-top: .1rem; }
    .user-bar-actions { display: flex; gap: .5rem; flex-shrink: 0; }
    .user-bar-btn {
      background: none; border: 1px solid var(--btn-toggle-border);
      color: var(--text-secondary); border-radius: .45rem;
      padding: .25rem .6rem; font-size: .76rem; cursor: pointer;
      text-decoration: none; transition: color .15s, border-color .15s;
      display: inline-flex; align-items: center; gap: .25rem;
    }
    .user-bar-btn:hover { color: var(--text-primary); border-color: var(--text-primary); }
    .user-bar-btn-danger { border-color: rgba(239,68,68,.5); color: #f87171; }
    .user-bar-btn-danger:hover { border-color: #ef4444; color: #ef4444; background: rgba(239,68,68,.08); }

    .btn-group-portal { display: flex; flex-direction: column; align-items: center; width: 100%; }
    .btn-portal {
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      width: 100%; max-width: 320px;
      padding: 1rem 1.5rem; border-radius: .75rem;
      font-size: 1.05rem; font-weight: 600; letter-spacing: .2px;
      border: none; transition: transform .15s, box-shadow .15s, filter .15s;
      text-decoration: none; margin-bottom: .85rem;
      color: #fff; text-align: center;
    }
    .btn-portal:hover { transform: translateY(-2px); box-shadow: 0 8px 24px rgba(0,0,0,.35); filter: brightness(1.08); color: #fff; }
    .btn-portal.btn-disabled {
      opacity: .38; cursor: not-allowed; pointer-events: none;
      transform: none !important; filter: none !important;
    }
    .btn-cbsl { background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%); }
    .btn-ifrs { background: linear-gradient(135deg, #15803d 0%, #16a34a 100%); }
    .btn-admin { background: linear-gradient(135deg, #92400e 0%, #b45309 100%); }
    .btn-portal .btn-label { display: block; font-size: .78rem; font-weight: 400; opacity: .8; margin-top: .18rem; }
    .divider { border: none; border-top: 1px solid var(--divider); margin: 1.75rem 0 1.25rem; }
    .footer-row { display: flex; align-items: center; justify-content: space-between; }
    .footer-note { color: var(--text-footer); font-size: .78rem; }
    .theme-toggle-btn {
      background: var(--btn-toggle-bg); border: 1px solid var(--btn-toggle-border);
      color: var(--text-secondary); border-radius: .6rem;
      padding: .3rem .7rem; font-size: .78rem; cursor: pointer;
      transition: border-color .15s, color .15s;
    }
    .theme-toggle-btn:hover { color: var(--text-primary); border-color: var(--text-secondary); }
    .lock-icon { font-size: .85rem; opacity: .5; }
  </style>
</head>
<body>
  <div class="portal-card">
    <img src="/static/logo/logo.png" alt="Assetline Logo" class="portal-logo" onerror="this.style.display='none'">
    <div class="portal-title">Report Automation Portal</div>
    <div class="portal-subtitle">ASSETLINE FINANCE LIMITED</div>

    <!-- User bar -->
    <div class="user-bar">
      <div>
        <div class="user-bar-name">{{ username }}</div>
        <div class="user-bar-perms">{{ perms_display }}</div>
      </div>
      <div class="user-bar-actions">
        <a href="/change-password" class="user-bar-btn" title="Change password">&#128274; Password</a>
        {% if is_admin %}
        <a href="/admin/users" class="user-bar-btn" title="Manage users">&#9881; Admin</a>
        {% endif %}
        <a href="/logout" class="user-bar-btn user-bar-btn-danger" title="Sign out">&#10148; Sign Out</a>
      </div>
    </div>

    <div class="btn-group-portal">
      {% if has_cbsl %}
      <a href="/cbsl/" class="btn-portal btn-cbsl">
        CBSL
        <span class="btn-label">Central Bank of Sri Lanka &mdash; Regulatory Reports</span>
      </a>
      {% else %}
      <span class="btn-portal btn-cbsl btn-disabled">
        CBSL <span class="lock-icon">&#128274;</span>
        <span class="btn-label">No access &mdash; contact admin</span>
      </span>
      {% endif %}

      {% if has_ifrs %}
      <a href="/ifrs/" class="btn-portal btn-ifrs">
        IFRS
        <span class="btn-label">International Financial Reporting Standards</span>
      </a>
      {% else %}
      <span class="btn-portal btn-ifrs btn-disabled">
        IFRS <span class="lock-icon">&#128274;</span>
        <span class="btn-label">No access &mdash; contact admin</span>
      </span>
      {% endif %}

      {% if is_admin %}
      <a href="/admin/users" class="btn-portal btn-admin" style="margin-bottom:0;">
        Admin Dashboard
        <span class="btn-label">Manage users &amp; permissions</span>
      </a>
      {% endif %}
    </div>

    <hr class="divider">
    <div class="footer-row">
      <span class="footer-note">Report Automation System &mdash; Internal Use Only</span>
      <button class="theme-toggle-btn" onclick="toggleTheme()" id="theme-btn">&#9680; Theme</button>
    </div>
  </div>

  <script>
    (function() {
      const saved = localStorage.getItem('cbsl-theme') || 'dark';
      document.documentElement.setAttribute('data-theme', saved);
      updateThemeBtn(saved);
    })();
    function toggleTheme() {
      const current = document.documentElement.getAttribute('data-theme') || 'dark';
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('cbsl-theme', next);
      updateThemeBtn(next);
    }
    function updateThemeBtn(theme) {
      const btn = document.getElementById('theme-btn');
      if (btn) btn.innerHTML = theme === 'dark' ? '&#9728; Light' : '&#9790; Dark';
    }
  </script>

  <!-- ── Inactivity timeout (10 min) ── -->
  <script>
  (function(){
    var TIMEOUT = 10 * 60 * 1000;
    var WARN_BEFORE = 60 * 1000;
    var RECHECK = 30 * 1000;
    var timer, warnTimer, toast, wasDeferred = false;
    function resetTimer() {
      clearTimeout(timer); clearTimeout(warnTimer);
      if (toast) { toast.remove(); toast = null; }
      warnTimer = setTimeout(showWarning, TIMEOUT - WARN_BEFORE);
      timer     = setTimeout(attemptLogout, TIMEOUT);
    }
    function attemptLogout() {
      fetch('/automation-status').then(function(r){ return r.json(); }).then(function(data){
        if (data && data.running) {
          // A report automation is mid-run — don't sign the user out, keep checking.
          wasDeferred = true;
          timer = setTimeout(attemptLogout, RECHECK);
        } else if (wasDeferred) {
          // Automation just finished while the user was away — give them a
          // fresh window to notice and grab the download.
          wasDeferred = false;
          resetTimer();
        } else {
          window.location.href = '/logout?reason=timeout';
        }
      }).catch(function(){
        window.location.href = '/logout?reason=timeout';
      });
    }
    function showWarning() {
      toast = document.createElement('div');
      toast.id = 'inactivity-toast';
      toast.style.cssText = 'position:fixed;bottom:1.5rem;right:1.5rem;z-index:9999;background:#1e293b;border:1px solid rgba(249,115,22,.5);color:#fdba74;padding:.85rem 1.25rem;border-radius:.75rem;font-size:.88rem;font-family:Segoe UI,system-ui,sans-serif;box-shadow:0 8px 24px rgba(0,0,0,.5);max-width:320px;line-height:1.5;';
      toast.innerHTML = '&#9888; You will be signed out in <strong id="toast-countdown">60</strong>s due to inactivity. <a href="#" onclick="resetTimer();return false;" style="color:#60a5fa;text-decoration:underline;margin-left:.5rem;">Stay signed in</a>';
      document.body.appendChild(toast);
      var secs = 60;
      var interval = setInterval(function(){
        secs--;
        var el = document.getElementById('toast-countdown');
        if (el) el.textContent = secs;
        if (secs <= 0) clearInterval(interval);
      }, 1000);
    }
    ['mousemove','mousedown','keydown','touchstart','scroll','click'].forEach(function(e){
      document.addEventListener(e, resetTimer, { passive: true });
    });
    resetTimer();
  })();
  </script>
</body>
</html>
"""


# ── Inactivity timeout ───────────────────────────────────────────────────────
INACTIVITY_TIMEOUT = 40 * 60  # 40 minutes in seconds


def _is_automation_running() -> bool:
    """True if the CBSL or IFRS module is mid-run in a background thread."""
    try:
        with cbsl_module._status_lock:
            if cbsl_module._status.get("running"):
                return True
    except Exception:
        pass
    try:
        with ifrs_module._lock:
            if ifrs_module._state.get("running"):
                return True
    except Exception:
        pass
    return False


@landing.route("/automation-status")
def automation_status():
    return jsonify(running=_is_automation_running())


@landing.before_request
def check_inactivity():
    # Skip for login/logout/static/status-poll
    if request.endpoint in ("login", "logout", "static", "automation_status"):
        return
    if not session.get("username"):
        return
    last = session.get("last_activity")
    now  = time.time()
    if last and (now - last) > INACTIVITY_TIMEOUT:
        if _is_automation_running():
            # Don't kick the user out while a report is mid-run — just push
            # the deadline out and let them come back to grab the download.
            session["last_activity"] = now
            return
        session.clear()
        flash("You were signed out due to inactivity.", "warning")
        return redirect(url_for("login"))
    session["last_activity"] = now


# ── Auth routes ──────────────────────────────────────────────────────────────

@landing.route("/login", methods=["GET", "POST"])
def login():
    if auth.current_user():
        return redirect(url_for("home"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        next_url = request.form.get("next", "/")

        user = auth.authenticate(username, password)
        if user:
            session["username"]       = username
            session["permissions"]    = user.get("permissions", "")
            session["is_admin"]       = auth.is_admin(username)
            session["last_activity"]  = time.time()
            return redirect(next_url if next_url.startswith("/") else "/")
        else:
            flash("Invalid username or password.", "error")
            return render_template("login.html", next=next_url, username=username)

    next_url = request.args.get("next", "/")
    return render_template("login.html", next=next_url, username="")


@landing.route("/logout")
def logout():
    session.clear()
    if request.args.get("reason") == "timeout":
        flash("You were signed out due to inactivity.", "warning")
    else:
        flash("You have been signed out.", "warning")
    return redirect(url_for("login"))


@landing.route("/")
def home():
    username = auth.current_user()
    if not username:
        return redirect(url_for("login"))

    perms    = auth.current_permissions()
    has_cbsl = auth.has_permission("CBSL")
    has_ifrs = auth.has_permission("IFRS")
    is_admin = auth.is_admin(username)

    parts = []
    if has_cbsl: parts.append("CBSL")
    if has_ifrs: parts.append("IFRS")
    if is_admin: parts.append("Admin")
    perms_display = "Access: " + ", ".join(parts) if parts else "No permissions assigned"

    return render_template_string(
        _LANDING_HTML,
        username=username,
        perms_display=perms_display,
        has_cbsl=has_cbsl,
        has_ifrs=has_ifrs,
        is_admin=is_admin,
    )


@landing.route("/change-password", methods=["GET", "POST"])
def change_password():
    username = auth.current_user()
    if not username:
        flash("Please log in to continue.", "warning")
        return redirect(url_for("login", next="/change-password"))

    if request.method == "POST":
        current_pw  = request.form.get("current_password", "")
        new_pw      = request.form.get("new_password", "")
        confirm_pw  = request.form.get("confirm_password", "")

        if not auth.authenticate(username, current_pw):
            flash("Current password is incorrect.", "error")
        elif new_pw != confirm_pw:
            flash("New passwords do not match.", "error")
        elif len(new_pw) < 6:
            flash("Password must be at least 6 characters.", "error")
        else:
            auth.update_password(username, new_pw)
            flash("Password updated successfully.", "success")
            return redirect(url_for("home"))

    return render_template("change_password.html", current_user=username)


# ── Admin routes ─────────────────────────────────────────────────────────────

def _require_admin():
    """Returns a redirect response if the current user is not admin, else None."""
    username = auth.current_user()
    if not username:
        flash("Please log in to continue.", "warning")
        return redirect(url_for("login", next=request.path))
    if not auth.has_permission("ADMIN"):
        flash("Admin access required.", "error")
        return redirect(url_for("home"))
    return None


@landing.route("/admin/users")
def admin_users():
    guard = _require_admin()
    if guard:
        return guard
    users = auth.list_users_display()
    return render_template("admin_users.html", users=users)


@landing.route("/admin/create-user", methods=["POST"])
def admin_create_user():
    guard = _require_admin()
    if guard:
        return guard

    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    perm_cbsl = request.form.get("perm_cbsl")
    perm_ifrs = request.form.get("perm_ifrs")

    perms_list = []
    if perm_cbsl: perms_list.append("CBSL")
    if perm_ifrs: perms_list.append("IFRS")
    permissions = ",".join(perms_list) if perms_list else "NONE"

    if not username or not password:
        flash("Username and password are required.", "error")
    elif auth.create_user(username, password, permissions):
        flash(f"User '{username}' created successfully.", "success")
    else:
        flash(f"Username '{username}' already exists.", "error")

    return redirect(url_for("admin_users"))


@landing.route("/admin/update-permissions", methods=["POST"])
def admin_update_permissions():
    guard = _require_admin()
    if guard:
        return guard

    username  = request.form.get("username", "").strip()
    perm_cbsl = request.form.get("perm_cbsl")
    perm_ifrs = request.form.get("perm_ifrs")

    perms_list = []
    if perm_cbsl: perms_list.append("CBSL")
    if perm_ifrs: perms_list.append("IFRS")
    permissions = ",".join(perms_list) if perms_list else "NONE"

    if auth.update_permissions(username, permissions):
        flash(f"Permissions updated for '{username}'.", "success")
    else:
        flash(f"User '{username}' not found.", "error")

    return redirect(url_for("admin_users"))


@landing.route("/admin/reset-password", methods=["POST"])
def admin_reset_password():
    guard = _require_admin()
    if guard:
        return guard

    username     = request.form.get("username", "").strip()
    new_password = request.form.get("new_password", "")

    if not new_password:
        flash("New password cannot be empty.", "error")
    elif auth.update_password(username, new_password):
        flash(f"Password reset for '{username}'.", "success")
    else:
        flash(f"User '{username}' not found.", "error")

    return redirect(url_for("admin_users"))


@landing.route("/admin/delete-user", methods=["POST"])
def admin_delete_user():
    guard = _require_admin()
    if guard:
        return guard

    username = request.form.get("username", "").strip()
    target   = auth.get_user(username)
    if target and "ADMIN" in str(target.get("permissions", "")).upper():
        flash("Admin accounts cannot be deleted.", "error")
    elif auth.delete_user(username):
        flash(f"User '{username}' deleted.", "success")
    else:
        flash(f"User '{username}' not found.", "error")

    return redirect(url_for("admin_users"))


# ── CBSL / IFRS permission gate middleware ────────────────────────────────────
class PermissionGate:
    """
    WSGI middleware that checks Flask session permission before forwarding
    requests to a sub-application.  Redirects to / on failure.
    """
    def __init__(self, app_wsgi, required_perm: str, landing_app):
        self._app          = app_wsgi
        self._required     = required_perm.upper()
        self._landing_app  = landing_app

    def __call__(self, environ, start_response):
        req = WerkzeugRequest(environ)
        raw_cookie = req.cookies.get("session", "")

        username    = None
        permissions = ""

        if raw_cookie:
            try:
                iface  = SecureCookieSessionInterface()
                signer = iface.get_signing_serializer(landing)
                if signer:
                    data        = signer.loads(raw_cookie)
                    username    = data.get("username")
                    permissions = data.get("permissions", "")
            except Exception:
                pass

        if not username:
            # Not logged in — redirect to login
            redirect_to = "/login?next=" + environ.get("PATH_INFO", "/")
            start_response("302 Found", [("Location", redirect_to), ("Content-Type", "text/html")])
            return [b""]

        if self._required not in permissions.upper():
            # No permission — redirect to home
            start_response("302 Found", [("Location", "/"), ("Content-Type", "text/html")])
            return [b""]

        return self._app(environ, start_response)


# ── Compose the unified WSGI application ──────────────────────────────────────
application = DispatcherMiddleware(landing, {
    "/cbsl": PermissionGate(cbsl_module.app, "CBSL", landing),
    "/ifrs": PermissionGate(ifrs_module.app, "IFRS", landing),
})


# ── Startup bootstrap ─────────────────────────────────────────────────────────
def _startup():
    auth.bootstrap_admin()


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    from cbsl import kill_excel_instances, ensure_directories_exist

    _startup()
    kill_excel_instances()
    ensure_directories_exist()

    print("Starting server on http://0.0.0.0:5000")
    run_simple("127.0.0.1", 5000, application, use_reloader=False, use_debugger=False)
