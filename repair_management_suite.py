"""
Repair Management Suite

Machine repair tracking, service orders and customer sign-off for field and
shop technicians.

Author:  tylersword12
Source:  https://github.com/tylersword12/RMS
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json
import hashlib
import os
import sys
import re
import threading
import socket
import ssl
import subprocess
import uuid
import urllib.request
import urllib.error
from datetime import datetime
import csv
import shutil
from pathlib import Path

# ============================================================
#  APPLICATION VERSION  /  AUTO-UPDATE
#
#  To ship an update:
#    1. Bump APP_VERSION below.
#    2. git commit -am "..."  &&  git push
#    3. git tag v1.3.0  &&  git push origin v1.3.0
#
#  GitHub Actions builds the Windows EXE and publishes it as a
#  release. Every installed copy sees it on next launch and
#  offers to install it on close.
#
#  The tag and APP_VERSION must match (tag v1.3.0 -> "1.3.0").
# ============================================================
APP_VERSION = "1.3.4"
GITHUB_OWNER = "tylersword12"
GITHUB_REPO = "RMS"
GITHUB_API_LATEST = (f"https://api.github.com/repos/{GITHUB_OWNER}/"
                     f"{GITHUB_REPO}/releases/latest")
GITHUB_RELEASES_PAGE = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases"

_HTTPS_CONTEXT = None


def https_context():
    """Secure-connection settings for talking to GitHub.

    Trusts the certificates Windows has AND the copy bundled with the program
    (certifi). Windows only downloads some root certificates when a browser
    first needs them, so relying on Windows alone can fail on a machine that
    hasn't visited the site yet - even though the browser works fine.
    """
    global _HTTPS_CONTEXT
    if _HTTPS_CONTEXT is None:
        ctx = ssl.create_default_context()
        try:
            import certifi
            ctx.load_verify_locations(cafile=certifi.where())
        except Exception as e:
            print(f"Bundled certificates not available: {e}")
        _HTTPS_CONTEXT = ctx
    return _HTTPS_CONTEXT


def describe_network_error(e):
    """Plain-English reason a request to GitHub failed."""
    reason = getattr(e, 'reason', e)
    text = str(reason)
    if isinstance(reason, ssl.SSLCertVerificationError) or 'CERTIFICATE_VERIFY_FAILED' in text:
        return ("The secure connection to GitHub couldn't be verified (certificate "
                "problem). Antivirus or network filtering that inspects secure "
                f"traffic often causes this.\n\nDetail: {text}")
    if isinstance(reason, ssl.SSLError):
        return f"The secure connection to GitHub failed.\n\nDetail: {text}"
    if isinstance(reason, (socket.timeout, TimeoutError)) or 'timed out' in text.lower():
        return ("GitHub didn't answer in time. The internet connection may be slow, "
                "or a firewall may be blocking it.")
    if isinstance(reason, socket.gaierror):
        return ("Couldn't find github.com. This computer may be offline, or DNS is "
                f"blocked.\n\nDetail: {text}")
    return f"Couldn't connect to GitHub.\n\nDetail: {type(reason).__name__}: {text}"


def parse_version(text):
    """'1.2.10' -> (1, 2, 10). None if it isn't a version string."""
    if not text:
        return None
    m = re.match(r'^\s*v?(\d+(?:\.\d+)*)\s*$', str(text))
    if not m:
        return None
    try:
        return tuple(int(p) for p in m.group(1).split('.'))
    except ValueError:
        return None


def version_key(ver):
    """Pad so (1, 2) and (1, 2, 0) compare equal, and 1.10 > 1.9."""
    t = parse_version(ver) if isinstance(ver, str) else tuple(ver or ())
    if not t:
        return (0,) * 4
    return tuple(t) + (0,) * (4 - len(t)) if len(t) < 4 else tuple(t[:4])


def format_version(ver):
    """(1, 3, 0) -> '1.3.0'."""
    parts = list(ver or ())
    return ".".join(str(p) for p in parts) if parts else "0.0"


# Excel support
OPENPYXL_ERROR = ""
try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
    EXCEL_AVAILABLE = True
except Exception as _openpyxl_err:
    # Was `except ImportError` with OPENPYXL_ERROR never defined, so the
    # price-list warning below raised NameError instead of explaining itself.
    OPENPYXL_ERROR = str(_openpyxl_err)
    OPENPYXL_AVAILABLE = False
    EXCEL_AVAILABLE = False

# Webcam barcode scanning support
WEBCAM_AVAILABLE = True
WEBCAM_IMPORT_ERRORS = []
PYZBAR_AVAILABLE = False
try:
    import cv2
except Exception as e:
    WEBCAM_AVAILABLE = False
    WEBCAM_IMPORT_ERRORS.append(f"opencv-python: {e}")

try:
    from pyzbar import pyzbar
    PYZBAR_AVAILABLE = True
except Exception as e:
    # pyzbar failed - we can still use OpenCV's built-in barcode detector as fallback
    WEBCAM_IMPORT_ERRORS.append(f"pyzbar (will use OpenCV fallback): {e}")

# Barcode printing support
BARCODE_AVAILABLE = True
BARCODE_IMPORT_ERRORS = []

# Try to setup libusb backend before importing brother_ql
# This handles the "No backend available" error from pyusb on Windows
_LIBUSB_BACKEND = None
try:
    import libusb_package
    import usb.backend.libusb1
    import usb.core
    
    # Get the bundled libusb-1.0 DLL path
    libusb1_lib = libusb_package.find_library("libusb-1.0")
    if libusb1_lib:
        _LIBUSB_BACKEND = usb.backend.libusb1.get_backend(find_library=lambda x: libusb1_lib)
        
        if _LIBUSB_BACKEND is not None:
            # Patch usb.core.find to use this backend by default
            _original_find = usb.core.find
            def _patched_find(*args, **kwargs):
                if 'backend' not in kwargs:
                    kwargs['backend'] = _LIBUSB_BACKEND
                return _original_find(*args, **kwargs)
            usb.core.find = _patched_find
except Exception as e:
    pass  # Will be reported below if brother_ql also fails

try:
    from brother_ql.conversion import convert
    from brother_ql.backends.helpers import send
    from brother_ql.raster import BrotherQLRaster
except Exception as e:
    BARCODE_AVAILABLE = False
    BARCODE_IMPORT_ERRORS.append(f"brother_ql: {e}")

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception as e:
    BARCODE_AVAILABLE = False
    BARCODE_IMPORT_ERRORS.append(f"Pillow: {e}")

try:
    import barcode
    from barcode.writer import ImageWriter
except Exception as e:
    BARCODE_AVAILABLE = False
    BARCODE_IMPORT_ERRORS.append(f"python-barcode: {e}")

from io import BytesIO
# PDF support for service orders
REPORTLAB_AVAILABLE = False
REPORTLAB_ERROR = ""
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.pdfgen import canvas as pdf_canvas
    REPORTLAB_AVAILABLE = True
except Exception as e:
    REPORTLAB_ERROR = str(e)

# Excel support for old .xls files (price lists)
XLRD_AVAILABLE = False
XLRD_ERROR = ""
try:
    import xlrd
    XLRD_AVAILABLE = True
except Exception as e:
    XLRD_ERROR = str(e)


if BARCODE_IMPORT_ERRORS:
    print(f"Barcode printing unavailable: {'; '.join(BARCODE_IMPORT_ERRORS)}")

def get_offline_service_orders_dir():
    """Local cache folder for service orders saved while the share was down."""
    base = os.environ.get('LOCALAPPDATA') or os.path.expanduser("~")
    path = os.path.join(base, "RepairManagementSuite", "service_orders_offline")
    try:
        os.makedirs(path, exist_ok=True)
        return path
    except Exception:
        fallback = os.path.join(os.path.expanduser("~"),
                                ".repair_management_suite", "service_orders_offline")
        os.makedirs(fallback, exist_ok=True)
        return fallback


def read_service_order_stamp(path):
    """Read 'last_saved' out of a service order file. Missing/bad -> ''."""
    try:
        with open(path, 'r') as f:
            return json.load(f).get('last_saved', '') or ''
    except Exception:
        return ''


def push_offline_service_orders(service_orders_dir):
    """Upload locally cached service orders to the share.

    Only uploads when the local copy is newer than the server copy, so a
    stale offline file can never clobber a later edit made elsewhere.
    Returns (uploaded_count, failure_messages).
    """
    if not service_orders_dir:
        return 0, []
    try:
        local_dir = get_offline_service_orders_dir()
    except Exception:
        return 0, []
    if not os.path.isdir(local_dir):
        return 0, []
    try:
        os.makedirs(service_orders_dir, exist_ok=True)
    except Exception:
        return 0, []  # share still unreachable

    uploaded, failed = 0, []
    for name in os.listdir(local_dir):
        if not name.endswith('_service_order.json'):
            continue
        local_path = os.path.join(local_dir, name)
        server_path = os.path.join(service_orders_dir, name)
        try:
            local_stamp = read_service_order_stamp(local_path)
            if os.path.exists(server_path):
                server_stamp = read_service_order_stamp(server_path)
                if server_stamp and local_stamp and server_stamp >= local_stamp:
                    continue
            shutil.copy2(local_path, server_path)
            uploaded += 1
        except Exception as e:
            failed.append(f"{name}: {e}")
    return uploaded, failed


def new_record_id():
    """Short, globally unique id for a machine record."""
    return uuid.uuid4().hex[:12]


def record_key(machine):
    """Stable identity for a record. Falls back to the estimate number
    for records created before record_id existed."""
    rid = machine.get('record_id')
    if rid:
        return str(rid)
    return 'est:' + str(machine.get('estimate_number', ''))


def record_stamp(machine):
    """When this record last changed. Used to pick a winner on merge."""
    return str(machine.get('updated_at') or machine.get('created_at') or '')


def legacy_record_id(machine):
    """Stable id for a job saved before record_id existed.

    Derived from the estimate number, so the shared copy, this computer's
    copy, and every other computer all give the same old job the same id.
    A random id here made one job look like two and triggered renumbering.
    """
    num = str(machine.get('estimate_number', '') or '').strip()
    if not num:
        return None
    return 'L' + hashlib.sha1(num.encode('utf-8')).hexdigest()[:11]


def adopt_record_ids(target, reference):
    """Give id-less records in `target` the id their twin in `reference`
    already has (matched by estimate number).

    Covers a computer still holding pre-id data after another computer has
    already assigned ids and saved them to the share.
    """
    ref = {}
    for m in reference or []:
        num = str(m.get('estimate_number', '') or '').strip()
        rid = m.get('record_id')
        if num and rid:
            ref.setdefault(num, []).append(str(rid))
    used = {str(m.get('record_id')) for m in target or [] if m.get('record_id')}
    adopted = 0
    for m in target or []:
        if m.get('record_id'):
            continue
        num = str(m.get('estimate_number', '') or '').strip()
        for rid in ref.get(num, []):
            if rid not in used:
                m['record_id'] = rid
                used.add(rid)
                adopted += 1
                break
    return adopted


def ensure_record_ids(machines):
    """Give any record without one a record_id and a created_at.

    Pass active and shipped together (a + b) so ids are unique across both.
    """
    changed = 0
    now = datetime.now().isoformat()
    used = {str(m.get('record_id')) for m in machines if m.get('record_id')}
    for m in machines:
        if not m.get('record_id'):
            rid = legacy_record_id(m)
            if not rid or rid in used:
                rid = new_record_id()
            m['record_id'] = rid
            used.add(rid)
            changed += 1
        if not m.get('created_at'):
            # Best guess for an existing record, so merges have something to sort on
            m['created_at'] = m.get('arrival_date') or now
        if not m.get('updated_at'):
            m['updated_at'] = m['created_at']
    return changed


def index_records(active, shipped):
    """{key: (record, 'active'|'shipped')} across both lists."""
    idx = {}
    for m in active or []:
        idx[record_key(m)] = (m, 'active')
    for m in shipped or []:
        idx[record_key(m)] = (m, 'shipped')
    return idx


def merge_machine_data(server_active, server_shipped,
                       local_active, local_shipped, synced_ids):
    """Combine the shared copy with this computer's copy without losing work.

    Rules:
      - In both copies  -> the one changed most recently wins.
      - Server only     -> another computer added it; take it.
      - Local only and previously synced -> it was deleted elsewhere; drop it.
      - Local only and never synced      -> work done offline here; keep it.
      - Server only and previously synced -> deleted on this computer; drop it.

    A key in synced_ids means both sides knew about the record, so either
    side no longer having it is a real deletion rather than a new addition.
    synced_ids starts empty on a fresh install, so nothing is ever inferred
    as deleted until this computer has actually seen the server copy.

    Returns (active, shipped, stats).
    """
    server_idx = index_records(server_active, server_shipped)
    local_idx = index_records(local_active, local_shipped)
    synced = set(synced_ids or ())

    stats = {'server_wins': 0, 'local_wins': 0, 'added_from_server': 0,
             'kept_offline': 0, 'deleted_elsewhere': 0, 'deleted_here': 0}
    chosen = {}

    for key in set(server_idx) | set(local_idx):
        s = server_idx.get(key)
        l = local_idx.get(key)
        if s and l:
            srec, slist = s
            lrec, llist = l
            if record_stamp(lrec) > record_stamp(srec):
                chosen[key] = (lrec, llist)
                stats['local_wins'] += 1
            else:
                chosen[key] = (srec, slist)
                stats['server_wins'] += 1
        elif s:
            if key in synced:
                # This computer had it, saw it on the server, and no longer
                # holds it - that is a deletion made here. Symmetric with the
                # local-only case below.
                stats['deleted_here'] += 1
                continue
            chosen[key] = s
            stats['added_from_server'] += 1
        else:
            lrec, llist = l
            if key in synced:
                # It reached the server once and is gone now - a real deletion
                stats['deleted_elsewhere'] += 1
                continue
            chosen[key] = (lrec, llist)
            stats['kept_offline'] += 1

    # Keep the server's ordering so the log doesn't reshuffle, then append
    # anything that only exists on this computer.
    active, shipped = [], []
    seen = set()
    for src in (server_active or [], server_shipped or []):
        for m in src:
            k = record_key(m)
            if k in chosen and k not in seen:
                rec, lst = chosen[k]
                (shipped if lst == 'shipped' else active).append(rec)
                seen.add(k)
    for src in (local_active or [], local_shipped or []):
        for m in src:
            k = record_key(m)
            if k in chosen and k not in seen:
                rec, lst = chosen[k]
                (shipped if lst == 'shipped' else active).append(rec)
                seen.add(k)

    return active, shipped, stats


# Fields that identify a physical job. If these all match, two records are
# almost certainly the same job entered (or synced) twice.
JOB_IDENTITY_FIELDS = ('customer_name', 'serial_number', 'make', 'model', 'arrival_date')


def _norm_field(value):
    return ' '.join(str(value or '').split()).lower()


def same_job(a, b):
    return all(_norm_field(a.get(f)) == _norm_field(b.get(f))
               for f in JOB_IDENTITY_FIELDS)


def find_estimate_conflicts(active, shipped):
    """Find estimate numbers that need a decision. Changes nothing.

    Two kinds:
      'shared'     - two or more jobs currently hold the same number.
      'renumbered' - an earlier version renumbered a job automatically
                     (it carries 'renumbered_from') and the job that kept
                     the original number is still here.

    Returns a list of {'number', 'kind', 'members': [(record, 'Active'|
    'Shipped')], 'identical'} sorted by number.
    """
    lists = (('Active', active or []), ('Shipped', shipped or []))
    by_number = {}
    for where, lst in lists:
        for m in lst:
            num = str(m.get('estimate_number', '') or '').strip()
            if num:
                by_number.setdefault(num, []).append((m, where))

    groups = {}
    for num, members in by_number.items():
        if len(members) > 1:
            groups[num] = {'number': num, 'kind': 'shared', 'members': list(members)}

    for where, lst in lists:
        for m in lst:
            old = str(m.get('renumbered_from', '') or '').strip()
            num = str(m.get('estimate_number', '') or '').strip()
            if not old or old == num or old not in by_number:
                continue
            g = groups.get(old)
            if g is None:
                g = groups[old] = {'number': old, 'kind': 'renumbered',
                                   'members': list(by_number[old])}
            if not any(x is m for x, _ in g['members']):
                g['members'].append((m, where))

    out = []
    for num in sorted(groups):
        g = groups[num]
        recs = [m for m, _ in g['members']]
        g['identical'] = all(same_job(recs[0], r) for r in recs[1:])
        out.append(g)
    return out


def conflict_signature(conflict):
    """Identifies a conflict so 'decide later' isn't re-asked every save."""
    return (conflict['number'],
            frozenset(record_key(m) for m, _ in conflict['members']))


def _current_records(active, shipped, conflict):
    """The live records for a conflict's members, looked up by id, since an
    autosave can swap in fresh record objects after the conflict was found."""
    keys = [record_key(m) for m, _ in conflict['members']]
    live = {record_key(m): m for m in list(active) + list(shipped)}
    seen, out = set(), []
    for k in keys:
        if k in live and k not in seen:
            out.append(live[k])
            seen.add(k)
    return out


def merge_conflict(active, shipped, conflict):
    """The user said these are the same job: keep one, drop the copies.

    The most recently edited record is kept and takes the group's estimate
    number. Lists are modified in place.
    Returns (kept, removed_list, other_numbers) or None if nothing to do.
    """
    number = conflict['number']
    recs = _current_records(active, shipped, conflict)
    if len(recs) < 2:
        return None
    other_numbers = {str(m.get('estimate_number', '') or '') for m in recs} - {number, ''}
    keep = max(recs, key=lambda m: (record_stamp(m),
                                    str(m.get('estimate_number', '')) == number))
    removed = [m for m in recs if m is not keep]
    gone = {id(m) for m in removed}
    active[:] = [m for m in active if id(m) not in gone]
    shipped[:] = [m for m in shipped if id(m) not in gone]
    keep['estimate_number'] = number
    keep.pop('renumbered_from', None)
    return keep, removed, other_numbers


def renumber_conflict(active, shipped, conflict):
    """The user said these are different jobs.

    For jobs that share a number, the one created first keeps it and the
    others get the next free number. Jobs an earlier version already
    renumbered keep the numbers they have. Lists are modified in place.
    Returns [(record, old_number, new_number)].
    """
    number = conflict['number']
    recs = _current_records(active, shipped, conflict)
    taken = {str(m.get('estimate_number', '') or '') for m in list(active) + list(shipped)}
    holders = [m for m in recs if str(m.get('estimate_number', '')) == number]
    holders.sort(key=lambda m: (str(m.get('created_at') or ''), record_key(m)))
    prefix = number.rsplit('-', 1)[0] if '-' in number else number
    changes = []
    for dup in holders[1:]:
        n = 1
        while f"{prefix}-{n:03d}" in taken:
            n += 1
        new_num = f"{prefix}-{n:03d}"
        dup['estimate_number'] = new_num
        taken.add(new_num)
        changes.append((dup, number, new_num))
    for m in recs:
        m.pop('renumbered_from', None)
    return changes


MAX_LABOR_RATES = 10
# Used until rates are set up, so new service orders price the same as before
DEFAULT_LABOR_RATES = [{'name': 'Standard', 'rate': '150.00'}]


def parse_money(text):
    """'$1,250.5' -> 1250.5. None if it isn't a number."""
    try:
        return float(str(text).replace('$', '').replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def clean_labor_rates(rates):
    """Keep only rates that have a name and a valid, non-negative price,
    in order, at most MAX_LABOR_RATES. Prices are stored as '150.00'."""
    out, seen = [], set()
    for r in rates or []:
        if not isinstance(r, dict):
            continue
        name = ' '.join(str(r.get('name', '') or '').split())
        rate = parse_money(r.get('rate', ''))
        if not name or rate is None or rate < 0 or name.lower() in seen:
            continue
        seen.add(name.lower())
        out.append({'name': name, 'rate': f"{rate:.2f}"})
    return out[:MAX_LABOR_RATES]


class MachineRepairApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Repair Management Suite")
        self.root.geometry("1200x700")
        self.root.configure(bg='#2c3e50')
        
        # Data storage
        self.active_machines = []
        self.shipped_machines = []
        self.save_location = os.path.join(os.path.expanduser("~"), "Desktop", "repair_data.json")
        self.attachments_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Machine_Attachments")
        
        # Printer settings
        self.printer_identifier = None  # Will be set by user
        
        # Company branding
        self.company_name = ""
        
        # Service order settings
        self.price_list_folder = ""
        self.column_mappings = {}
        self.price_lists = {}
        self.service_orders_dir = ""
        self._service_order_windows = []
        self.labor_rates = [dict(r) for r in DEFAULT_LABOR_RATES]
        self.labor_rates_updated = ""   # '' = never set up on this computer
        
        # Initialize GUI components to None
        self.status_label = None
        self.estimate_label = None
        self.notebook = None
        self.active_tree = None
        self.shipped_tree = None
        self.customer_details = None
        self.active_customer_filter = None
        self.active_status_filter = None
        self.active_location_filter = None
        self.shipped_location_filter = None
        self.shipped_customer_filter = None
        self.customer_select = None
        
        # Offline mode tracking
        self._was_offline = False
        self.connection_label = None
        
        # Barcode scanner listener
        self._scan_buffer = ""
        self._scan_last_time = 0
        self._scan_popup_open = False

        # Sync state
        self._synced_ids = set()
        self._record_snapshot = {}
        self._pending_conflicts = []
        self._conflicts_deferred = set()
        self._conflict_prompt_open = False
        self._conflict_prompt_scheduled = False
        self._last_merge_stats = {}
        
        # Auto-update state
        self._pending_update = None
        self.update_label = None
        self._update_result_message = None
        self._update_result_color = '#27ae60'
        
        # Create attachments directory if it doesn't exist
        os.makedirs(self.attachments_dir, exist_ok=True)
        
        # Load saved data and settings FIRST
        self.load_settings()

        # Tell the updater this build started cleanly, and report the outcome
        self._confirm_update_launch()
        
        # Update title with company name
        if self.company_name:
            self.root.title(f"{self.company_name} - Repair Management Suite")
        
        # Check for first-time setup (trigger only if company name not set)
        if not self.company_name:
            self.first_time_setup()
        
        # Setup service orders directory (may fail if server offline - that's ok)
        self.service_orders_dir = os.path.join(os.path.dirname(self.save_location), "Service_Orders")
        try:
            os.makedirs(self.service_orders_dir, exist_ok=True)
        except Exception as e:
            print(f"Could not create service orders dir (offline?): {e}")

        try:
            self.load_labor_rates()
        except Exception as e:
            print(f"Could not load labor rates: {e}")

        self.load_data()
        
        # Create GUI components
        self.create_widgets()
        
        # Then update displays AFTER GUI is created
        self.update_estimate_display()
        self.refresh_tables()
        
        # Setup global barcode scanner listener
        self.setup_barcode_scanner_listener()
        
        # Start offline sync watcher (checks server every 15s)
        self.start_offline_sync_watcher()
        
        # Show initial connection status
        self._update_connection_indicator(self.is_server_reachable())
        
        # Say what the merge did (recovered offline jobs, renumbering)
        try:
            self.report_sync_results()
        except Exception as e:
            print(f"Could not report sync results: {e}")
        
        # Report any update that just happened, then look for a new one
        self._show_update_result()
        self.check_for_updates_async()
        
        # Upload any service orders that were saved while offline last time
        try:
            pushed, _ = push_offline_service_orders(self.service_orders_dir)
            if pushed and self.status_label:
                self.status_label.config(
                    text=f"✅ Uploaded {pushed} service order(s) saved offline",
                    fg='#27ae60')
        except Exception as e:
            print(f"Startup service order sync failed: {e}")
        
        # Auto-save timer
        self.auto_save_interval = 30000  # 30 seconds
        self.schedule_auto_save()
        
        # Handle window closing
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    

    def first_time_setup(self):
        """Show first-time setup wizard"""
        setup_win = tk.Toplevel(self.root)
        setup_win.title("Welcome - First Time Setup")
        setup_win.geometry("550x400")
        setup_win.configure(bg='#2c3e50')
        setup_win.transient(self.root)
        setup_win.grab_set()
        setup_win.protocol("WM_DELETE_WINDOW", lambda: None)  # Prevent closing
        
        # Center
        setup_win.update_idletasks()
        x = (setup_win.winfo_screenwidth() // 2) - 275
        y = (setup_win.winfo_screenheight() // 2) - 200
        setup_win.geometry(f"+{x}+{y}")
        
        tk.Label(setup_win, text="Welcome to Repair Management Suite",
                font=('Arial', 18, 'bold'), fg='white', bg='#2c3e50').pack(pady=(30, 5))
        tk.Label(setup_win, text="Let\'s get you set up.",
                font=('Arial', 12), fg='#bdc3c7', bg='#2c3e50').pack(pady=(0, 20))
        
        # Company name
        name_frame = tk.Frame(setup_win, bg='#34495e', padx=20, pady=15)
        name_frame.pack(fill='x', padx=30, pady=10)
        
        tk.Label(name_frame, text="Company Name:", font=('Arial', 12, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w')
        tk.Label(name_frame, text="This will appear on service orders and reports.",
                font=('Arial', 9), fg='#95a5a6', bg='#34495e').pack(anchor='w')
        company_var = tk.StringVar(value=self.company_name)
        tk.Entry(name_frame, textvariable=company_var, width=40, font=('Arial', 12)).pack(pady=(5, 0))
        
        # Save location
        loc_frame = tk.Frame(setup_win, bg='#34495e', padx=20, pady=15)
        loc_frame.pack(fill='x', padx=30, pady=10)
        
        tk.Label(loc_frame, text="Data Save Location:", font=('Arial', 12, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w')
        tk.Label(loc_frame, text="Choose where to store repair data, service orders, and attachments.",
                font=('Arial', 9), fg='#95a5a6', bg='#34495e').pack(anchor='w')
        
        loc_var = tk.StringVar(value=os.path.dirname(self.save_location))
        loc_inner = tk.Frame(loc_frame, bg='#34495e')
        loc_inner.pack(fill='x', pady=(5, 0))
        tk.Entry(loc_inner, textvariable=loc_var, width=35, font=('Arial', 10)).pack(side='left')
        
        def browse_location():
            folder = filedialog.askdirectory(title="Select Data Folder")
            if folder:
                loc_var.set(folder)
        
        tk.Button(loc_inner, text="Browse...", command=browse_location,
                 bg='#3498db', fg='white', font=('Arial', 10)).pack(side='left', padx=5)
        
        def finish_setup():
            name = company_var.get().strip()
            location = loc_var.get().strip()
            
            if not name:
                messagebox.showwarning("Required", "Please enter a company name.", parent=setup_win)
                return
            if not location:
                messagebox.showwarning("Required", "Please choose a save location.", parent=setup_win)
                return
            
            self.company_name = name
            self.save_location = os.path.join(location, "repair_data.json")
            self.attachments_dir = os.path.join(location, "Machine_Attachments")
            os.makedirs(self.attachments_dir, exist_ok=True)
            
            self.root.title(f"{self.company_name} - Repair Management Suite")
            self.save_settings()
            setup_win.destroy()
        
        tk.Button(setup_win, text="Complete Setup", command=finish_setup,
                 bg='#27ae60', fg='white', font=('Arial', 14, 'bold'),
                 width=20).pack(pady=25)
        
        setup_win.wait_window()

    def load_settings(self):
        """Load app settings including save location"""
        settings_file = "repair_settings.json"
        if os.path.exists(settings_file):
            try:
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
                    self.save_location = settings.get('save_location', self.save_location)
                    self.attachments_dir = settings.get('attachments_dir', self.attachments_dir)
                    self.printer_identifier = settings.get('printer_identifier', None)
                    self.company_name = settings.get('company_name', '')
                    self.price_list_folder = settings.get('price_list_folder', '')
                    self.column_mappings = settings.get('column_mappings', {})
                    if 'labor_rates' in settings:
                        self.labor_rates = clean_labor_rates(settings.get('labor_rates'))
                        self.labor_rates_updated = str(settings.get('labor_rates_updated', '') or '')
            except Exception as e:
                print(f"Error loading settings: {e}")

    def save_settings(self):
        """Save app settings"""
        settings = {
            'save_location': self.save_location,
            'attachments_dir': self.attachments_dir,
            'printer_identifier': self.printer_identifier,
            'company_name': self.company_name,
            'price_list_folder': self.price_list_folder,
            'column_mappings': self.column_mappings,
            # Local copy of the shared labor rates, used when the share is offline
            'labor_rates': self.labor_rates,
            'labor_rates_updated': self.labor_rates_updated,
        }
        try:
            with open("repair_settings.json", 'w') as f:
                json.dump(settings, f)
        except Exception as e:
            print(f"Error saving settings: {e}")

    # ---------------- Labor rates ----------------
    # Kept in the shared data folder so every computer uses the same rates,
    # with a copy in this computer's settings for when the share is down.
    # Whichever copy was changed most recently wins.

    def get_labor_rates_path(self):
        return os.path.join(os.path.dirname(self.save_location), "labor_rates.json")

    def load_labor_rates(self):
        """Pick up the shared labor rates (another computer may have changed
        them). Pushes this computer's copy up if it is the newer one."""
        shared = self._read_json(self.get_labor_rates_path())
        if shared and 'rates' in shared:
            shared_stamp = str(shared.get('updated_at', '') or '')
            if shared_stamp >= (self.labor_rates_updated or ''):
                rates = clean_labor_rates(shared.get('rates'))
                if rates != self.labor_rates or shared_stamp != self.labor_rates_updated:
                    self.labor_rates = rates
                    self.labor_rates_updated = shared_stamp
                    self.save_settings()
                return
        # Share has nothing, or an older copy: offer ours if it was ever set
        if self.labor_rates_updated:
            self._write_shared_labor_rates()

    def _write_shared_labor_rates(self):
        try:
            path = self.get_labor_rates_path()
            tmp = path + '.tmp'
            with open(tmp, 'w') as f:
                json.dump({'rates': self.labor_rates,
                           'updated_at': self.labor_rates_updated}, f, indent=2)
            os.replace(tmp, path)
            return True
        except Exception as e:
            print(f"Could not write shared labor rates: {e}")
            return False

    def save_labor_rates(self, rates):
        """Store new labor rates. Returns True if they reached the share."""
        self.labor_rates = clean_labor_rates(rates)
        self.labor_rates_updated = datetime.now().isoformat()
        self.save_settings()
        shared_ok = self._write_shared_labor_rates()
        for win in list(getattr(self, '_service_order_windows', []) or []):
            try:
                win.refresh_labor_rate_choices()
            except Exception as e:
                print(f"Could not refresh labor rates in an open service order: {e}")
        return shared_ok
    
    def get_local_backup_path(self):
        """Get path for local backup file (stored in user's local AppData)"""
        try:
            local_dir = os.path.join(os.path.expanduser("~"), ".repair_management_suite")
            os.makedirs(local_dir, exist_ok=True)
            return os.path.join(local_dir, "local_backup.json")
        except:
            # Fallback to same directory as executable
            try:
                import sys
                base = os.path.dirname(os.path.abspath(sys.argv[0]))
                return os.path.join(base, "local_backup.json")
            except:
                return "local_backup.json"
    
    def is_server_reachable(self):
        """Check if the save location is reachable"""
        try:
            save_dir = os.path.dirname(self.save_location)
            if not save_dir:
                return True
            # Try to access the directory
            return os.path.exists(save_dir) and os.access(save_dir, os.W_OK)
        except:
            return False
    
    def _stamp_changed_records(self):
        """Mark records whose contents changed since the last save.

        Doing it here means every edit anywhere in the app gets a timestamp
        without each one having to remember to set it - and the timestamp is
        what decides the winner when two computers have both touched a job.
        """
        if not hasattr(self, '_record_snapshot'):
            self._record_snapshot = {}
        now = datetime.now().isoformat()
        for m in list(self.active_machines) + list(self.shipped_machines):
            key = record_key(m)
            try:
                body = json.dumps({k: v for k, v in m.items()
                                   if k != 'updated_at'}, sort_keys=True, default=str)
            except Exception:
                body = str(m)
            if self._record_snapshot.get(key) != body:
                m['updated_at'] = now
                self._record_snapshot[key] = body

    def save_data(self):
        """Save locally, then merge-and-write to the shared copy.

        The server write re-reads what is there first and merges, so two
        computers saving within the same 30-second window cannot wipe out
        each other's jobs.
        """
        ensure_record_ids(self.active_machines + self.shipped_machines)
        self._stamp_changed_records()

        if not hasattr(self, '_synced_ids'):
            self._synced_ids = set()

        local_payload = {
            'active_machines': self.active_machines,
            'shipped_machines': self.shipped_machines,
            'synced_ids': sorted(self._synced_ids),
            'last_saved': datetime.now().isoformat()
        }
        
        # Always save local backup first (fast, always works)
        try:
            local_path = self.get_local_backup_path()
            with open(local_path, 'w') as f:
                json.dump(local_payload, f, indent=2)
        except Exception as e:
            print(f"Local backup save error: {e}")
        
        # Try to save to server
        server_ok = False
        error_msg = ""
        try:
            os.makedirs(os.path.dirname(self.save_location), exist_ok=True)

            # Fold in anything another computer wrote since we last looked
            existing = self._read_json(self.save_location)
            if existing:
                ex_active = existing.get('active_machines', []) or []
                ex_shipped = existing.get('shipped_machines', []) or []
                ex_all = ex_active + ex_shipped
                adopt_record_ids(ex_all, self.active_machines + self.shipped_machines)
                ensure_record_ids(ex_all)
                merged_a, merged_s, _stats = merge_machine_data(
                    ex_active, ex_shipped,
                    self.active_machines, self.shipped_machines,
                    self._synced_ids)
                self.active_machines, self.shipped_machines = merged_a, merged_s
                # Never renumber here - just ask, if the merge produced a clash
                self._schedule_conflict_prompt()

            data = {
                'active_machines': self.active_machines,
                'shipped_machines': self.shipped_machines,
                'last_saved': datetime.now().isoformat()
            }
            tmp = self.save_location + '.tmp'
            with open(tmp, 'w') as f:
                json.dump(data, f, indent=2)
            os.replace(tmp, self.save_location)
            server_ok = True

            # Everything now on the server counts as synced, which is how a
            # later deletion elsewhere is told apart from new offline work.
            self._synced_ids = {record_key(m) for m in
                                self.active_machines + self.shipped_machines}
            try:
                with open(self.get_local_backup_path(), 'w') as f:
                    json.dump({'active_machines': self.active_machines,
                               'shipped_machines': self.shipped_machines,
                               'synced_ids': sorted(self._synced_ids),
                               'last_saved': data['last_saved']}, f, indent=2)
            except Exception as e:
                print(f"Local backup refresh error: {e}")
        except Exception as e:
            error_msg = str(e)
        
        # Update status
        if hasattr(self, 'status_label') and self.status_label:
            if server_ok:
                if self._was_offline:
                    self.status_label.config(text=f"✅ Reconnected & synced at {datetime.now().strftime('%H:%M:%S')}",
                                            fg='#27ae60')
                    self._was_offline = False
                else:
                    self.status_label.config(text=f"✅ Saved at {datetime.now().strftime('%H:%M:%S')}",
                                            fg='#27ae60')
            else:
                self._was_offline = True
                self.status_label.config(text=f"⚠ OFFLINE MODE - Saved locally at {datetime.now().strftime('%H:%M:%S')}",
                                        fg='#e67e22')
        
        # Update connection indicator if it exists
        self._update_connection_indicator(server_ok)
        
        if not server_ok:
            print(f"Server save failed (using local backup): {error_msg}")
        
        return server_ok
    
    def _update_connection_indicator(self, connected):
        """Update the visual connection status indicator"""
        if hasattr(self, 'connection_label') and self.connection_label:
            if connected:
                self.connection_label.config(text="🟢 Online", fg='#27ae60')
            else:
                self.connection_label.config(text="🔴 Offline", fg='#e74c3c')
    
    def start_offline_sync_watcher(self):
        """Start background timer to check server connection and sync when back online"""
        def check_and_sync():
            try:
                if self._was_offline and self.is_server_reachable():
                    # Server is back! Try to sync
                    print("Server reachable again - attempting sync...")
                    # Push up any service orders saved while offline
                    try:
                        pushed, _ = push_offline_service_orders(self.service_orders_dir)
                        if pushed:
                            print(f"Uploaded {pushed} offline service order(s)")
                    except Exception as e:
                        print(f"Offline service order sync failed: {e}")
                    if self.save_data():
                        print("Sync successful")
                        # Show notification if GUI is up
                        if hasattr(self, 'status_label') and self.status_label:
                            self.status_label.config(
                                text=f"✅ Back online! Data synced at {datetime.now().strftime('%H:%M:%S')}",
                                fg='#27ae60'
                            )
                # Update indicator regardless
                self._update_connection_indicator(self.is_server_reachable())
            except Exception as e:
                print(f"Sync check error: {e}")
            
            # Schedule next check
            try:
                self.root.after(15000, check_and_sync)  # Every 15 seconds
            except:
                pass
        
        # Start first check after 15 seconds
        try:
            self.root.after(15000, check_and_sync)
        except:
            pass
    
    def manual_sync(self):
        """Manually trigger a sync attempt"""
        if self.is_server_reachable():
            if self.save_data():
                messagebox.showinfo("Sync Complete",
                    f"Data successfully synced to server.\n\n"
                    f"Active: {len(self.active_machines)} machines\n"
                    f"Shipped: {len(self.shipped_machines)} machines")
            else:
                messagebox.showerror("Sync Failed",
                    "Server appeared reachable but save failed. Check the log.")
        else:
            local_path = self.get_local_backup_path()
            messagebox.showwarning("Server Unreachable",
                f"Cannot reach the server at:\n{os.path.dirname(self.save_location)}\n\n"
                f"Your data is safe in local backup:\n{local_path}\n\n"
                f"The app will automatically sync when the connection is restored.")
    
    def _read_json(self, path):
        """Read a data file. Returns {} if it is missing or unreadable."""
        try:
            if not os.path.exists(path):
                return {}
            with open(path, 'r') as f:
                return json.load(f) or {}
        except Exception as e:
            print(f"Could not read {path}: {e}")
            return {}

    def load_data(self):
        """Load the log by MERGING the shared copy with this computer's copy.

        The old version replaced memory with the server copy whenever the
        share was reachable, which silently threw away anything added while
        offline. Nothing is dropped now unless it was genuinely deleted
        somewhere else.
        """
        server = self._read_json(self.save_location)
        local = self._read_json(self.get_local_backup_path())

        server_active = server.get('active_machines', []) or []
        server_shipped = server.get('shipped_machines', []) or []
        local_active = local.get('active_machines', []) or []
        local_shipped = local.get('shipped_machines', []) or []

        # Which records this computer has previously seen on the server.
        # That is how an offline addition is told apart from a deletion.
        self._synced_ids = set(local.get('synced_ids', []) or [])

        server_reachable = self.is_server_reachable()
        have_server_file = os.path.exists(self.save_location)

        # Give every record a stable id before anything is matched up.
        # Old jobs get an id derived from their estimate number (the same on
        # both copies), or adopt the id the other copy already gave them.
        server_all = server_active + server_shipped
        local_all = local_active + local_shipped
        adopt_record_ids(local_all, server_all)
        adopt_record_ids(server_all, local_all)
        ensure_record_ids(server_all)
        ensure_record_ids(local_all)

        if not have_server_file and not server_reachable:
            # Offline with no server copy to compare against: trust local
            self.active_machines = local_active
            self.shipped_machines = local_shipped
            self._was_offline = True
            print(f"Offline start - {len(self.active_machines)} active, "
                  f"{len(self.shipped_machines)} shipped from local copy")
        elif not have_server_file:
            self.active_machines = local_active
            self.shipped_machines = local_shipped
        else:
            self.active_machines, self.shipped_machines, stats = merge_machine_data(
                server_active, server_shipped,
                local_active, local_shipped,
                self._synced_ids)
            self._last_merge_stats = stats
            if stats['kept_offline'] or stats['local_wins'] or stats['deleted_elsewhere']:
                print(f"Merge: {stats}")
            if not server_reachable:
                self._was_offline = True

        ensure_record_ids(self.active_machines + self.shipped_machines)

        # Shared estimate numbers are NOT fixed here. They are listed and the
        # user decides what to do once the window is up (report_sync_results).
        self._pending_conflicts = find_estimate_conflicts(
            self.active_machines, self.shipped_machines)
        for c in self._pending_conflicts:
            print(f"Estimate number needs a decision: {c['number']} "
                  f"({len(c['members'])} jobs, {c['kind']})")

        # Records created before Field/Shop tracking existed default to Shop
        self._migrate_service_location()

        # Remember what each record looks like now, so the first save only
        # timestamps records that were actually edited. Stamping everything
        # on the first save made this computer's copy win every merge.
        self._seed_record_snapshot()

    def _seed_record_snapshot(self):
        self._record_snapshot = {}
        for m in list(self.active_machines) + list(self.shipped_machines):
            try:
                body = json.dumps({k: v for k, v in m.items()
                                   if k != 'updated_at'}, sort_keys=True, default=str)
            except Exception:
                body = str(m)
            self._record_snapshot[record_key(m)] = body
    
    def _service_order_dirs(self):
        dirs = []
        if self.service_orders_dir:
            dirs.append(self.service_orders_dir)
        try:
            dirs.append(get_offline_service_orders_dir())
        except Exception:
            pass
        return dirs

    @staticmethod
    def _service_order_owner(path):
        """record_id inside a service order file: '' if it has none, None if
        the file can't be read."""
        try:
            with open(path, 'r') as f:
                return str(json.load(f).get('record_id', '') or '')
        except Exception:
            return None

    @staticmethod
    def _set_aside(path, suffix):
        """Rename a file out of the way instead of deleting it."""
        dst = path + suffix
        if os.path.exists(dst):
            dst = f"{path}{suffix}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        os.replace(path, dst)
        return dst

    def _rename_service_order_files(self, old_number, new_number, record_id):
        """Follow a renumbered job's service order files to the new name.

        Only moves a file that names this job's record_id. A file with no
        record_id stays with the number it was filed under - moving those
        is what took service orders away from the jobs they belonged to.
        """
        if not old_number or not new_number or not record_id:
            return
        for d in self._service_order_dirs():
            try:
                src = os.path.join(d, f"{old_number}_service_order.json")
                if not os.path.exists(src):
                    continue
                try:
                    with open(src, 'r') as f:
                        blob = json.load(f)
                    owner = str(blob.get('record_id', '') or '')
                    if owner != str(record_id):
                        continue
                    blob['linked_estimate'] = new_number
                    if record_id:
                        blob['record_id'] = str(record_id)
                except Exception:
                    continue
                dst = os.path.join(d, f"{new_number}_service_order.json")
                if os.path.exists(dst):
                    continue
                with open(dst, 'w') as f:
                    json.dump(blob, f, indent=2)
                self._set_aside(src, ".renamed")
                print(f"Service order {old_number} -> {new_number} in {d}")
            except Exception as e:
                print(f"Could not rename service order {old_number}: {e}")

    def report_sync_results(self):
        """Tell the user plainly when a merge recovered or renumbered work."""
        stats = getattr(self, '_last_merge_stats', {}) or {}

        recovered = stats.get('kept_offline', 0)
        if recovered and self.status_label:
            self.status_label.config(
                text=f"✅ {recovered} job(s) saved offline were restored and uploaded",
                fg='#27ae60')

        conflicts = getattr(self, '_pending_conflicts', []) or []
        self._pending_conflicts = []
        if conflicts:
            self.prompt_estimate_conflicts(conflicts)

    # ------------------------------------------------------------
    #  Duplicate estimate numbers - always ask, never auto-fix
    # ------------------------------------------------------------

    def _schedule_conflict_prompt(self):
        """Called after a save merges in another computer's data. Shows the
        prompt shortly afterwards (never in the middle of a save)."""
        if self._conflict_prompt_open or self._conflict_prompt_scheduled:
            return
        pending = [c for c in find_estimate_conflicts(self.active_machines,
                                                      self.shipped_machines)
                   if conflict_signature(c) not in self._conflicts_deferred]
        if not pending or not getattr(self, 'status_label', None):
            return
        self._conflict_prompt_scheduled = True

        def show():
            self._conflict_prompt_scheduled = False
            try:
                if self.root.winfo_exists():
                    self.prompt_estimate_conflicts()
            except Exception as e:
                print(f"Could not show duplicate-number prompt: {e}")
        try:
            self.root.after(500, show)
        except Exception:
            self._conflict_prompt_scheduled = False

    def prompt_estimate_conflicts(self, conflicts=None):
        """List jobs that share an estimate number and let the user decide.
        Nothing is renumbered or removed without a choice made here."""
        if self._conflict_prompt_open:
            return
        if conflicts is None:
            conflicts = find_estimate_conflicts(self.active_machines, self.shipped_machines)
        conflicts = [c for c in conflicts
                     if conflict_signature(c) not in self._conflicts_deferred]
        if not conflicts:
            return
        self._conflict_prompt_open = True
        try:
            choices = self._ask_conflict_actions(conflicts)
        finally:
            self._conflict_prompt_open = False

        if choices is None:
            choices = [(c, 'later') for c in conflicts]
        self._apply_conflict_actions(choices)

    def _ask_conflict_actions(self, conflicts):
        """Modal window: one decision per estimate number.
        Returns [(conflict, 'merge'|'separate'|'later')] or None if closed."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Duplicate Estimate Numbers")
        dlg.configure(bg='#2c3e50')
        dlg.geometry("940x640")
        dlg.minsize(700, 420)
        dlg.transient(self.root)

        tk.Label(dlg, text=f"{len(conflicts)} estimate number(s) need your decision",
                 font=('Arial', 15, 'bold'), fg='white', bg='#2c3e50'
                 ).pack(anchor='w', padx=16, pady=(14, 2))
        tk.Label(dlg, justify='left', wraplength=900, fg='#bdc3c7', bg='#2c3e50',
                 font=('Arial', 10),
                 text=("These jobs share an estimate number, or were renumbered "
                       "automatically by an earlier version ([auto-copy]). Nothing "
                       "has been changed. Choose what should happen to each one, then "
                       "click Apply. 'Decide later' leaves it exactly as it is "
                       "and asks again next time the program starts.")
                 ).pack(anchor='w', padx=16, pady=(0, 8))

        # Set-all row
        bar = tk.Frame(dlg, bg='#2c3e50')
        bar.pack(fill='x', padx=16, pady=(0, 6))
        tk.Label(bar, text="Set all to:", fg='#ecf0f1', bg='#2c3e50',
                 font=('Arial', 10, 'bold')).pack(side='left')

        # Scrollable list of groups
        holder = tk.Frame(dlg, bg='#2c3e50')
        holder.pack(fill='both', expand=True, padx=16)
        canvas = tk.Canvas(holder, bg='#2c3e50', highlightthickness=0)
        sb = ttk.Scrollbar(holder, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg='#2c3e50')
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        win_id = canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.bind('<Configure>', lambda e: canvas.itemconfigure(win_id, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        def on_wheel(event):
            if event.delta:
                canvas.yview_scroll(int(-event.delta / 120) or (-1 if event.delta > 0 else 1), 'units')
            elif getattr(event, 'num', None) == 4:
                canvas.yview_scroll(-1, 'units')
            elif getattr(event, 'num', None) == 5:
                canvas.yview_scroll(1, 'units')
        dlg.bind_all('<MouseWheel>', on_wheel)
        dlg.bind_all('<Button-4>', on_wheel)
        dlg.bind_all('<Button-5>', on_wheel)

        rows = []
        for c in conflicts:
            identical = c['identical']
            box = tk.Frame(inner, bg='#34495e', padx=12, pady=8)
            box.pack(fill='x', pady=5)

            head = tk.Frame(box, bg='#34495e')
            head.pack(fill='x')
            tk.Label(head, text=f"Estimate {c['number']}", font=('Arial', 12, 'bold'),
                     fg='white', bg='#34495e').pack(side='left')
            count = len(c['members'])
            verdict = (f"Details match - looks like the same job entered {count} times"
                       if identical else "Details differ - looks like different jobs")
            tk.Label(head, text=verdict, font=('Arial', 10, 'bold'),
                     fg='#2ecc71' if identical else '#f39c12',
                     bg='#34495e').pack(side='left', padx=12)

            for m, where in c['members']:
                make_model = f"{m.get('make', '')} {m.get('model', '')}".strip() or '-'
                tag = "  [auto-copy]" if m.get('renumbered_from') else ""
                line = (f"{m.get('estimate_number', '')}{tag}  ·  "
                        f"{m.get('customer_name', '') or '-'}  ·  "
                        f"S/N {m.get('serial_number', '') or '-'}  ·  "
                        f"{make_model}  ·  in {m.get('arrival_date', '') or '-'}  ·  "
                        f"{where}")
                tk.Label(box, text=line, font=('Consolas', 9), fg='#ecf0f1',
                         bg='#34495e', anchor='w', justify='left', wraplength=840
                         ).pack(fill='x', padx=(8, 0))

            extra = count - 1
            copies = "copy" if extra == 1 else f"{extra} copies"
            if c['kind'] == 'renumbered':
                separate_text = ("Different jobs - keep " + ("both" if count == 2 else "all")
                                 + " with the numbers they have now")
            else:
                separate_text = "Different jobs - give the newer one(s) the next free number"
            options = [
                ('merge', f"Same job - keep one, remove the extra {copies} "
                          "(service order goes back under this number)"),
                ('separate', separate_text),
                ('later', "Decide later - change nothing"),
            ]
            var = tk.StringVar(value='merge' if identical else 'later')
            for key, text in options:
                tk.Radiobutton(box, text=text, value=key, variable=var,
                               fg='#ecf0f1', bg='#34495e', selectcolor='#2c3e50',
                               activebackground='#34495e', activeforeground='white',
                               font=('Arial', 10), anchor='w', bd=0,
                               highlightthickness=0
                               ).pack(fill='x', padx=(8, 0), pady=1)
            rows.append((c, var))

        def set_all(key):
            for _, var in rows:
                var.set(key)

        def set_suggested():
            for c, var in rows:
                var.set('merge' if c['identical'] else 'later')

        for text, cmd in (("Suggested", set_suggested),
                          ("Same job", lambda: set_all('merge')),
                          ("Different jobs", lambda: set_all('separate')),
                          ("Decide later", lambda: set_all('later'))):
            tk.Button(bar, text=text, command=cmd, bg='#7f8c8d', fg='white',
                      font=('Arial', 9)).pack(side='left', padx=4)

        result = {'value': None}

        def close():
            try:
                dlg.unbind_all('<MouseWheel>')
                dlg.unbind_all('<Button-4>')
                dlg.unbind_all('<Button-5>')
            except Exception:
                pass
            dlg.destroy()

        def apply():
            result['value'] = [(c, var.get()) for c, var in rows]
            close()

        foot = tk.Frame(dlg, bg='#2c3e50')
        foot.pack(fill='x', padx=16, pady=12)
        tk.Button(foot, text="Apply", command=apply, bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'), width=14).pack(side='right')
        tk.Button(foot, text="Decide Later (change nothing)", command=close,
                  bg='#7f8c8d', fg='white', font=('Arial', 11)
                  ).pack(side='right', padx=8)
        dlg.protocol("WM_DELETE_WINDOW", close)

        dlg.update_idletasks()
        x = max(0, (dlg.winfo_screenwidth() // 2) - (dlg.winfo_width() // 2))
        y = max(0, (dlg.winfo_screenheight() // 2) - (dlg.winfo_height() // 2))
        dlg.geometry(f"+{x}+{y}")
        try:
            dlg.wait_visibility()
            dlg.grab_set()
        except tk.TclError:
            pass
        dlg.focus_set()
        dlg.wait_window()
        return result['value']

    def _apply_conflict_actions(self, choices):
        """Carry out what the user picked in the duplicate-number window."""
        merged, separated, deferred = [], [], 0
        for c, action in choices:
            try:
                if action == 'merge':
                    res = merge_conflict(self.active_machines, self.shipped_machines, c)
                    if res:
                        keep, removed, other_numbers = res
                        self._merge_service_order_files(c['number'], keep, removed,
                                                        other_numbers)
                        merged.append((c['number'], keep, len(removed)))
                elif action == 'separate':
                    for rec, old, new in renumber_conflict(
                            self.active_machines, self.shipped_machines, c):
                        self._rename_service_order_files(old, new, rec.get('record_id'))
                        separated.append((rec, old, new))
                    if c['kind'] == 'renumbered':
                        separated.append((None, c['number'], None))
                else:
                    self._conflicts_deferred.add(conflict_signature(c))
                    deferred += 1
            except Exception as e:
                print(f"Could not apply decision for {c.get('number')}: {e}")

        if not merged and not separated:
            if deferred and self.status_label:
                self.status_label.config(
                    text=f"⚠ {deferred} duplicate estimate number(s) left as they are "
                         "- you'll be asked again next start",
                    fg='#e67e22')
            return

        self.save_data()
        try:
            self.refresh_tables()
            self.update_estimate_display()
        except Exception as e:
            print(f"Refresh after duplicate fix failed: {e}")

        parts = []
        if merged:
            n = sum(k for _, _, k in merged)
            parts.append(f"Removed {n} extra cop{'y' if n == 1 else 'ies'} across "
                         f"{len(merged)} estimate number(s).")
        renumbered = [(r, o, n) for r, o, n in separated if r is not None]
        kept_both = sum(1 for r, _, _ in separated if r is None)
        if renumbered:
            lines = "\n".join(f"   {o}  ->  {n}   ({r.get('customer_name', '')})"
                              for r, o, n in renumbered[:10])
            more = f"\n   ...and {len(renumbered) - 10} more" if len(renumbered) > 10 else ""
            parts.append(f"Renumbered {len(renumbered)} job(s):\n{lines}{more}")
        if kept_both:
            parts.append(f"Kept {kept_both} group(s) as separate jobs with their current numbers.")
        if deferred:
            parts.append(f"Left {deferred} for later.")
        messagebox.showinfo("Duplicate Estimate Numbers", "\n\n".join(parts))

    def _merge_service_order_files(self, number, keep, removed, other_numbers):
        """Bring a merged job's service order back under its estimate number.

        An earlier version moved a job's service order to the copy's new
        number and left the original as '<number>_service_order.json.renamed'.
        The most recently saved of those becomes the job's service order
        again. The rest are set aside with a '.duplicate' suffix, never
        deleted.
        """
        group_ids = {str(m.get('record_id') or '') for m in [keep] + list(removed)} - {''}
        keep_id = str(keep.get('record_id') or '')
        for d in self._service_order_dirs():
            try:
                target = os.path.join(d, f"{number}_service_order.json")
                paths = [target, target + '.renamed'] + [
                    os.path.join(d, f"{n}_service_order.json") for n in sorted(other_numbers)]
                candidates = []
                for p in paths:
                    if not os.path.exists(p):
                        continue
                    owner = self._service_order_owner(p)
                    if owner is None:
                        continue                      # unreadable - leave it alone
                    if owner and owner not in group_ids:
                        continue                      # another job's paperwork
                    candidates.append(p)
                if not candidates:
                    continue
                if candidates == [target]:
                    owner = self._service_order_owner(target)
                    if not owner or owner == keep_id:
                        continue                      # already correct

                def saved_key(p):
                    try:
                        mtime = os.path.getmtime(p)
                    except Exception:
                        mtime = 0
                    return (read_service_order_stamp(p), mtime)
                newest = max(candidates, key=saved_key)

                with open(newest, 'r') as f:
                    blob = json.load(f)
                blob['linked_estimate'] = number
                if keep_id:
                    blob['record_id'] = keep_id

                if os.path.exists(target) and newest != target:
                    self._set_aside(target, '.before-merge')
                tmp = target + '.tmp'
                with open(tmp, 'w') as f:
                    json.dump(blob, f, indent=2)
                os.replace(tmp, target)
                for p in candidates:
                    if p != target and os.path.exists(p):
                        self._set_aside(p, '.duplicate')
                print(f"Service order for {number} restored from "
                      f"{os.path.basename(newest)} in {d}")
            except Exception as e:
                print(f"Could not merge service order files for {number}: {e}")

    def _migrate_service_location(self):
        """Give older records a Field/Shop value so filters behave predictably."""
        changed = 0
        for machine in list(self.active_machines) + list(self.shipped_machines):
            if not machine.get('service_location'):
                machine['service_location'] = 'Shop'
                changed += 1
        if changed:
            print(f"Set {changed} existing record(s) to Shop (Field/Shop default)")
        return changed
    
    def schedule_auto_save(self):
        """Schedule the next auto-save"""
        # Hold off while the duplicate-number window is waiting on the user
        if not self._conflict_prompt_open:
            self.save_data()
        self.root.after(self.auto_save_interval, self.schedule_auto_save)
    
    def create_widgets(self):
        """Create the main GUI"""
        # Header
        header_frame = tk.Frame(self.root, bg='#34495e', height=80)
        header_frame.pack(fill='x', padx=10, pady=5)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="Machine Repair Management System", 
                              font=('Arial', 20, 'bold'), fg='white', bg='#34495e')
        title_label.pack(pady=(12, 0))
        
        tk.Label(header_frame, text=f"Version {APP_VERSION}", font=('Arial', 9),
                fg='#95a5a6', bg='#34495e').pack()
        
        # Status bar
        status_frame = tk.Frame(self.root, bg='#2c3e50')
        status_frame.pack(fill='x', padx=10)
        
        self.status_label = tk.Label(status_frame, text="Ready", font=('Arial', 10), 
                                    fg='#ecf0f1', bg='#2c3e50')
        self.status_label.pack(side='left')
        
        # Connection status indicator
        self.connection_label = tk.Label(status_frame, text="🟢 Online", font=('Arial', 10, 'bold'),
                                        fg='#27ae60', bg='#2c3e50')
        self.connection_label.pack(side='left', padx=15)
        
        # Sync button (manual retry)
        sync_btn = tk.Button(status_frame, text="🔄 Sync Now", command=self.manual_sync,
                            bg='#2980b9', fg='white', font=('Arial', 9, 'bold'))
        sync_btn.pack(side='left', padx=5)
        
        # Update notice - hidden until a newer build is found
        self.update_label = tk.Label(status_frame, text="", font=('Arial', 10, 'bold'),
                                    fg='#f39c12', bg='#2c3e50')
        
        update_btn = tk.Button(status_frame, text="⬆ Check for Updates",
                              command=self.check_for_updates_manual,
                              bg='#8e44ad', fg='white', font=('Arial', 9, 'bold'))
        update_btn.pack(side='left', padx=5)
        
        # Printer setup button
        printer_btn = tk.Button(status_frame, text="Setup Printer", 
                               command=self.setup_printer, bg='#e67e22', fg='white', 
                               font=('Arial', 10, 'bold'))
        printer_btn.pack(side='right', padx=5)
        
        # Manual barcode lookup button
        scan_btn = tk.Button(status_frame, text="🔍 Lookup Barcode",
                            command=self.manual_barcode_lookup, bg='#16a085', fg='white',
                            font=('Arial', 10, 'bold'))
        scan_btn.pack(side='right', padx=5)
        
        # Webcam barcode scanner button
        webcam_btn = tk.Button(status_frame, text="📷 Scan with Webcam",
                              command=self.open_webcam_scanner, bg='#8e44ad', fg='white',
                              font=('Arial', 10, 'bold'))
        webcam_btn.pack(side='right', padx=5)
        
        # Load data file button
        load_btn = tk.Button(status_frame, text="Load Data File", 
                            command=self.load_data_file, bg='#3498db', fg='white', 
                            font=('Arial', 10, 'bold'))
        load_btn.pack(side='right', padx=5)
        
        # Save location button
        location_btn = tk.Button(status_frame, text="Set Save Location", 
                               command=self.set_save_location, bg='#9b59b6', fg='white', 
                               font=('Arial', 10, 'bold'))
        location_btn.pack(side='right', padx=5)
        
        # Attachments location button
        attach_btn = tk.Button(status_frame, text="Set Attachments Folder", 
                              command=self.set_attachments_location, bg='#e74c3c', fg='white', 
                              font=('Arial', 10, 'bold'))
        attach_btn.pack(side='right', padx=5)
        
        # Main content area
        main_frame = tk.Frame(self.root, bg='#2c3e50')
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Left panel - Input form
        self.create_input_panel(main_frame)
        
        # Right panel - Notebook with tabs
        self.create_tabs_panel(main_frame)
    
    def setup_printer(self):
        """Setup Brother QL-570 printer"""
        if not BARCODE_AVAILABLE:
            error_details = '\n'.join(BARCODE_IMPORT_ERRORS) if BARCODE_IMPORT_ERRORS else 'Unknown import error'
            messagebox.showerror("Error", 
                f"Barcode printing unavailable.\n\n"
                f"Import errors:\n{error_details}\n\n"
                "Please install:\n"
                "pip install brother_ql pillow python-barcode pyusb")
            return
        
        setup_window = tk.Toplevel(self.root)
        setup_window.title("Printer Setup")
        setup_window.geometry("500x300")
        setup_window.configure(bg='#2c3e50')
        setup_window.transient(self.root)
        setup_window.grab_set()
        
        # Center the window
        setup_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 200, self.root.winfo_rooty() + 100))
        
        tk.Label(setup_window, text="Brother QL-570 Printer Setup", 
                font=('Arial', 16, 'bold'), fg='white', bg='#2c3e50').pack(pady=20)
        
        tk.Label(setup_window, text="Enter printer identifier:", 
                font=('Arial', 12), fg='#ecf0f1', bg='#2c3e50').pack(pady=10)
        
        # Info text
        info_text = tk.Label(setup_window, 
                            text="Examples:\n"
                                 "Windows: usb://0x04f9:0x2028\n"
                                 "macOS: usb://0x04f9:0x2028\n"
                                 "Linux: /dev/usb/lp0\n\n"
                                 "Or use 'file://labels.bin' for testing",
                            font=('Arial', 9), fg='#bdc3c7', bg='#2c3e50', justify='left')
        info_text.pack(pady=10)
        
        printer_var = tk.StringVar(value=self.printer_identifier or "usb://0x04f9:0x2028")
        printer_entry = tk.Entry(setup_window, textvariable=printer_var, 
                                font=('Arial', 11), width=40)
        printer_entry.pack(pady=10)
        
        def save_printer():
            self.printer_identifier = printer_var.get().strip()
            if self.printer_identifier:
                self.save_settings()
                messagebox.showinfo("Success", "Printer setup saved!")
                setup_window.destroy()
            else:
                messagebox.showwarning("Warning", "Please enter a printer identifier")
        
        btn_frame = tk.Frame(setup_window, bg='#2c3e50')
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="Save", command=save_printer,
                 bg='#27ae60', fg='white', font=('Arial', 11, 'bold'), width=10).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Cancel", command=setup_window.destroy,
                 bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'), width=10).pack(side='left', padx=5)
    
    def print_barcode_label(self, machine_idx):
        """Print barcode label for a machine"""
        if not BARCODE_AVAILABLE:
            error_details = '\n'.join(BARCODE_IMPORT_ERRORS) if BARCODE_IMPORT_ERRORS else 'Unknown import error'
            messagebox.showerror("Error", 
                f"Barcode printing unavailable.\n\n"
                f"Import errors:\n{error_details}\n\n"
                "Please install:\n"
                "pip install brother_ql pillow python-barcode pyusb")
            return
        
        if not self.printer_identifier:
            messagebox.showwarning("Warning", 
                "Please setup your printer first.\n\n"
                "Click 'Setup Printer' button in the status bar.")
            return
        
        if machine_idx >= len(self.active_machines):
            messagebox.showerror("Error", "Machine not found")
            return
        
        machine = self.active_machines[machine_idx]
        
        # An entry logged before the machine arrived may have no details yet
        missing = self.missing_machine_details(machine)
        if missing:
            if not messagebox.askyesno(
                    "Machine Details Missing",
                    f"This entry is still missing: {', '.join(missing)}.\n\n"
                    "The label will print with those lines blank.\n\n"
                    "Print anyway?"):
                return
        
        try:
            # Create label image for 29x90mm die-cut labels  
            # Create in PORTRAIT orientation (306 wide x 991 tall) - matches printer expectations exactly
            # Content flows top-to-bottom along the long 90mm dimension
            label_width = 306   # Width in pixels for 29mm
            label_height = 991  # Height in pixels for 90mm
            
            # Create white background
            img = Image.new('RGB', (label_width, label_height), 'white')
            draw = ImageDraw.Draw(img)
            
            # Try to use a TrueType font, fallback to default
            try:
                font_large = ImageFont.truetype("arial.ttf", 28)
                font_small = ImageFont.truetype("arial.ttf", 18)
            except:
                try:
                    font_large = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 28)
                    font_small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 18)
                except:
                    font_large = ImageFont.load_default()
                    font_small = ImageFont.load_default()
            
            # VERTICAL PORTRAIT LAYOUT: Customer name at TOP, Barcode BELOW
            customer_text = machine['customer_name']
            
            # Truncate customer name if too long
            if len(customer_text) > 25:
                customer_text = customer_text[:22] + "..."
            
            # Draw customer name at the top, centered horizontally
            try:
                bbox = draw.textbbox((0, 0), customer_text, font=font_large)
                text_width = bbox[2] - bbox[0]
            except:
                text_width = len(customer_text) * 12
            
            text_x = (label_width - text_width) // 2
            draw.text((text_x, 20), customer_text, fill='black', font=font_large)
            
            # Generate barcode - try multiple methods for best quality
            estimate_num = machine['estimate_number']
            CODE128 = barcode.get_barcode_class('code128')
            
            barcode_generated = False
            barcode_target_width = label_width - 20  # Max width available
            
            # Method 1: Try ImageWriter (produces real Code128 barcodes - BEST)
            try:
                from barcode.writer import ImageWriter
                
                # Calculate module_width to fit label width
                # Code128 uses about 11 modules per character + start/stop/checksum (~24 extra modules)
                num_chars = len(estimate_num)
                approx_modules = (num_chars * 11) + 35  # rough estimate with quiet zones
                target_module_px = max(1.0, barcode_target_width / approx_modules)
                # Convert pixels to mm (at 300 DPI: 1mm ≈ 11.8px)
                module_width_mm = target_module_px / 11.8
                # Clamp to reasonable range for scannability
                module_width_mm = max(0.2, min(0.6, module_width_mm))
                
                writer_options = {
                    'module_width': module_width_mm,
                    'module_height': 20.0,     # Height in mm (will be scaled by DPI)
                    'quiet_zone': 2.0,         # White space around barcode in mm
                    'font_size': 14,           # Text size
                    'text_distance': 3.0,      # Space between bars and text
                    'background': 'white',
                    'foreground': 'black',
                    'write_text': True,
                    'dpi': 300
                }
                
                writer = ImageWriter()
                code = CODE128(estimate_num, writer=writer)
                barcode_io = BytesIO()
                code.write(barcode_io, writer_options)
                barcode_io.seek(0)
                barcode_img = Image.open(barcode_io)
                barcode_generated = True
                print(f"✓ Barcode generated with ImageWriter: {barcode_img.size[0]}x{barcode_img.size[1]}px")
                
            except Exception as e:
                print(f"✗ ImageWriter failed: {e}")
            
            # Method 2: Try SVG writer with cairosvg (GOOD)
            if not barcode_generated:
                try:
                    from barcode.writer import SVGWriter
                    import cairosvg
                    
                    barcode_svg_io = BytesIO()
                    code = CODE128(estimate_num, writer=SVGWriter())
                    code.write(barcode_svg_io, {
                        'module_width': 0.4,
                        'module_height': 20,
                        'quiet_zone': 2,
                        'font_size': 14,
                        'text_distance': 3,
                        'write_text': True
                    })
                    
                    barcode_svg_io.seek(0)
                    png_data = cairosvg.svg2png(bytestring=barcode_svg_io.read(), scale=2.0)
                    barcode_img = Image.open(BytesIO(png_data))
                    barcode_generated = True
                    print(f"✓ Barcode generated with SVG+cairosvg: {barcode_img.size[0]}x{barcode_img.size[1]}px")
                    
                except Exception as e:
                    print(f"✗ SVG method failed: {e}")
            
            # Method 3: Manual fallback (FALLBACK)
            if not barcode_generated:
                print("⚠ Using manual barcode drawing (may not scan well)")
                barcode_img = self.draw_barcode_manually(estimate_num, barcode_target_width, 350)
            
            # Scale barcode to fit label - ONLY scale uniformly to preserve bar ratios
            bw, bh = barcode_img.size
            
            # Scale to fit width if too wide, maintaining aspect ratio
            if bw > barcode_target_width:
                scale = barcode_target_width / bw
                new_w = int(bw * scale)
                new_h = int(bh * scale)
                barcode_img = barcode_img.resize((new_w, new_h), Image.LANCZOS)
                bw, bh = new_w, new_h
            
            # If barcode is too short vertically, scale up uniformly
            min_barcode_height = 500
            if bh < min_barcode_height and bw < barcode_target_width:
                scale = min(barcode_target_width / bw, min_barcode_height / bh)
                new_w = int(bw * scale)
                new_h = int(bh * scale)
                # Ensure we don't exceed label width
                if new_w > barcode_target_width:
                    scale = barcode_target_width / bw
                    new_w = int(bw * scale)
                    new_h = int(bh * scale)
                barcode_img = barcode_img.resize((new_w, new_h), Image.LANCZOS)
                bw, bh = new_w, new_h
            
            # Cap height to fit on label
            max_barcode_height = label_height - 180
            if bh > max_barcode_height:
                scale = max_barcode_height / bh
                new_w = int(bw * scale)
                new_h = int(bh * scale)
                barcode_img = barcode_img.resize((new_w, new_h), Image.LANCZOS)
                bw, bh = new_w, new_h
            
            # Center barcode horizontally on label
            barcode_x = (label_width - bw) // 2
            barcode_y = 150  # Below customer name
            
            # Convert barcode to RGB if it's not already
            if barcode_img.mode != 'RGB':
                barcode_img = barcode_img.convert('RGB')
            
            img.paste(barcode_img, (barcode_x, barcode_y))
            
            # Convert to black and white (required for Brother QL)
            img = img.convert('1')
            
            # Print to Brother QL-570
            qlr = BrotherQLRaster('QL-570')
            
            # Convert image to raster format
            # Image is already in portrait (306x991), no rotation needed
            # Content flows top-to-bottom along the long 90mm axis
            instructions = convert(
                qlr=qlr,
                images=[img],
                label='29x90',  # 29mm x 90mm die-cut label (matches 1-1/7" x 3-1/2")
                rotate='0',  # No rotation - image is already correct orientation
                threshold=70.0,
                dither=False,
                compress=False,
                red=False,
                dpi_600=False,
                hq=True,
                cut=True
            )
            
            # Send to printer
            try:
                send(instructions=instructions, printer_identifier=self.printer_identifier, backend_identifier='pyusb', blocking=True)
            except Exception as usb_err:
                # Ignore USB timeout errors - label already printed successfully
                usb_err_str = str(usb_err).lower()
                if 'timeout' in usb_err_str or 'libusb0' in usb_err_str:
                    print(f"USB timeout after print (label printed OK): {usb_err}")
                else:
                    raise  # Re-raise if it's a different error
            
            if self.status_label:
                self.status_label.config(text=f"✅ Label printed for {estimate_num}")
            
            messagebox.showinfo("Success", f"Barcode label printed for {machine['estimate_number']}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to print barcode label:\n\n{str(e)}\n\nMake sure the printer is connected and turned on.")
            print(f"Barcode printing error: {e}")
            import traceback
            traceback.print_exc()
    
    def draw_barcode_manually(self, text, width, height):
        """Manually draw Code128 barcode to avoid library compatibility issues"""
        # Create a simple barcode representation
        img = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(img)
        
        # Create a more realistic barcode pattern with varied bar widths
        bar_width = 5  # Base bar width
        x = 30  # Start position with margin
        
        # Start pattern - 3 bars
        for i in range(3):
            draw.rectangle([x, 15, x + bar_width, height - 65], fill='black')
            x += bar_width * 2
        
        x += bar_width  # Space after start
        
        # Draw bars for each character with varied patterns
        for i, char in enumerate(text):
            char_val = ord(char)
            
            # Pattern 1: Wide black bar + narrow space
            if char_val % 4 == 0:
                draw.rectangle([x, 15, x + bar_width * 2, height - 65], fill='black')
                x += bar_width * 3
            # Pattern 2: Narrow black bar + narrow black bar + space
            elif char_val % 4 == 1:
                draw.rectangle([x, 15, x + bar_width, height - 65], fill='black')
                x += bar_width * 2
                draw.rectangle([x, 15, x + bar_width, height - 65], fill='black')
                x += bar_width * 2
            # Pattern 3: Medium black bar + space
            elif char_val % 4 == 2:
                draw.rectangle([x, 15, x + int(bar_width * 1.5), height - 65], fill='black')
                x += bar_width * 3
            # Pattern 4: Two narrow bars with gap
            else:
                draw.rectangle([x, 15, x + bar_width, height - 65], fill='black')
                x += int(bar_width * 1.5)
                draw.rectangle([x, 15, x + bar_width, height - 65], fill='black')
                x += bar_width * 2
        
        x += bar_width  # Space before end
        
        # End pattern - 3 bars
        for i in range(3):
            draw.rectangle([x, 15, x + bar_width, height - 65], fill='black')
            x += bar_width * 2
        
        # Draw the text below with larger font
        try:
            font = ImageFont.truetype("arial.ttf", 26)
        except:
            try:
                font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 26)
            except:
                font = ImageFont.load_default()
        
        # Center the text
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
        except:
            text_width = len(text) * 12
        
        text_x = (width - text_width) // 2
        draw.text((text_x, height - 50), text, fill='black', font=font)
        
        return img
    
    def create_input_panel(self, parent):
        """Create the input form panel"""
        input_frame = tk.LabelFrame(parent, text="Add New Machine", font=('Arial', 12, 'bold'),
                                   fg='#ecf0f1', bg='#34495e', width=350)
        input_frame.pack(side='left', fill='y', padx=(0, 10), pady=5)
        input_frame.pack_propagate(False)
        
        # Estimate number display
        self.estimate_frame = tk.Frame(input_frame, bg='#f39c12', height=50)
        self.estimate_frame.pack(fill='x', padx=10, pady=10)
        self.estimate_frame.pack_propagate(False)
        
        self.estimate_label = tk.Label(self.estimate_frame, text="Next Estimate #: 090325-001",
                                      font=('Arial', 12, 'bold'), fg='white', bg='#f39c12')
        self.estimate_label.pack(expand=True)
        
        # Input fields
        fields_frame = tk.Frame(input_frame, bg='#34495e')
        fields_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Customer Name
        tk.Label(fields_frame, text="Customer Name:", font=('Arial', 10, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w', pady=(0, 5))
        self.customer_var = tk.StringVar()
        tk.Entry(fields_frame, textvariable=self.customer_var, font=('Arial', 10),
                width=30).pack(fill='x', pady=(0, 10))
        
        # Arrival Date
        tk.Label(fields_frame, text="Date of Arrival:", font=('Arial', 10, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w', pady=(0, 5))
        self.arrival_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(fields_frame, textvariable=self.arrival_var, font=('Arial', 10),
                width=30).pack(fill='x', pady=(0, 10))
        
        # Serial Number
        tk.Label(fields_frame, text="Serial Number:", font=('Arial', 10, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w', pady=(0, 5))
        self.serial_var = tk.StringVar()
        tk.Entry(fields_frame, textvariable=self.serial_var, font=('Arial', 10),
                width=30).pack(fill='x', pady=(0, 10))
        
        # Make
        tk.Label(fields_frame, text="Make:", font=('Arial', 10, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w', pady=(0, 5))
        self.make_var = tk.StringVar()
        tk.Entry(fields_frame, textvariable=self.make_var, font=('Arial', 10),
                width=30).pack(fill='x', pady=(0, 10))
        
        # Model
        tk.Label(fields_frame, text="Model:", font=('Arial', 10, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w', pady=(0, 5))
        self.model_var = tk.StringVar()
        tk.Entry(fields_frame, textvariable=self.model_var, font=('Arial', 10),
                width=30).pack(fill='x', pady=(0, 10))
        
        # Service Location - Field vs Shop
        tk.Label(fields_frame, text="Service Location:", font=('Arial', 10, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w', pady=(0, 5))
        self.location_var = tk.StringVar(value='Shop')
        loc_frame = tk.Frame(fields_frame, bg='#34495e')
        loc_frame.pack(fill='x', pady=(0, 10))
        tk.Radiobutton(loc_frame, text="🏭 Shop", variable=self.location_var, value='Shop',
                      bg='#34495e', fg='#ecf0f1', selectcolor='#2c3e50',
                      activebackground='#34495e', activeforeground='white',
                      font=('Arial', 10, 'bold')).pack(side='left', padx=(0, 15))
        tk.Radiobutton(loc_frame, text="🚚 Field", variable=self.location_var, value='Field',
                      bg='#34495e', fg='#ecf0f1', selectcolor='#2c3e50',
                      activebackground='#34495e', activeforeground='white',
                      font=('Arial', 10, 'bold')).pack(side='left')
        
        # PDF Attachment button
        self.pdf_path_var = tk.StringVar(value="No PDF attached")
        pdf_frame = tk.Frame(fields_frame, bg='#34495e')
        pdf_frame.pack(fill='x', pady=(10, 5))
        
        tk.Label(pdf_frame, text="PDF Attachment:", font=('Arial', 10, 'bold'),
                fg='#ecf0f1', bg='#34495e').pack(anchor='w')
        
        attach_btn = tk.Button(pdf_frame, text="Attach PDF", command=self.attach_pdf,
                              bg='#e67e22', fg='white', font=('Arial', 9, 'bold'))
        attach_btn.pack(side='left', pady=(5, 0))
        
        pdf_label = tk.Label(pdf_frame, textvariable=self.pdf_path_var, font=('Arial', 8),
                            fg='#bdc3c7', bg='#34495e', wraplength=250)
        pdf_label.pack(side='left', padx=(10, 0), pady=(5, 0))
        
        # Buttons
        btn_frame = tk.Frame(fields_frame, bg='#34495e')
        btn_frame.pack(fill='x', pady=10)
        
        add_btn = tk.Button(btn_frame, text="Add Machine", command=self.add_machine,
                           bg='#3498db', fg='white', font=('Arial', 12, 'bold'), height=2)
        add_btn.pack(fill='x', pady=(0, 5))
        
        export_btn = tk.Button(btn_frame, text="Export Options", command=self.show_export_menu,
                              bg='#27ae60', fg='white', font=('Arial', 10, 'bold'))
        export_btn.pack(fill='x')
        
        # Bind arrival date change to update estimate
        # Use trace_add for Tcl 9 compatibility
        try:
            self.arrival_var.trace_add('write', lambda *args: self.update_estimate_display())
        except AttributeError:
            # Fallback for older Python/Tkinter
            self.arrival_var.trace('w', self.update_estimate_display)
        
        # Reset PDF attachment variable
        self.current_pdf_path = None
    
    def attach_pdf(self):
        """Attach a PDF file to the machine"""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.current_pdf_path = file_path
            self.pdf_path_var.set(os.path.basename(file_path))
    
    def set_attachments_location(self):
        """Set the attachments directory"""
        new_location = filedialog.askdirectory(
            title="Select Attachments Folder",
            initialdir=self.attachments_dir
        )
        if new_location:
            self.attachments_dir = new_location
            self.save_settings()
            if self.status_label:
                self.status_label.config(text=f"Attachments folder updated: {os.path.basename(new_location)}")
    
    def create_tabs_panel(self, parent):
        """Create the tabs panel with data tables"""
        tabs_frame = tk.Frame(parent, bg='#34495e')
        tabs_frame.pack(side='right', fill='both', expand=True, pady=5)
        
        # Create notebook
        self.notebook = ttk.Notebook(tabs_frame)
        self.notebook.pack(fill='both', expand=True)
        
        # Active repairs tab
        self.active_frame = tk.Frame(self.notebook, bg='#ecf0f1')
        self.notebook.add(self.active_frame, text="Active Repairs")
        self.create_active_table()
        
        # Past repairs tab
        self.shipped_frame = tk.Frame(self.notebook, bg='#ecf0f1')
        self.notebook.add(self.shipped_frame, text="Past Repairs")
        self.create_shipped_table()
        
        # Customer view tab
        self.customer_frame = tk.Frame(self.notebook, bg='#ecf0f1')
        self.notebook.add(self.customer_frame, text="Customer View")
        self.create_customer_view()
    
    def create_active_table(self):
        """Create the active repairs table"""
        # Customer filter
        filter_frame = tk.Frame(self.active_frame, bg='#ecf0f1')
        filter_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(filter_frame, text="Filter by Customer:", bg='#ecf0f1',
                font=('Arial', 10, 'bold')).pack(side='left')
        
        self.active_customer_filter = ttk.Combobox(filter_frame, width=20)
        self.active_customer_filter.pack(side='left', padx=10)
        self.active_customer_filter.bind('<<ComboboxSelected>>', lambda e: self.refresh_active_table())
        
        # Status filter
        tk.Label(filter_frame, text="    Status Filter:", bg='#ecf0f1',
                font=('Arial', 10, 'bold')).pack(side='left')
        
        self.active_status_filter = ttk.Combobox(filter_frame, width=25, state='readonly')
        self.active_status_filter['values'] = [
            '',
            'Estimate Not Sent',
            'Estimate Not Complete',
            'Needs Estimate (Not Sent & Not Complete)',
            'Needs Machine Info',
            'Parts Not Ordered',
            'P.O. Not Received',
            'Repair Not Complete',
        ]
        self.active_status_filter.pack(side='left', padx=10)
        self.active_status_filter.bind('<<ComboboxSelected>>', lambda e: self.refresh_active_table())
        
        # Field / Shop filter
        tk.Label(filter_frame, text="    Field/Shop:", bg='#ecf0f1',
                font=('Arial', 10, 'bold')).pack(side='left')
        
        self.active_location_filter = ttk.Combobox(filter_frame, width=12, state='readonly')
        self.active_location_filter['values'] = ['', 'Field', 'Shop']
        self.active_location_filter.pack(side='left', padx=10)
        self.active_location_filter.bind('<<ComboboxSelected>>',
                                        lambda e: self.refresh_active_table())
        
        # Clear filters button
        tk.Button(filter_frame, text="Clear Filters", command=self.clear_active_filters,
                 bg='#95a5a6', fg='white', font=('Arial', 9, 'bold')).pack(side='left', padx=5)
        
        # Table frame for proper scrollbar layout
        table_frame = tk.Frame(self.active_frame)
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Table
        self.active_tree = ttk.Treeview(table_frame, columns=('estimate', 'customer', 'location', 'arrival', 
                                       'serial', 'make', 'model', 'est_sent', 'est_complete', 'parts_ord', 'po_recv', 'po_num', 'repair_done', 'pdf'), show='headings')
        
        # Column headings
        self.active_tree.heading('estimate', text='Estimate #')
        self.active_tree.heading('customer', text='Customer')
        self.active_tree.heading('location', text='Field/Shop',
                                command=lambda: self.sort_tree_by_column(self.active_tree, 'location'))
        self.active_tree.heading('arrival', text='Arrival')
        self.active_tree.heading('serial', text='Serial #')
        self.active_tree.heading('make', text='Make')
        self.active_tree.heading('model', text='Model')
        self.active_tree.heading('est_sent', text='Est. Sent')
        self.active_tree.heading('est_complete', text='Est. Done')
        self.active_tree.heading('parts_ord', text='Parts Ord.')
        self.active_tree.heading('po_recv', text='P.O. Recv\'d')
        self.active_tree.heading('po_num', text='P.O. Number')
        self.active_tree.heading('repair_done', text='Repair Done')
        self.active_tree.heading('pdf', text='PDF')
        
        # Column widths
        self.active_tree.column('estimate', width=100)
        self.active_tree.column('customer', width=120)
        self.active_tree.column('location', width=80, anchor='center')
        self.active_tree.column('arrival', width=80)
        self.active_tree.column('serial', width=100)
        self.active_tree.column('make', width=80)
        self.active_tree.column('model', width=80)
        self.active_tree.column('est_sent', width=65)
        self.active_tree.column('est_complete', width=65)
        self.active_tree.column('parts_ord', width=65)
        self.active_tree.column('po_recv', width=65)
        self.active_tree.column('po_num', width=90)
        self.active_tree.column('repair_done', width=75)
        self.active_tree.column('pdf', width=40)
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.active_tree.yview)
        h_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.active_tree.xview)
        self.active_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # Pack table and scrollbars with proper grid layout
        self.active_tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        
        # Configure grid weights
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Entries still waiting on serial / make / model
        self.active_tree.tag_configure('needsinfo', background='#fdebd0')
        
        # Context menu
        self.active_tree.bind("<Button-3>", self.show_active_context_menu)
        self.active_tree.bind("<Double-1>", self.on_active_double_click)
    
    def create_shipped_table(self):
        """Create the shipped repairs table"""
        # Customer filter
        filter_frame = tk.Frame(self.shipped_frame, bg='#ecf0f1')
        filter_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(filter_frame, text="Filter by Customer:", bg='#ecf0f1',
                font=('Arial', 10, 'bold')).pack(side='left')
        
        self.shipped_customer_filter = ttk.Combobox(filter_frame, width=20)
        self.shipped_customer_filter.pack(side='left', padx=10)
        self.shipped_customer_filter.bind('<<ComboboxSelected>>', lambda e: self.refresh_shipped_table())
        
        # Field / Shop filter
        tk.Label(filter_frame, text="    Field/Shop:", bg='#ecf0f1',
                font=('Arial', 10, 'bold')).pack(side='left')
        
        self.shipped_location_filter = ttk.Combobox(filter_frame, width=12, state='readonly')
        self.shipped_location_filter['values'] = ['', 'Field', 'Shop']
        self.shipped_location_filter.pack(side='left', padx=10)
        self.shipped_location_filter.bind('<<ComboboxSelected>>',
                                         lambda e: self.refresh_shipped_table())
        
        tk.Button(filter_frame, text="Clear Filters", command=self.clear_shipped_filters,
                 bg='#95a5a6', fg='white', font=('Arial', 9, 'bold')).pack(side='left', padx=5)
        
        # Table frame for proper scrollbar layout
        table_frame2 = tk.Frame(self.shipped_frame)
        table_frame2.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Table
        self.shipped_tree = ttk.Treeview(table_frame2, columns=('estimate', 'customer', 'location', 'arrival',
                                        'ship_date', 'serial', 'make', 'model', 'po_num', 'pdf'), show='headings')
        
        # Column headings
        self.shipped_tree.heading('estimate', text='Estimate #')
        self.shipped_tree.heading('customer', text='Customer')
        self.shipped_tree.heading('location', text='Field/Shop',
                                 command=lambda: self.sort_tree_by_column(self.shipped_tree, 'location'))
        self.shipped_tree.heading('arrival', text='Arrival')
        self.shipped_tree.heading('ship_date', text='Ship Date')
        self.shipped_tree.heading('serial', text='Serial #')
        self.shipped_tree.heading('make', text='Make')
        self.shipped_tree.heading('model', text='Model')
        self.shipped_tree.heading('po_num', text='P.O. Number')
        self.shipped_tree.heading('pdf', text='PDF')
        
        # Column widths
        for col in ('estimate', 'customer', 'arrival', 'ship_date', 'serial', 'make', 'model', 'po_num'):
            self.shipped_tree.column(col, width=100)
        self.shipped_tree.column('location', width=80, anchor='center')
        self.shipped_tree.column('pdf', width=50)
        
        # Scrollbars
        v_scroll2 = ttk.Scrollbar(table_frame2, orient="vertical", command=self.shipped_tree.yview)
        h_scroll2 = ttk.Scrollbar(table_frame2, orient="horizontal", command=self.shipped_tree.xview)
        self.shipped_tree.configure(yscrollcommand=v_scroll2.set, xscrollcommand=h_scroll2.set)
        
        # Pack table and scrollbars with proper grid layout
        self.shipped_tree.grid(row=0, column=0, sticky='nsew')
        v_scroll2.grid(row=0, column=1, sticky='ns')
        h_scroll2.grid(row=1, column=0, sticky='ew')
        
        # Configure grid weights
        table_frame2.grid_rowconfigure(0, weight=1)
        table_frame2.grid_columnconfigure(0, weight=1)
        
        # Context menu for shipped items
        self.shipped_tree.tag_configure('needsinfo', background='#fdebd0')
        self.shipped_tree.bind("<Button-3>", self.show_shipped_context_menu)
        self.shipped_tree.bind("<Double-1>", self.open_pdf_from_shipped)
    
    def create_customer_view(self):
        """Create the customer overview"""
        # Customer selection
        select_frame = tk.Frame(self.customer_frame, bg='#ecf0f1')
        select_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(select_frame, text="Select Customer:", bg='#ecf0f1',
                font=('Arial', 12, 'bold')).pack(side='left')
        
        self.customer_select = ttk.Combobox(select_frame, width=25)
        self.customer_select.pack(side='left', padx=10)
        self.customer_select.bind('<<ComboboxSelected>>', self.show_customer_details)
        
        # Details frame for proper scrollbar layout
        details_frame = tk.Frame(self.customer_frame)
        details_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Details area
        self.customer_details = tk.Text(details_frame, wrap='word', bg='white')
        customer_scroll = ttk.Scrollbar(details_frame, orient="vertical", command=self.customer_details.yview)
        self.customer_details.configure(yscrollcommand=customer_scroll.set)
        
        # Pack with grid
        self.customer_details.grid(row=0, column=0, sticky='nsew')
        customer_scroll.grid(row=0, column=1, sticky='ns')
        
        # Configure grid weights
        details_frame.grid_rowconfigure(0, weight=1)
        details_frame.grid_columnconfigure(0, weight=1)
    
    def generate_estimate_number(self, date_str):
        """Generate estimate number in format MMDDYY-XXX"""
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            date_part = date_obj.strftime('%m%d%y')
            
            # Collect every number already used - in memory, in the local
            # backup, and on the share. Counting only what is loaded is how
            # the same number got issued twice after an offline sync.
            taken = set()
            for m in self.active_machines + self.shipped_machines:
                taken.add(str(m.get('estimate_number', '')))
            for path in (self.get_local_backup_path(), self.save_location):
                try:
                    blob = self._read_json(path)
                    for keyname in ('active_machines', 'shipped_machines'):
                        for m in blob.get(keyname, []) or []:
                            taken.add(str(m.get('estimate_number', '')))
                except Exception:
                    pass

            n = 1
            while n <= 999:
                candidate = f"{date_part}-{str(n).zfill(3)}"
                if candidate not in taken:
                    return candidate
                n += 1
            return f"{date_part}-{uuid.uuid4().hex[:3]}"
        except:
            return "Invalid-Date"
    
    def update_estimate_display(self, *args):
        """Update the estimate number display"""
        if self.estimate_label:
            try:
                estimate = self.generate_estimate_number(self.arrival_var.get())
                self.estimate_label.config(text=f"Next Estimate #: {estimate}")
            except:
                self.estimate_label.config(text="Next Estimate #: Invalid Date")
    
    def add_machine(self):
        """Add a new machine.

        Only the customer and the arrival date are required. Serial, make and
        model can be filled in later (right-click the entry -> Edit Machine
        Details), so a job can be logged the moment it is booked in, before
        anyone has laid eyes on the machine.
        """
        # Validate the two fields we genuinely need
        if not self.customer_var.get().strip():
            messagebox.showerror("Error", "Please enter a customer name.")
            return
        if not self.arrival_var.get().strip():
            messagebox.showerror("Error", "Please enter a date of arrival.")
            return
        
        try:
            # Validate date
            datetime.strptime(self.arrival_var.get(), '%Y-%m-%d')
        except:
            messagebox.showerror("Error", "Please enter date in YYYY-MM-DD format")
            return
        
        # Let the user know what is still blank, so it is a conscious choice
        missing = []
        if not self.serial_var.get().strip():
            missing.append("Serial Number")
        if not self.make_var.get().strip():
            missing.append("Make")
        if not self.model_var.get().strip():
            missing.append("Model")
        if missing:
            if not messagebox.askyesno(
                    "Add Without Machine Details?",
                    f"{', '.join(missing)} not filled in.\n\n"
                    "The entry will be added and flagged as needing machine info. "
                    "You can right-click it in the log later and choose "
                    "'Edit Machine Details' to finish it.\n\n"
                    "Add it now?"):
                return
        
        # Create machine record
        now_iso = datetime.now().isoformat()
        machine = {
            'record_id': new_record_id(),
            'created_at': now_iso,
            'updated_at': now_iso,
            'estimate_number': self.generate_estimate_number(self.arrival_var.get()),
            'customer_name': self.customer_var.get().strip(),
            'arrival_date': self.arrival_var.get(),
            'serial_number': self.serial_var.get().strip(),
            'make': self.make_var.get().strip(),
            'model': self.model_var.get().strip(),
            'service_location': self.location_var.get(),
            'po_received': False,
            'po_number': '',
            'pdf_attachment': None,
            'estimate_sent': False,
            'estimate_complete': False,
            'parts_ordered': False,
            'repair_complete': False
        }
        
        # Handle PDF attachment
        if self.current_pdf_path:
            try:
                # Create customer directory if it doesn't exist
                customer_dir = os.path.join(self.attachments_dir, machine['customer_name'])
                os.makedirs(customer_dir, exist_ok=True)
                
                # Copy PDF to attachments directory
                pdf_filename = f"{machine['estimate_number']}_{os.path.basename(self.current_pdf_path)}"
                pdf_dest = os.path.join(customer_dir, pdf_filename)
                shutil.copy2(self.current_pdf_path, pdf_dest)
                machine['pdf_attachment'] = pdf_dest
                
            except Exception as e:
                messagebox.showwarning("Warning", f"Could not copy PDF file: {str(e)}")
        
        # Add to active machines
        self.active_machines.append(machine)
        
        # Clear form (except customer and date)
        self.serial_var.set('')
        self.make_var.set('')
        self.model_var.set('')
        self.pdf_path_var.set("No PDF attached")
        self.current_pdf_path = None
        
        # Refresh display
        self.update_estimate_display()
        self.refresh_tables()
        
        if missing:
            messagebox.showinfo(
                "Added - Needs Machine Info",
                f"Entry {machine['estimate_number']} added for "
                f"{machine['customer_name']}.\n\n"
                f"Still needed: {', '.join(missing)}\n\n"
                "It is highlighted in the log. Filter by 'Needs Machine Info' "
                "to find everything waiting on details.")
        else:
            messagebox.showinfo("Success", f"Machine {machine['estimate_number']} added successfully!")
    
    @staticmethod
    def missing_machine_details(machine):
        """Which of Serial / Make / Model are still blank on this record."""
        missing = []
        if not str(machine.get('serial_number', '') or '').strip():
            missing.append('Serial Number')
        if not str(machine.get('make', '') or '').strip():
            missing.append('Make')
        if not str(machine.get('model', '') or '').strip():
            missing.append('Model')
        return missing

    def needs_machine_info(self, machine):
        """True when an entry was logged before the machine details were known."""
        return bool(self.missing_machine_details(machine))

    def edit_machine_details(self, machine_idx, shipped=False):
        """Fill in Serial / Make / Model on an entry added earlier."""
        records = self.shipped_machines if shipped else self.active_machines
        if machine_idx >= len(records):
            messagebox.showerror("Error", "Entry not found.")
            return
        machine = records[machine_idx]

        dlg = tk.Toplevel(self.root)
        dlg.title("Edit Machine Details")
        dlg.configure(bg='#2c3e50')
        dlg.resizable(False, False)
        dlg.transient(self.root)

        tk.Label(dlg, text="Machine Details", font=('Arial', 14, 'bold'),
                 fg='#3498db', bg='#2c3e50').pack(anchor='w', padx=20, pady=(15, 2))
        tk.Label(dlg,
                 text=f"{machine.get('estimate_number', '')}  -  "
                      f"{machine.get('customer_name', '')}",
                 font=('Arial', 10), fg='#bdc3c7',
                 bg='#2c3e50').pack(anchor='w', padx=20, pady=(0, 12))

        grid = tk.Frame(dlg, bg='#2c3e50')
        grid.pack(fill='x', padx=20)

        serial_v = tk.StringVar(master=dlg, value=machine.get('serial_number', ''))
        make_v = tk.StringVar(master=dlg, value=machine.get('make', ''))
        model_v = tk.StringVar(master=dlg, value=machine.get('model', ''))

        entries = []
        for row, (label, var) in enumerate((("Serial Number:", serial_v),
                                            ("Make:", make_v),
                                            ("Model:", model_v))):
            tk.Label(grid, text=label, font=('Arial', 10, 'bold'), fg='white',
                     bg='#2c3e50').grid(row=row, column=0, sticky='e', pady=6, padx=(0, 8))
            ent = tk.Entry(grid, textvariable=var, width=30, font=('Arial', 11))
            ent.grid(row=row, column=1, pady=6)
            entries.append(ent)

        tk.Label(dlg, text="Leave anything you still don't know blank.",
                 font=('Arial', 8, 'italic'), fg='#95a5a6',
                 bg='#2c3e50').pack(anchor='w', padx=20, pady=(8, 0))

        btns = tk.Frame(dlg, bg='#2c3e50')
        btns.pack(fill='x', padx=20, pady=(12, 16))

        def save():
            machine['serial_number'] = serial_v.get().strip()
            machine['make'] = make_v.get().strip()
            machine['model'] = model_v.get().strip()
            self.save_data()
            self.refresh_tables()
            still = self.missing_machine_details(machine)
            if self.status_label:
                if still:
                    self.status_label.config(
                        text=f"{machine.get('estimate_number', '')} updated - "
                             f"still needs: {', '.join(still)}", fg='#e67e22')
                else:
                    self.status_label.config(
                        text=f"{machine.get('estimate_number', '')} machine details complete",
                        fg='#27ae60')
            dlg.destroy()

        tk.Button(btns, text="Save", command=save, bg='#27ae60', fg='white',
                  font=('Arial', 10, 'bold'), width=12).pack(side='right', padx=5)
        tk.Button(btns, text="Cancel", command=dlg.destroy, bg='#7f8c8d', fg='white',
                  font=('Arial', 10, 'bold'), width=12).pack(side='right', padx=5)

        dlg.bind('<Return>', lambda e: save())
        dlg.update_idletasks()
        try:
            px = self.root.winfo_rootx() + 250
            py = self.root.winfo_rooty() + 200
            dlg.geometry(f"+{px}+{py}")
        except Exception:
            pass
        if entries:
            entries[0].focus_set()
        dlg.grab_set()

    def set_service_location(self, machine_idx, location, shipped=False):
        """Switch a record between Field and Shop."""
        try:
            records = self.shipped_machines if shipped else self.active_machines
            records[machine_idx]['service_location'] = location
            self.save_data()
            self.refresh_tables()
            if self.status_label:
                est = records[machine_idx].get('estimate_number', '')
                self.status_label.config(text=f"{est} set to {location}", fg='#27ae60')
        except Exception as e:
            messagebox.showerror("Error", f"Could not change service location: {e}")

    def refresh_tables(self):
        """Refresh all tables"""
        self.refresh_active_table()
        self.refresh_shipped_table()
        self.update_customer_filters()
    
    def update_customer_filters(self):
        """Update customer filter dropdowns"""
        all_customers = set()
        for machine in self.active_machines + self.shipped_machines:
            all_customers.add(machine['customer_name'])
        
        customers = [''] + sorted(list(all_customers))
        
        if self.active_customer_filter:
            self.active_customer_filter['values'] = customers
        if self.shipped_customer_filter:
            self.shipped_customer_filter['values'] = customers
        if self.customer_select:
            self.customer_select['values'] = sorted(list(all_customers))
    
    def refresh_active_table(self):
        """Refresh the active repairs table"""
        if not self.active_tree:
            return
            
        # Clear existing items
        for item in self.active_tree.get_children():
            self.active_tree.delete(item)
        
        # Filter machines by customer
        filter_customer = self.active_customer_filter.get() if self.active_customer_filter else ''
        machines = self.active_machines
        if filter_customer:
            machines = [m for m in machines if m['customer_name'] == filter_customer]
        
        # Filter machines by status
        status_filter = self.active_status_filter.get() if self.active_status_filter else ''
        if status_filter == 'Estimate Not Sent':
            machines = [m for m in machines if not m.get('estimate_sent', False)]
        elif status_filter == 'Estimate Not Complete':
            machines = [m for m in machines if not m.get('estimate_complete', False)]
        elif status_filter == 'P.O. Not Received':
            machines = [m for m in machines if not m.get('po_received', False)]
        elif status_filter == 'Needs Estimate (Not Sent & Not Complete)':
            machines = [m for m in machines if not m.get('estimate_sent', False) and not m.get('estimate_complete', False)]
        elif status_filter == 'Parts Not Ordered':
            machines = [m for m in machines if not m.get('parts_ordered', False)]
        elif status_filter == 'Repair Not Complete':
            machines = [m for m in machines if not m.get('repair_complete', False)]
        elif status_filter == 'Needs Machine Info':
            machines = [m for m in machines if self.needs_machine_info(m)]
        
        # Filter by Field / Shop
        loc_filter = self.active_location_filter.get() if self.active_location_filter else ''
        if loc_filter in ('Field', 'Shop'):
            machines = [m for m in machines
                        if m.get('service_location', 'Shop') == loc_filter]
        
        # Add machines to table
        for machine in machines:
            row_tags = ('needsinfo',) if self.needs_machine_info(machine) else ()
            self.active_tree.insert('', 'end', tags=row_tags, values=(
                machine['estimate_number'],
                machine['customer_name'],
                machine.get('service_location', 'Shop'),
                machine['arrival_date'],
                machine['serial_number'],
                machine['make'],
                machine['model'],
                'Yes' if machine.get('estimate_sent', False) else 'No',
                'Yes' if machine.get('estimate_complete', False) else 'No',
                'Yes' if machine.get('parts_ordered', False) else 'No',
                'Yes' if machine['po_received'] else 'No',
                machine.get('po_number', ''),
                'Yes' if machine.get('repair_complete', False) else 'No',
                '📎' if machine.get('pdf_attachment') else ''
            ))
    
    def clear_active_filters(self):
        """Clear all active table filters"""
        if self.active_customer_filter:
            self.active_customer_filter.set('')
        if self.active_status_filter:
            self.active_status_filter.set('')
        if self.active_location_filter:
            self.active_location_filter.set('')
        self.refresh_active_table()
    
    def clear_shipped_filters(self):
        """Clear all past-repair table filters"""
        if self.shipped_customer_filter:
            self.shipped_customer_filter.set('')
        if self.shipped_location_filter:
            self.shipped_location_filter.set('')
        self.refresh_shipped_table()

    def sort_tree_by_column(self, tree, col):
        """Click a column heading to sort by it; click again to reverse."""
        try:
            if not hasattr(self, '_tree_sort_state'):
                self._tree_sort_state = {}
            key = (str(tree), col)
            reverse = self._tree_sort_state.get(key, False)

            rows = [(tree.set(item, col), item) for item in tree.get_children('')]
            rows.sort(key=lambda r: str(r[0]).lower(), reverse=reverse)
            for index, (_val, item) in enumerate(rows):
                tree.move(item, '', index)

            self._tree_sort_state[key] = not reverse
        except Exception as e:
            print(f"Sort failed: {e}")

    def refresh_shipped_table(self):
        """Refresh the shipped repairs table"""
        if not self.shipped_tree:
            return
            
        # Clear existing items
        for item in self.shipped_tree.get_children():
            self.shipped_tree.delete(item)
        
        # Filter machines
        filter_customer = self.shipped_customer_filter.get() if self.shipped_customer_filter else ''
        machines = self.shipped_machines
        if filter_customer:
            machines = [m for m in machines if m['customer_name'] == filter_customer]
        
        # Filter by Field / Shop
        loc_filter = self.shipped_location_filter.get() if self.shipped_location_filter else ''
        if loc_filter in ('Field', 'Shop'):
            machines = [m for m in machines
                        if m.get('service_location', 'Shop') == loc_filter]
        
        # Add machines to table
        for machine in machines:
            row_tags = ('needsinfo',) if self.needs_machine_info(machine) else ()
            self.shipped_tree.insert('', 'end', tags=row_tags, values=(
                machine['estimate_number'],
                machine['customer_name'],
                machine.get('service_location', 'Shop'),
                machine['arrival_date'],
                machine.get('ship_date', ''),
                machine['serial_number'],
                machine['make'],
                machine['model'],
                machine.get('po_number', ''),
                '📎' if machine.get('pdf_attachment') else ''
            ))
    

    def on_active_double_click(self, event):
        """Handle double-click on active repair to open service order"""
        if not self.active_tree:
            return
        row_id = self.active_tree.identify_row(event.y)
        if not row_id:
            return
        self.active_tree.selection_set(row_id)
        item = self.active_tree.selection()
        if not item:
            return
        values = self.active_tree.item(item[0])['values']
        estimate_num = str(values[0])
        
        # Find the machine
        for machine in self.active_machines:
            if machine['estimate_number'] == estimate_num:
                self.open_service_order(machine)
                return
    
    def open_service_order(self, machine):
        """Open a service order window for the given machine"""
        # Check if a window is already open for this machine
        for win_ref in self._service_order_windows[:]:
            try:
                if win_ref.linked_estimate_number == machine['estimate_number']:
                    win_ref.window.lift()
                    win_ref.window.focus_force()
                    return
            except:
                self._service_order_windows.remove(win_ref)
        
        # Create new service order window
        so_window = ServiceOrderWindow(self, machine)
        self._service_order_windows.append(so_window)

    def show_active_context_menu(self, event):
        """Show context menu for active repairs"""
        if not self.active_tree:
            return
        
        # Select the row under the cursor (right-click doesn't auto-select)
        row_id = self.active_tree.identify_row(event.y)
        if not row_id:
            return
        self.active_tree.selection_set(row_id)
        self.active_tree.focus(row_id)
        
        item = self.active_tree.selection()
        if not item:
            return
        
        # Get machine data
        values = self.active_tree.item(item[0])['values']
        estimate_num = str(values[0])
        
        # Find machine in list
        machine_idx = None
        for i, machine in enumerate(self.active_machines):
            if machine['estimate_number'] == estimate_num:
                machine_idx = i
                break
        
        if machine_idx is None:
            return
        
        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0)
        
        # Add Print Barcode Label option at the top
        context_menu.add_command(label="🖨️ Print Barcode Label", 
                                command=lambda: self.print_barcode_label(machine_idx))
        context_menu.add_command(label="📋 Open Service Order",
                                command=lambda: self.open_service_order(self.active_machines[machine_idx]))
        context_menu.add_separator()
        
        # Edit options
        _missing = self.missing_machine_details(self.active_machines[machine_idx])
        _label = "🔧 Edit Machine Details"
        if _missing:
            _label += f"  (needs {len(_missing)})"
        context_menu.add_command(label=_label,
                                command=lambda: self.edit_machine_details(machine_idx))
        context_menu.add_command(label="Edit Estimate Number", command=lambda: self.edit_estimate_number(machine_idx))
        
        # Field / Shop toggle
        cur_loc = self.active_machines[machine_idx].get('service_location', 'Shop')
        other_loc = 'Field' if cur_loc == 'Shop' else 'Shop'
        context_menu.add_command(
            label=f"Change to {other_loc}  (now: {cur_loc})",
            command=lambda: self.set_service_location(machine_idx, other_loc))
        
        # Estimate sent toggle
        est_sent_text = "Remove Estimate Sent" if self.active_machines[machine_idx].get('estimate_sent', False) else "Mark Estimate Sent"
        context_menu.add_command(label=est_sent_text, command=lambda: self.toggle_estimate_sent(machine_idx))
        
        est_complete_text = "Remove Estimate Complete" if self.active_machines[machine_idx].get('estimate_complete', False) else "Mark Estimate Complete"
        context_menu.add_command(label=est_complete_text, command=lambda: self.toggle_estimate_complete(machine_idx))
        
        context_menu.add_separator()
        
        # Parts ordered toggle
        parts_ord_text = "Remove Parts Ordered" if self.active_machines[machine_idx].get('parts_ordered', False) else "Mark Parts Ordered"
        context_menu.add_command(label=parts_ord_text, command=lambda: self.toggle_parts_ordered(machine_idx))
        
        # PO options
        po_text = "Remove P.O. Received" if self.active_machines[machine_idx]['po_received'] else "Mark P.O. Received"
        context_menu.add_command(label=po_text, command=lambda: self.toggle_po_received(machine_idx))
        
        if self.active_machines[machine_idx]['po_received']:
            context_menu.add_command(label="Edit P.O. Number", command=lambda: self.edit_po_number(machine_idx))
        
        context_menu.add_separator()
        
        # Repair complete toggle
        repair_text = "Remove Repair Complete" if self.active_machines[machine_idx].get('repair_complete', False) else "Mark Repair Complete"
        context_menu.add_command(label=repair_text, command=lambda: self.toggle_repair_complete(machine_idx))
        
        context_menu.add_separator()
        context_menu.add_command(label="Ship Machine", command=lambda: self.ship_machine(machine_idx))
        
        # Add PDF options
        if self.active_machines[machine_idx].get('pdf_attachment'):
            context_menu.add_separator()
            context_menu.add_command(label="View PDF", command=lambda: self.open_pdf(machine_idx))
            context_menu.add_command(label="Replace PDF", command=lambda: self.replace_pdf(machine_idx))
            context_menu.add_command(label="Remove PDF", command=lambda: self.remove_pdf(machine_idx))
        else:
            context_menu.add_separator()
            context_menu.add_command(label="Add PDF", command=lambda: self.add_pdf_to_machine(machine_idx))
        
        # Delete option at bottom
        context_menu.add_separator()
        context_menu.add_command(label="❌ Delete Entry", command=lambda: self.delete_active_entry(machine_idx),
                                foreground='red')
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def show_shipped_context_menu(self, event):
        """Show context menu for shipped repairs"""
        if not self.shipped_tree:
            return
        
        # Select the row under the cursor (right-click doesn't auto-select)
        row_id = self.shipped_tree.identify_row(event.y)
        if not row_id:
            return
        self.shipped_tree.selection_set(row_id)
        self.shipped_tree.focus(row_id)
        
        item = self.shipped_tree.selection()
        if not item:
            return
        
        # Get machine data
        values = self.shipped_tree.item(item[0])['values']
        estimate_num = str(values[0])
        
        # Find machine in list
        machine_idx = None
        for i, machine in enumerate(self.shipped_machines):
            if machine['estimate_number'] == estimate_num:
                machine_idx = i
                break
        
        if machine_idx is None:
            return
        
        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0)
        
        context_menu.add_command(label="Revert to Active", command=lambda: self.revert_to_active(machine_idx))
        
        _missing = self.missing_machine_details(self.shipped_machines[machine_idx])
        _label = "🔧 Edit Machine Details"
        if _missing:
            _label += f"  (needs {len(_missing)})"
        context_menu.add_command(
            label=_label,
            command=lambda: self.edit_machine_details(machine_idx, shipped=True))
        
        # Field / Shop toggle
        cur_loc = self.shipped_machines[machine_idx].get('service_location', 'Shop')
        other_loc = 'Field' if cur_loc == 'Shop' else 'Shop'
        context_menu.add_command(
            label=f"Change to {other_loc}  (now: {cur_loc})",
            command=lambda: self.set_service_location(machine_idx, other_loc, shipped=True))
        
        # Add PDF options
        if self.shipped_machines[machine_idx].get('pdf_attachment'):
            context_menu.add_separator()
            context_menu.add_command(label="View PDF", command=lambda: self.open_pdf(machine_idx, shipped=True))
            context_menu.add_command(label="Replace PDF", command=lambda: self.replace_pdf(machine_idx, shipped=True))
            context_menu.add_command(label="Remove PDF", command=lambda: self.remove_pdf(machine_idx, shipped=True))
        else:
            context_menu.add_separator()
            context_menu.add_command(label="Add PDF", command=lambda: self.add_pdf_to_machine(machine_idx, shipped=True))
        
        # Delete option at bottom
        context_menu.add_separator()
        context_menu.add_command(label="❌ Delete Entry", command=lambda: self.delete_shipped_entry(machine_idx),
                                foreground='red')
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def revert_to_active(self, machine_idx):
        """Revert a shipped machine back to active status"""
        if machine_idx < len(self.shipped_machines):
            machine = self.shipped_machines[machine_idx]
            if messagebox.askyesno("Revert Machine", f"Revert machine {machine['estimate_number']} back to active status?"):
                # Remove ship date and move back to active
                if 'ship_date' in machine:
                    del machine['ship_date']
                self.active_machines.append(machine)
                del self.shipped_machines[machine_idx]
                self.refresh_tables()
                messagebox.showinfo("Success", f"Machine {machine['estimate_number']} reverted to active status!")
    
    def open_pdf(self, machine_idx, shipped=False):
        """Open the PDF attachment for a machine"""
        machine_list = self.shipped_machines if shipped else self.active_machines
        if machine_idx < len(machine_list):
            machine = machine_list[machine_idx]
            pdf_path = machine.get('pdf_attachment')
            if pdf_path and os.path.exists(pdf_path):
                try:
                    os.startfile(pdf_path)  # Works on Windows
                except:
                    try:
                        import subprocess
                        subprocess.run(['open', pdf_path])  # Works on macOS
                    except:
                        try:
                            subprocess.run(['xdg-open', pdf_path])  # Works on Linux
                        except:
                            messagebox.showerror("Error", "Could not open PDF file")
            else:
                messagebox.showwarning("Warning", "No PDF attachment found or file doesn't exist")
    
    def open_pdf_from_active(self, event):
        """Open PDF when double-clicking on active table"""
        item = self.active_tree.selection()
        if item:
            values = self.active_tree.item(item[0])['values']
            if values and values[-1] == '📎':  # Check if PDF icon is present
                estimate_num = values[0]
                for i, machine in enumerate(self.active_machines):
                    if machine['estimate_number'] == estimate_num:
                        self.open_pdf(i)
                        break
    
    def open_pdf_from_shipped(self, event):
        """Open PDF when double-clicking on shipped table"""
        item = self.shipped_tree.selection()
        if item:
            values = self.shipped_tree.item(item[0])['values']
            if values and values[-1] == '📎':  # Check if PDF icon is present
                estimate_num = values[0]
                for i, machine in enumerate(self.shipped_machines):
                    if machine['estimate_number'] == estimate_num:
                        self.open_pdf(i, shipped=True)
                        break
    
    def replace_pdf(self, machine_idx, shipped=False):
        """Replace the PDF attachment for a machine"""
        machine_list = self.shipped_machines if shipped else self.active_machines
        if machine_idx < len(machine_list):
            machine = machine_list[machine_idx]
            file_path = filedialog.askopenfilename(
                title="Select PDF File",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            if file_path:
                try:
                    # Create customer directory if it doesn't exist
                    customer_dir = os.path.join(self.attachments_dir, machine['customer_name'])
                    os.makedirs(customer_dir, exist_ok=True)
                    
                    # Copy new PDF
                    pdf_filename = f"{machine['estimate_number']}_{os.path.basename(file_path)}"
                    pdf_dest = os.path.join(customer_dir, pdf_filename)
                    shutil.copy2(file_path, pdf_dest)
                    
                    # Remove old PDF if it exists
                    old_pdf = machine.get('pdf_attachment')
                    if old_pdf and os.path.exists(old_pdf):
                        os.remove(old_pdf)
                    
                    machine['pdf_attachment'] = pdf_dest
                    self.refresh_tables()
                    messagebox.showinfo("Success", "PDF attachment replaced successfully!")
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Could not replace PDF: {str(e)}")
    
    def remove_pdf(self, machine_idx, shipped=False):
        """Remove the PDF attachment from a machine"""
        machine_list = self.shipped_machines if shipped else self.active_machines
        if machine_idx < len(machine_list):
            machine = machine_list[machine_idx]
            pdf_path = machine.get('pdf_attachment')
            if pdf_path:
                if messagebox.askyesno("Remove PDF", "Are you sure you want to remove the PDF attachment?"):
                    try:
                        if os.path.exists(pdf_path):
                            os.remove(pdf_path)
                        machine['pdf_attachment'] = None
                        self.refresh_tables()
                        messagebox.showinfo("Success", "PDF attachment removed successfully!")
                    except Exception as e:
                        messagebox.showerror("Error", f"Could not remove PDF: {str(e)}")
            else:
                messagebox.showwarning("Warning", "No PDF attachment to remove")
    
    def add_pdf_to_machine(self, machine_idx, shipped=False):
        """Add a PDF attachment to an existing machine"""
        machine_list = self.shipped_machines if shipped else self.active_machines
        if machine_idx < len(machine_list):
            machine = machine_list[machine_idx]
            file_path = filedialog.askopenfilename(
                title="Select PDF File",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            if file_path:
                try:
                    # Create customer directory if it doesn't exist
                    customer_dir = os.path.join(self.attachments_dir, machine['customer_name'])
                    os.makedirs(customer_dir, exist_ok=True)
                    
                    # Copy PDF to attachments directory
                    pdf_filename = f"{machine['estimate_number']}_{os.path.basename(file_path)}"
                    pdf_dest = os.path.join(customer_dir, pdf_filename)
                    shutil.copy2(file_path, pdf_dest)
                    
                    machine['pdf_attachment'] = pdf_dest
                    self.refresh_tables()
                    messagebox.showinfo("Success", "PDF attachment added successfully!")
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Could not add PDF: {str(e)}")
    
    def toggle_po_received(self, machine_idx):
        """Toggle P.O. received status"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            machine['po_received'] = not machine['po_received']
            if not machine['po_received']:
                machine['po_number'] = ''
            self.refresh_active_table()
    
    def edit_po_number(self, machine_idx):
        """Edit P.O. number"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            po_number = simpledialog.askstring("P.O. Number", 
                                             f"Enter P.O. Number for {machine['estimate_number']}:",
                                             initialvalue=machine.get('po_number', ''))
            if po_number is not None:
                machine['po_number'] = po_number
                self.refresh_active_table()
    
    def ship_machine(self, machine_idx):
        """Ship a machine"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            if messagebox.askyesno("Ship Machine", f"Ship machine {machine['estimate_number']}?"):
                machine['ship_date'] = datetime.now().strftime('%Y-%m-%d')
                self.shipped_machines.append(machine)
                del self.active_machines[machine_idx]
                self.refresh_tables()
                messagebox.showinfo("Success", f"Machine {machine['estimate_number']} marked as shipped!")
    
    def edit_estimate_number(self, machine_idx):
        """Edit estimate number for a machine"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            old_estimate = machine['estimate_number']
            
            new_estimate = simpledialog.askstring("Edit Estimate Number", 
                                                 f"Edit estimate number for {old_estimate}:\n\n"
                                                 "Format: MMDDYY-XXX (e.g., 020625-001)",
                                                 initialvalue=old_estimate)
            
            if new_estimate is not None and new_estimate.strip():
                new_estimate = new_estimate.strip()
                
                # Check if estimate number already exists
                all_machines = self.active_machines + self.shipped_machines
                existing = [m for m in all_machines if m['estimate_number'] == new_estimate and m != machine]
                
                if existing:
                    messagebox.showerror("Error", f"Estimate number '{new_estimate}' already exists!")
                    return
                
                # Update estimate number
                machine['estimate_number'] = new_estimate
                
                # Update PDF filename if exists
                if machine.get('pdf_attachment'):
                    old_pdf = machine['pdf_attachment']
                    if os.path.exists(old_pdf):
                        pdf_dir = os.path.dirname(old_pdf)
                        pdf_basename = os.path.basename(old_pdf)
                        # Replace old estimate number in filename with new one
                        new_pdf_basename = pdf_basename.replace(old_estimate, new_estimate, 1)
                        new_pdf = os.path.join(pdf_dir, new_pdf_basename)
                        try:
                            os.rename(old_pdf, new_pdf)
                            machine['pdf_attachment'] = new_pdf
                        except:
                            pass  # Keep old path if rename fails
                
                self.refresh_active_table()
                messagebox.showinfo("Success", f"Estimate number updated from {old_estimate} to {new_estimate}")
    
    def toggle_estimate_sent(self, machine_idx):
        """Toggle estimate sent status"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            machine['estimate_sent'] = not machine.get('estimate_sent', False)
            self.refresh_active_table()
            
            status = "marked as sent" if machine['estimate_sent'] else "marked as not sent"
            if self.status_label:
                self.status_label.config(text=f"Estimate {machine['estimate_number']} {status}")
    
    def toggle_estimate_complete(self, machine_idx):
        """Toggle estimate complete status"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            machine['estimate_complete'] = not machine.get('estimate_complete', False)
            self.refresh_active_table()
            
            status = "marked as complete" if machine['estimate_complete'] else "marked as not complete"
            if self.status_label:
                self.status_label.config(text=f"Estimate {machine['estimate_number']} {status}")
    
    def toggle_parts_ordered(self, machine_idx):
        """Toggle parts ordered status"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            machine['parts_ordered'] = not machine.get('parts_ordered', False)
            self.refresh_active_table()
            
            status = "parts marked as ordered" if machine['parts_ordered'] else "parts marked as not ordered"
            if self.status_label:
                self.status_label.config(text=f"{machine['estimate_number']} {status}")
    
    def toggle_repair_complete(self, machine_idx):
        """Toggle repair complete status"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            machine['repair_complete'] = not machine.get('repair_complete', False)
            self.refresh_active_table()
            
            status = "repair marked as complete" if machine['repair_complete'] else "repair marked as not complete"
            if self.status_label:
                self.status_label.config(text=f"{machine['estimate_number']} {status}")
    
    def delete_active_entry(self, machine_idx):
        """Delete an active repair entry"""
        if machine_idx < len(self.active_machines):
            machine = self.active_machines[machine_idx]
            
            # Confirm deletion
            result = messagebox.askyesno("Delete Entry", 
                                        f"Are you sure you want to delete this entry?\n\n"
                                        f"Estimate: {machine['estimate_number']}\n"
                                        f"Customer: {machine['customer_name']}\n"
                                        f"Make/Model: {machine['make']} {machine['model']}\n\n"
                                        "This action cannot be undone!",
                                        icon='warning')
            
            if result:
                # Delete PDF attachment if exists
                if machine.get('pdf_attachment'):
                    pdf_path = machine['pdf_attachment']
                    if os.path.exists(pdf_path):
                        try:
                            os.remove(pdf_path)
                        except:
                            pass  # Continue even if PDF deletion fails
                
                # Remove from list
                estimate_num = machine['estimate_number']
                del self.active_machines[machine_idx]
                self.refresh_tables()
                messagebox.showinfo("Deleted", f"Entry {estimate_num} has been deleted")
    
    def delete_shipped_entry(self, machine_idx):
        """Delete a shipped repair entry"""
        if machine_idx < len(self.shipped_machines):
            machine = self.shipped_machines[machine_idx]
            
            # Confirm deletion
            result = messagebox.askyesno("Delete Entry", 
                                        f"Are you sure you want to delete this entry?\n\n"
                                        f"Estimate: {machine['estimate_number']}\n"
                                        f"Customer: {machine['customer_name']}\n"
                                        f"Make/Model: {machine['make']} {machine['model']}\n\n"
                                        "This action cannot be undone!",
                                        icon='warning')
            
            if result:
                # Delete PDF attachment if exists
                if machine.get('pdf_attachment'):
                    pdf_path = machine['pdf_attachment']
                    if os.path.exists(pdf_path):
                        try:
                            os.remove(pdf_path)
                        except:
                            pass  # Continue even if PDF deletion fails
                
                # Remove from list
                estimate_num = machine['estimate_number']
                del self.shipped_machines[machine_idx]
                self.refresh_tables()
                messagebox.showinfo("Deleted", f"Entry {estimate_num} has been deleted")
    
    def show_customer_details(self, event=None):
        """Show details for selected customer"""
        if not self.customer_select or not self.customer_details:
            return
            
        customer = self.customer_select.get()
        if not customer:
            return
        
        # Clear details
        self.customer_details.delete(1.0, tk.END)
        
        # Get customer's machines
        active = [m for m in self.active_machines if m['customer_name'] == customer]
        shipped = [m for m in self.shipped_machines if m['customer_name'] == customer]
        
        # Display details
        details = f"=== {customer} - Repair History ===\n\n"
        
        if active:
            details += f"ACTIVE REPAIRS ({len(active)}):\n"
            details += "-" * 50 + "\n"
            for machine in active:
                est_sent = "Est. Sent" if machine.get('estimate_sent', False) else "Est. Not Sent"
                est_complete = "Est. Complete" if machine.get('estimate_complete', False) else "Est. Incomplete"
                parts_ord = "Parts Ordered" if machine.get('parts_ordered', False) else "Parts Not Ordered"
                repair_done = "Repair Complete" if machine.get('repair_complete', False) else "Repair Incomplete"
                po_status = f"P.O. Received: {machine['po_number']}" if machine['po_received'] else "P.O. Pending"
                pdf_status = " | PDF: Yes" if machine.get('pdf_attachment') else ""
                details += f"• {machine['estimate_number']} - {machine['make']} {machine['model']}\n"
                details += f"  Serial: {machine['serial_number']} | Arrival: {machine['arrival_date']}\n"
                details += f"  {est_sent} | {est_complete} | {parts_ord} | {po_status}\n"
                details += f"  {repair_done}{pdf_status}\n\n"
        
        if shipped:
            details += f"\nPAST REPAIRS ({len(shipped)}):\n"
            details += "-" * 50 + "\n"
            for machine in shipped:
                po_text = f" | P.O.: {machine['po_number']}" if machine.get('po_number') else ""
                pdf_status = " | PDF: Yes" if machine.get('pdf_attachment') else ""
                details += f"• {machine['estimate_number']} - {machine['make']} {machine['model']}\n"
                details += f"  Serial: {machine['serial_number']} | Shipped: {machine.get('ship_date', 'N/A')}{po_text}{pdf_status}\n\n"
        
        if not active and not shipped:
            details += "No repairs found for this customer."
        
        self.customer_details.insert(1.0, details)
    
    def load_data_file(self):
        """Load an existing data file"""
        file_path = filedialog.askopenfilename(
            title="Open Data File",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.save_location)
        )
        
        if not file_path:
            return
        
        # Check if file exists
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "Selected file does not exist")
            return
        
        try:
            # Load data from the selected file
            with open(file_path, 'r') as f:
                data = json.load(f)
                new_active = data.get('active_machines', [])
                new_shipped = data.get('shipped_machines', [])
            
            # Ask user what to do with existing data
            if self.active_machines or self.shipped_machines:
                merge_window = tk.Toplevel(self.root)
                merge_window.title("Load Data File")
                merge_window.geometry("450x250")
                merge_window.configure(bg='#2c3e50')
                merge_window.transient(self.root)
                merge_window.grab_set()
                
                # Center the window
                merge_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 200, self.root.winfo_rooty() + 150))
                
                tk.Label(merge_window, text="Load Data File", 
                        font=('Arial', 16, 'bold'), fg='white', bg='#2c3e50').pack(pady=20)
                
                info_text = f"File contains:\n"
                info_text += f"  • {len(new_active)} active repairs\n"
                info_text += f"  • {len(new_shipped)} past repairs\n\n"
                info_text += f"Current data:\n"
                info_text += f"  • {len(self.active_machines)} active repairs\n"
                info_text += f"  • {len(self.shipped_machines)} past repairs"
                
                tk.Label(merge_window, text=info_text, 
                        font=('Arial', 10), fg='#ecf0f1', bg='#2c3e50', justify='left').pack(pady=10)
                
                tk.Label(merge_window, text="How do you want to load this file?", 
                        font=('Arial', 11, 'bold'), fg='#ecf0f1', bg='#2c3e50').pack(pady=10)
                
                result = {'action': None}
                
                def replace_data():
                    result['action'] = 'replace'
                    merge_window.destroy()
                
                def merge_data():
                    result['action'] = 'merge'
                    merge_window.destroy()
                
                def cancel():
                    result['action'] = 'cancel'
                    merge_window.destroy()
                
                btn_frame = tk.Frame(merge_window, bg='#2c3e50')
                btn_frame.pack(pady=10)
                
                tk.Button(btn_frame, text="Replace Current Data", command=replace_data,
                         bg='#e74c3c', fg='white', font=('Arial', 10, 'bold'), width=18).pack(side='left', padx=5)
                tk.Button(btn_frame, text="Merge with Current", command=merge_data,
                         bg='#27ae60', fg='white', font=('Arial', 10, 'bold'), width=18).pack(side='left', padx=5)
                tk.Button(btn_frame, text="Cancel", command=cancel,
                         bg='#95a5a6', fg='white', font=('Arial', 10, 'bold'), width=10).pack(side='left', padx=5)
                
                merge_window.wait_window()
                
                if result['action'] == 'cancel':
                    return
                elif result['action'] == 'replace':
                    self.active_machines = new_active
                    self.shipped_machines = new_shipped
                    self.save_location = file_path
                    self.save_settings()
                    self.refresh_tables()
                    messagebox.showinfo("Success", f"Data loaded from:\n{os.path.basename(file_path)}\n\nCurrent data was replaced.")
                elif result['action'] == 'merge':
                    # Merge by checking for duplicate estimate numbers
                    existing_estimates = set(m['estimate_number'] for m in self.active_machines + self.shipped_machines)
                    
                    merged_active = 0
                    merged_shipped = 0
                    skipped = 0
                    
                    for machine in new_active:
                        if machine['estimate_number'] not in existing_estimates:
                            self.active_machines.append(machine)
                            existing_estimates.add(machine['estimate_number'])
                            merged_active += 1
                        else:
                            skipped += 1
                    
                    for machine in new_shipped:
                        if machine['estimate_number'] not in existing_estimates:
                            self.shipped_machines.append(machine)
                            existing_estimates.add(machine['estimate_number'])
                            merged_shipped += 1
                        else:
                            skipped += 1
                    
                    self.save_location = file_path
                    self.save_settings()
                    self.refresh_tables()
                    
                    merge_msg = f"Data merged from:\n{os.path.basename(file_path)}\n\n"
                    merge_msg += f"Added: {merged_active} active, {merged_shipped} past\n"
                    if skipped > 0:
                        merge_msg += f"Skipped {skipped} duplicate(s)"
                    messagebox.showinfo("Success", merge_msg)
            
            else:
                # No existing data, just load it
                self.active_machines = new_active
                self.shipped_machines = new_shipped
                self.save_location = file_path
                self.save_settings()
                self.refresh_tables()
                messagebox.showinfo("Success", f"Data loaded from:\n{os.path.basename(file_path)}")
            
            if self.status_label:
                self.status_label.config(text=f"✅ Loaded: {os.path.basename(file_path)}")
                
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Invalid JSON file format")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data file:\n\n{str(e)}")
    
    def set_save_location(self):
        """Set the save location for data files"""
        new_location = filedialog.asksaveasfilename(
            title="Choose Save Location",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.save_location),
            initialfile="repair_data.json"
        )
        if new_location:
            self.save_location = new_location
            self.save_settings()
            self.save_data()
            if self.status_label:
                self.status_label.config(text=f"Save location updated: {os.path.basename(new_location)}")
    
    def show_export_menu(self):
        """Show comprehensive export menu options"""
        export_window = tk.Toplevel(self.root)
        export_window.title("Export Options")
        export_window.geometry("500x700")
        export_window.configure(bg='#2c3e50')
        export_window.transient(self.root)
        export_window.grab_set()
        
        # Center the window
        export_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 100, self.root.winfo_rooty() + 100))
        
        # Title
        title_label = tk.Label(export_window, text="Export Data", font=('Arial', 16, 'bold'),
                              fg='white', bg='#2c3e50')
        title_label.pack(pady=20)
        
        # Check if Excel is available
        if not EXCEL_AVAILABLE:
            warning_label = tk.Label(export_window, text="⚠️ Excel exports require openpyxl package\nInstall with: pip install openpyxl", 
                                   font=('Arial', 10), fg='#e74c3c', bg='#2c3e50')
            warning_label.pack(pady=5)
        
        # Excel exports section
        excel_frame = tk.LabelFrame(export_window, text="Excel Exports", font=('Arial', 12, 'bold'),
                                   fg='#ecf0f1', bg='#34495e')
        excel_frame.pack(fill='x', padx=20, pady=10)
        
        # All data export
        tk.Button(excel_frame, text="Export All Data to Excel", command=self.export_all_excel,
                 bg='#3498db', fg='white', font=('Arial', 10, 'bold'), width=40,
                 state='normal' if EXCEL_AVAILABLE else 'disabled').pack(pady=3)
        
        # Customer-specific exports
        tk.Button(excel_frame, text="Export by Customer (All Data)", command=self.export_by_customer_excel,
                 bg='#9b59b6', fg='white', font=('Arial', 10, 'bold'), width=40,
                 state='normal' if EXCEL_AVAILABLE else 'disabled').pack(pady=3)
        
        # Active repairs only
        tk.Button(excel_frame, text="Export Active Repairs Only", command=self.export_active_only_excel,
                 bg='#f39c12', fg='white', font=('Arial', 10, 'bold'), width=40,
                 state='normal' if EXCEL_AVAILABLE else 'disabled').pack(pady=3)
        
        # Past repairs only
        tk.Button(excel_frame, text="Export Past Repairs Only", command=self.export_past_only_excel,
                 bg='#e67e22', fg='white', font=('Arial', 10, 'bold'), width=40,
                 state='normal' if EXCEL_AVAILABLE else 'disabled').pack(pady=3)
        
        # Customer active only
        tk.Button(excel_frame, text="Export Customer Active Repairs Only", command=self.export_customer_active_excel,
                 bg='#1abc9c', fg='white', font=('Arial', 10, 'bold'), width=40,
                 state='normal' if EXCEL_AVAILABLE else 'disabled').pack(pady=3)
        
        # Customer past only
        tk.Button(excel_frame, text="Export Customer Past Repairs Only", command=self.export_customer_past_excel,
                 bg='#16a085', fg='white', font=('Arial', 10, 'bold'), width=40,
                 state='normal' if EXCEL_AVAILABLE else 'disabled').pack(pady=3)
        
        # Customer active/past separate sheets
        tk.Button(excel_frame, text="Export Customer Active/Past Separate Sheets", command=self.export_customer_separate_excel,
                 bg='#8e44ad', fg='white', font=('Arial', 10, 'bold'), width=40,
                 state='normal' if EXCEL_AVAILABLE else 'disabled').pack(pady=3)
        
        # Export with PDF links
        tk.Button(excel_frame, text="Export All Data with PDF Links", command=self.export_with_pdf_links,
                 bg='#d35400', fg='white', font=('Arial', 10, 'bold'), width=40,
                 state='normal' if EXCEL_AVAILABLE else 'disabled').pack(pady=3)
        
        # CSV exports section (as backup option)
        csv_frame = tk.LabelFrame(export_window, text="CSV Exports (Backup Option)", font=('Arial', 12, 'bold'),
                                 fg='#ecf0f1', bg='#34495e')
        csv_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Button(csv_frame, text="Export All Data to CSV", command=self.export_all_csv,
                 bg='#27ae60', fg='white', font=('Arial', 10, 'bold'), width=40).pack(pady=3)
        
        tk.Button(csv_frame, text="Export by Customer to CSV", command=self.export_by_customer_csv,
                 bg='#2ecc71', fg='white', font=('Arial', 10, 'bold'), width=40).pack(pady=3)
        
        # JSON backup section
        json_frame = tk.LabelFrame(export_window, text="Backup Options", font=('Arial', 12, 'bold'),
                                  fg='#ecf0f1', bg='#34495e')
        json_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Button(json_frame, text="Create JSON Backup", command=self.create_json_backup,
                 bg='#34495e', fg='white', font=('Arial', 10, 'bold'), width=40).pack(pady=3)
        
        # Export attachments
        tk.Button(json_frame, text="Export All Attachments", command=self.export_all_attachments,
                 bg='#7f8c8d', fg='white', font=('Arial', 10, 'bold'), width=40).pack(pady=3)
        
        # Close button
        tk.Button(export_window, text="Close", command=export_window.destroy,
                 bg='#95a5a6', fg='white', font=('Arial', 10, 'bold')).pack(pady=20)
    
    # Excel Export Functions
    def create_styled_excel_workbook(self):
        """Create a styled Excel workbook with formatting"""
        if not EXCEL_AVAILABLE:
            raise ImportError("openpyxl package is required for Excel exports")
        
        wb = Workbook()
        
        # Define styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        return wb, header_font, header_fill, header_alignment, thin_border
    
    def style_excel_sheet(self, sheet, header_font, header_fill, header_alignment, thin_border):
        """Apply styling to an Excel sheet"""
        # Style header row
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # Apply borders to all cells with data
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.border = thin_border
        
        # Auto-adjust column widths
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            sheet.column_dimensions[column_letter].width = adjusted_width
    
    def export_all_excel(self):
        """Export all data to Excel with separate sheets for active and past repairs"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires the openpyxl package.\nPlease install it with: pip install openpyxl")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export All Data to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"all_repairs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                wb, header_font, header_fill, header_alignment, thin_border = self.create_styled_excel_workbook()
                
                # Remove default sheet and create our sheets
                wb.remove(wb.active)
                
                # Active Repairs Sheet
                active_sheet = wb.create_sheet("Active Repairs")
                active_headers = ['Estimate #', 'Customer', 'Field/Shop', 'Arrival Date', 'Serial #', 
                                'Make', 'Model', 'Est. Sent', 'Est. Complete', 'Parts Ordered', 'P.O. Received', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                active_sheet.append(active_headers)
                
                for machine in self.active_machines:
                    active_sheet.append([
                        machine['estimate_number'],
                        machine['customer_name'],
                        machine.get('service_location', 'Shop'),
                        machine['arrival_date'],
                        machine['serial_number'],
                        machine['make'],
                        machine['model'],
                        'Yes' if machine.get('estimate_sent', False) else 'No',
                        'Yes' if machine.get('estimate_complete', False) else 'No',
                        'Yes' if machine.get('parts_ordered', False) else 'No',
                        'Yes' if machine['po_received'] else 'No',
                        machine.get('po_number', ''),
                        'Yes' if machine.get('repair_complete', False) else 'No',
                        'Yes' if machine.get('pdf_attachment') else 'No'
                    ])
                
                self.style_excel_sheet(active_sheet, header_font, header_fill, header_alignment, thin_border)
                
                # Past Repairs Sheet
                past_sheet = wb.create_sheet("Past Repairs")
                past_headers = ['Estimate #', 'Customer', 'Field/Shop', 'Arrival Date', 'Ship Date',
                              'Serial #', 'Make', 'Model', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                past_sheet.append(past_headers)
                
                for machine in self.shipped_machines:
                    past_sheet.append([
                        machine['estimate_number'],
                        machine['customer_name'],
                        machine.get('service_location', 'Shop'),
                        machine['arrival_date'],
                        machine.get('ship_date', ''),
                        machine['serial_number'],
                        machine['make'],
                        machine['model'],
                        machine.get('po_number', ''),
                        'Yes' if machine.get('repair_complete', False) else 'No',
                        'Yes' if machine.get('pdf_attachment') else 'No'
                    ])
                
                self.style_excel_sheet(past_sheet, header_font, header_fill, header_alignment, thin_border)
                
                # Summary Sheet
                summary_sheet = wb.create_sheet("Summary")
                summary_sheet.append(['Repair Management Summary', '', ''])
                summary_sheet.append(['Total Active Repairs:', len(self.active_machines), ''])
                summary_sheet.append(['Total Past Repairs:', len(self.shipped_machines), ''])
                summary_sheet.append(['Total All Repairs:', len(self.active_machines) + len(self.shipped_machines), ''])
                summary_sheet.append(['Active with PDF:', sum(1 for m in self.active_machines if m.get('pdf_attachment')), ''])
                summary_sheet.append(['Past with PDF:', sum(1 for m in self.shipped_machines if m.get('pdf_attachment')), ''])
                summary_sheet.append(['Report Generated:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'), ''])
                
                # Style summary sheet
                summary_sheet['A1'].font = Font(bold=True, size=14)
                for row in range(2, 9):
                    summary_sheet[f'A{row}'].font = Font(bold=True)
                
                wb.save(filename)
                messagebox.showinfo("Success", f"All data exported to Excel: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
    
    def export_with_pdf_links(self):
        """Export data with clickable PDF links (Windows only)"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires the openpyxl package.\nPlease install it with: pip install openpyxl")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Data with PDF Links",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"repairs_with_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                wb, header_font, header_fill, header_alignment, thin_border = self.create_styled_excel_workbook()
                wb.remove(wb.active)
                
                # Combined data sheet
                sheet = wb.create_sheet("All Repairs")
                headers = ['Estimate #', 'Customer', 'Field/Shop', 'Arrival Date', 'Ship Date', 'Serial #', 
                          'Make', 'Model', 'P.O. Received', 'P.O. Number', 'PDF File', 'Status']
                sheet.append(headers)
                
                # Add active machines
                for machine in self.active_machines:
                    pdf_path = machine.get('pdf_attachment', '')
                    pdf_display = os.path.basename(pdf_path) if pdf_path else 'No PDF'
                    sheet.append([
                        machine['estimate_number'],
                        machine['customer_name'],
                        machine.get('service_location', 'Shop'),
                        machine['arrival_date'],
                        '',
                        machine['serial_number'],
                        machine['make'],
                        machine['model'],
                        'Yes' if machine['po_received'] else 'No',
                        machine.get('po_number', ''),
                        pdf_display,
                        'Active'
                    ])
                
                # Add shipped machines
                for machine in self.shipped_machines:
                    pdf_path = machine.get('pdf_attachment', '')
                    pdf_display = os.path.basename(pdf_path) if pdf_path else 'No PDF'
                    sheet.append([
                        machine['estimate_number'],
                        machine['customer_name'],
                        machine.get('service_location', 'Shop'),
                        machine['arrival_date'],
                        machine.get('ship_date', ''),
                        machine['serial_number'],
                        machine['make'],
                        machine['model'],
                        'Yes' if machine.get('po_received', False) else 'No',
                        machine.get('po_number', ''),
                        pdf_display,
                        'Shipped'
                    ])
                
                self.style_excel_sheet(sheet, header_font, header_fill, header_alignment, thin_border)
                wb.save(filename)
                messagebox.showinfo("Success", f"Data with PDF links exported to Excel: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
    
    def get_customer_list(self):
        """Get list of all customers"""
        all_customers = set()
        for machine in self.active_machines + self.shipped_machines:
            all_customers.add(machine['customer_name'])
        
        if not all_customers:
            messagebox.showwarning("Warning", "No customer data to export")
            return []
        
        return sorted(list(all_customers))
    
    def choose_customer_dialog(self, customers, title):
        """Show dialog to choose a customer"""
        customer_window = tk.Toplevel(self.root)
        customer_window.title(title)
        customer_window.geometry("350x200")
        customer_window.configure(bg='#2c3e50')
        customer_window.transient(self.root)
        customer_window.grab_set()
        
        # Center the window
        customer_window.geometry("+%d+%d" % (self.root.winfo_rootx() + 200, self.root.winfo_rooty() + 200))
        
        tk.Label(customer_window, text="Select Customer:", 
                fg='white', bg='#2c3e50', font=('Arial', 12, 'bold')).pack(pady=20)
        
        customer_var = tk.StringVar()
        customer_combo = ttk.Combobox(customer_window, textvariable=customer_var, 
                                     values=customers, width=30, state='readonly')
        customer_combo.pack(pady=10)
        customer_combo.focus()
        
        result = {'customer': None}
        
        def on_export():
            selected = customer_var.get()
            if not selected:
                messagebox.showwarning("Warning", "Please select a customer")
                return
            result['customer'] = selected
            customer_window.destroy()
        
        def on_cancel():
            customer_window.destroy()
        
        btn_frame = tk.Frame(customer_window, bg='#2c3e50')
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="Export", command=on_export,
                 bg='#27ae60', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Cancel", command=on_cancel,
                 bg='#95a5a6', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        
        # Bind Enter key to export
        customer_window.bind('<Return>', lambda e: on_export())
        customer_window.bind('<Escape>', lambda e: on_cancel())
        
        customer_window.wait_window()
        return result['customer']
    
    def export_by_customer_excel(self):
        """Export data for a specific customer to Excel"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires the openpyxl package.\nPlease install it with: pip install openpyxl")
            return
        
        customers = self.get_customer_list()
        if not customers:
            return
        
        customer = self.choose_customer_dialog(customers, "Export Customer Data to Excel")
        if not customer:
            return
        
        filename = filedialog.asksaveasfilename(
            title=f"Export {customer} Data to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{customer.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                wb, header_font, header_fill, header_alignment, thin_border = self.create_styled_excel_workbook()
                wb.remove(wb.active)
                
                # Get customer data
                customer_active = [m for m in self.active_machines if m['customer_name'] == customer]
                customer_past = [m for m in self.shipped_machines if m['customer_name'] == customer]
                
                # Customer Active Sheet
                if customer_active:
                    active_sheet = wb.create_sheet(f"{customer} - Active")
                    headers = ['Estimate #', 'Field/Shop', 'Arrival Date', 'Serial #', 'Make', 'Model', 'Est. Sent', 'Est. Complete', 'Parts Ordered', 'P.O. Received', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                    active_sheet.append(headers)
                    
                    for machine in customer_active:
                        active_sheet.append([
                            machine['estimate_number'],
                            machine.get('service_location', 'Shop'),
                            machine['arrival_date'],
                            machine['serial_number'],
                            machine['make'],
                            machine['model'],
                            'Yes' if machine.get('estimate_sent', False) else 'No',
                            'Yes' if machine.get('estimate_complete', False) else 'No',
                            'Yes' if machine.get('parts_ordered', False) else 'No',
                            'Yes' if machine['po_received'] else 'No',
                            machine.get('po_number', ''),
                            'Yes' if machine.get('repair_complete', False) else 'No',
                            'Yes' if machine.get('pdf_attachment') else 'No'
                        ])
                    
                    self.style_excel_sheet(active_sheet, header_font, header_fill, header_alignment, thin_border)
                
                # Customer Past Sheet
                if customer_past:
                    past_sheet = wb.create_sheet(f"{customer} - Past")
                    headers = ['Estimate #', 'Field/Shop', 'Arrival Date', 'Ship Date', 'Serial #', 'Make', 'Model', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                    past_sheet.append(headers)
                    
                    for machine in customer_past:
                        past_sheet.append([
                            machine['estimate_number'],
                            machine.get('service_location', 'Shop'),
                            machine['arrival_date'],
                            machine.get('ship_date', ''),
                            machine['serial_number'],
                            machine['make'],
                            machine['model'],
                            machine.get('po_number', ''),
                            'Yes' if machine.get('repair_complete', False) else 'No',
                            'Yes' if machine.get('pdf_attachment') else 'No'
                        ])
                    
                    self.style_excel_sheet(past_sheet, header_font, header_fill, header_alignment, thin_border)
                
                # Summary sheet
                summary_sheet = wb.create_sheet(f"{customer} - Summary")
                summary_sheet.append([f'{customer} Repair Summary', '', ''])
                summary_sheet.append(['Active Repairs:', len(customer_active), ''])
                summary_sheet.append(['Past Repairs:', len(customer_past), ''])
                summary_sheet.append(['Total Repairs:', len(customer_active) + len(customer_past), ''])
                summary_sheet.append(['Active with PDF:', sum(1 for m in customer_active if m.get('pdf_attachment')), ''])
                summary_sheet.append(['Past with PDF:', sum(1 for m in customer_past if m.get('pdf_attachment')), ''])
                summary_sheet.append(['Report Generated:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'), ''])
                
                summary_sheet['A1'].font = Font(bold=True, size=14)
                for row in range(2, 8):
                    summary_sheet[f'A{row}'].font = Font(bold=True)
                
                wb.save(filename)
                messagebox.showinfo("Success", f"{customer} data exported to Excel: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
    
    def export_active_only_excel(self):
        """Export only active repairs to Excel"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires the openpyxl package.\nPlease install it with: pip install openpyxl")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Active Repairs to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"active_repairs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                wb, header_font, header_fill, header_alignment, thin_border = self.create_styled_excel_workbook()
                sheet = wb.active
                sheet.title = "Active Repairs"
                
                headers = ['Estimate #', 'Customer', 'Field/Shop', 'Arrival Date', 'Serial #', 
                          'Make', 'Model', 'Est. Sent', 'Est. Complete', 'Parts Ordered', 'P.O. Received', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                sheet.append(headers)
                
                for machine in self.active_machines:
                    sheet.append([
                        machine['estimate_number'],
                        machine['customer_name'],
                        machine.get('service_location', 'Shop'),
                        machine['arrival_date'],
                        machine['serial_number'],
                        machine['make'],
                        machine['model'],
                        'Yes' if machine.get('estimate_sent', False) else 'No',
                        'Yes' if machine.get('estimate_complete', False) else 'No',
                        'Yes' if machine.get('parts_ordered', False) else 'No',
                        'Yes' if machine['po_received'] else 'No',
                        machine.get('po_number', ''),
                        'Yes' if machine.get('repair_complete', False) else 'No',
                        'Yes' if machine.get('pdf_attachment') else 'No'
                    ])
                
                self.style_excel_sheet(sheet, header_font, header_fill, header_alignment, thin_border)
                
                wb.save(filename)
                messagebox.showinfo("Success", f"Active repairs exported to Excel: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
    
    def export_past_only_excel(self):
        """Export only past repairs to Excel"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires the openpyxl package.\nPlease install it with: pip install openpyxl")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Past Repairs to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"past_repairs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                wb, header_font, header_fill, header_alignment, thin_border = self.create_styled_excel_workbook()
                sheet = wb.active
                sheet.title = "Past Repairs"
                
                headers = ['Estimate #', 'Customer', 'Field/Shop', 'Arrival Date', 'Ship Date',
                          'Serial #', 'Make', 'Model', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                sheet.append(headers)
                
                for machine in self.shipped_machines:
                    sheet.append([
                        machine['estimate_number'],
                        machine['customer_name'],
                        machine.get('service_location', 'Shop'),
                        machine['arrival_date'],
                        machine.get('ship_date', ''),
                        machine['serial_number'],
                        machine['make'],
                        machine['model'],
                        machine.get('po_number', ''),
                        'Yes' if machine.get('repair_complete', False) else 'No',
                        'Yes' if machine.get('pdf_attachment') else 'No'
                    ])
                
                self.style_excel_sheet(sheet, header_font, header_fill, header_alignment, thin_border)
                
                wb.save(filename)
                messagebox.showinfo("Success", f"Past repairs exported to Excel: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
    
    def export_customer_active_excel(self):
        """Export only active repairs for a specific customer"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires the openpyxl package.\nPlease install it with: pip install openpyxl")
            return
        
        customers = self.get_customer_list()
        if not customers:
            return
        
        customer = self.choose_customer_dialog(customers, "Export Customer Active Repairs")
        if not customer:
            return
        
        customer_active = [m for m in self.active_machines if m['customer_name'] == customer]
        if not customer_active:
            messagebox.showwarning("Warning", f"No active repairs found for {customer}")
            return
        
        filename = filedialog.asksaveasfilename(
            title=f"Export {customer} Active Repairs",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{customer.replace(' ', '_')}_active_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                wb, header_font, header_fill, header_alignment, thin_border = self.create_styled_excel_workbook()
                sheet = wb.active
                sheet.title = f"{customer} - Active"
                
                headers = ['Estimate #', 'Field/Shop', 'Arrival Date', 'Serial #', 'Make', 'Model', 'Est. Sent', 'Est. Complete', 'Parts Ordered', 'P.O. Received', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                sheet.append(headers)
                
                for machine in customer_active:
                    sheet.append([
                        machine['estimate_number'],
                        machine.get('service_location', 'Shop'),
                        machine['arrival_date'],
                        machine['serial_number'],
                        machine['make'],
                        machine['model'],
                        'Yes' if machine.get('estimate_sent', False) else 'No',
                        'Yes' if machine.get('estimate_complete', False) else 'No',
                        'Yes' if machine.get('parts_ordered', False) else 'No',
                        'Yes' if machine['po_received'] else 'No',
                        machine.get('po_number', ''),
                        'Yes' if machine.get('repair_complete', False) else 'No',
                        'Yes' if machine.get('pdf_attachment') else 'No'
                    ])
                
                self.style_excel_sheet(sheet, header_font, header_fill, header_alignment, thin_border)
                
                wb.save(filename)
                messagebox.showinfo("Success", f"{customer} active repairs exported to Excel: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
    
    def export_customer_past_excel(self):
        """Export only past repairs for a specific customer"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires the openpyxl package.\nPlease install it with: pip install openpyxl")
            return
        
        customers = self.get_customer_list()
        if not customers:
            return
        
        customer = self.choose_customer_dialog(customers, "Export Customer Past Repairs")
        if not customer:
            return
        
        customer_past = [m for m in self.shipped_machines if m['customer_name'] == customer]
        if not customer_past:
            messagebox.showwarning("Warning", f"No past repairs found for {customer}")
            return
        
        filename = filedialog.asksaveasfilename(
            title=f"Export {customer} Past Repairs",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{customer.replace(' ', '_')}_past_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                wb, header_font, header_fill, header_alignment, thin_border = self.create_styled_excel_workbook()
                sheet = wb.active
                sheet.title = f"{customer} - Past"
                
                headers = ['Estimate #', 'Field/Shop', 'Arrival Date', 'Ship Date', 'Serial #', 'Make', 'Model', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                sheet.append(headers)
                
                for machine in customer_past:
                    sheet.append([
                        machine['estimate_number'],
                        machine.get('service_location', 'Shop'),
                        machine['arrival_date'],
                        machine.get('ship_date', ''),
                        machine['serial_number'],
                        machine['make'],
                        machine['model'],
                        machine.get('po_number', ''),
                        'Yes' if machine.get('repair_complete', False) else 'No',
                        'Yes' if machine.get('pdf_attachment') else 'No'
                    ])
                
                self.style_excel_sheet(sheet, header_font, header_fill, header_alignment, thin_border)
                
                wb.save(filename)
                messagebox.showinfo("Success", f"{customer} past repairs exported to Excel: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
    
    def export_customer_separate_excel(self):
        """Export customer data with active and past repairs on separate sheets"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires the openpyxl package.\nPlease install it with: pip install openpyxl")
            return
        
        customers = self.get_customer_list()
        if not customers:
            return
        
        customer = self.choose_customer_dialog(customers, "Export Customer Data (Separate Sheets)")
        if not customer:
            return
        
        customer_active = [m for m in self.active_machines if m['customer_name'] == customer]
        customer_past = [m for m in self.shipped_machines if m['customer_name'] == customer]
        
        if not customer_active and not customer_past:
            messagebox.showwarning("Warning", f"No repairs found for {customer}")
            return
        
        filename = filedialog.asksaveasfilename(
            title=f"Export {customer} Data (Separate Sheets)",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{customer.replace(' ', '_')}_separate_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            # This is the same as export_by_customer_excel, just with a different dialog
            try:
                wb, header_font, header_fill, header_alignment, thin_border = self.create_styled_excel_workbook()
                wb.remove(wb.active)
                
                # Active repairs sheet
                if customer_active:
                    active_sheet = wb.create_sheet("Active Repairs")
                    headers = ['Estimate #', 'Field/Shop', 'Arrival Date', 'Serial #', 'Make', 'Model', 'Est. Sent', 'Est. Complete', 'Parts Ordered', 'P.O. Received', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                    active_sheet.append(headers)
                    
                    for machine in customer_active:
                        active_sheet.append([
                            machine['estimate_number'],
                            machine.get('service_location', 'Shop'),
                            machine['arrival_date'],
                            machine['serial_number'],
                            machine['make'],
                            machine['model'],
                            'Yes' if machine.get('estimate_sent', False) else 'No',
                            'Yes' if machine.get('estimate_complete', False) else 'No',
                            'Yes' if machine.get('parts_ordered', False) else 'No',
                            'Yes' if machine['po_received'] else 'No',
                            machine.get('po_number', ''),
                            'Yes' if machine.get('repair_complete', False) else 'No',
                            'Yes' if machine.get('pdf_attachment') else 'No'
                        ])
                    
                    self.style_excel_sheet(active_sheet, header_font, header_fill, header_alignment, thin_border)
                
                # Past repairs sheet
                if customer_past:
                    past_sheet = wb.create_sheet("Past Repairs")
                    headers = ['Estimate #', 'Field/Shop', 'Arrival Date', 'Ship Date', 'Serial #', 'Make', 'Model', 'P.O. Number', 'Repair Complete', 'PDF Attachment']
                    past_sheet.append(headers)
                    
                    for machine in customer_past:
                        past_sheet.append([
                            machine['estimate_number'],
                            machine.get('service_location', 'Shop'),
                            machine['arrival_date'],
                            machine.get('ship_date', ''),
                            machine['serial_number'],
                            machine['make'],
                            machine['model'],
                            machine.get('po_number', ''),
                            'Yes' if machine.get('repair_complete', False) else 'No',
                            'Yes' if machine.get('pdf_attachment') else 'No'
                        ])
                    
                    self.style_excel_sheet(past_sheet, header_font, header_fill, header_alignment, thin_border)
                
                # Summary sheet
                summary_sheet = wb.create_sheet("Summary")
                summary_sheet.append([f'{customer} Repair Summary', '', ''])
                summary_sheet.append(['Active Repairs:', len(customer_active), ''])
                summary_sheet.append(['Past Repairs:', len(customer_past), ''])
                summary_sheet.append(['Total Repairs:', len(customer_active) + len(customer_past), ''])
                summary_sheet.append(['Active with PDF:', sum(1 for m in customer_active if m.get('pdf_attachment')), ''])
                summary_sheet.append(['Past with PDF:', sum(1 for m in customer_past if m.get('pdf_attachment')), ''])
                summary_sheet.append(['Report Generated:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'), ''])
                
                summary_sheet['A1'].font = Font(bold=True, size=14)
                for row in range(2, 8):
                    summary_sheet[f'A{row}'].font = Font(bold=True)
                
                wb.save(filename)
                messagebox.showinfo("Success", f"{customer} data exported to Excel: {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export to Excel: {str(e)}")
    
    def export_all_attachments(self):
        """Export all PDF attachments to a zip file"""
        import zipfile
        
        # Get all attachments
        all_attachments = []
        for machine in self.active_machines + self.shipped_machines:
            if machine.get('pdf_attachment'):
                all_attachments.append(machine['pdf_attachment'])
        
        if not all_attachments:
            messagebox.showwarning("Warning", "No PDF attachments found to export")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export All Attachments",
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip")],
            initialfile=f"attachments_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        )
        
        if filename:
            try:
                with zipfile.ZipFile(filename, 'w') as zipf:
                    for attachment in all_attachments:
                        if os.path.exists(attachment):
                            zipf.write(attachment, os.path.basename(attachment))
                
                messagebox.showinfo("Success", f"All attachments exported to: {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export attachments: {str(e)}")
    
    # CSV Export Functions (Backup Options)
    def export_all_csv(self):
        """Export all data to CSV"""
        filename = filedialog.asksaveasfilename(
            title="Export All Data",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"all_repairs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if filename:
            try:
                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # Write header
                    writer.writerow(['Estimate #', 'Customer', 'Field/Shop', 'Arrival Date', 'Ship Date', 
                                   'Serial #', 'Make', 'Model', 'Est. Sent', 'Est. Complete', 'Parts Ordered', 'P.O. Received', 'P.O. Number', 'Repair Complete', 'PDF Attachment', 'Status'])
                    
                    # Write active machines
                    for machine in self.active_machines:
                        writer.writerow([
                            machine['estimate_number'],
                            machine['customer_name'],
                            machine.get('service_location', 'Shop'),
                            machine['arrival_date'],
                            '',  # No ship date for active
                            machine['serial_number'],
                            machine['make'],
                            machine['model'],
                            'Yes' if machine.get('estimate_sent', False) else 'No',
                            'Yes' if machine.get('estimate_complete', False) else 'No',
                            'Yes' if machine.get('parts_ordered', False) else 'No',
                            'Yes' if machine['po_received'] else 'No',
                            machine.get('po_number', ''),
                            'Yes' if machine.get('repair_complete', False) else 'No',
                            'Yes' if machine.get('pdf_attachment') else 'No',
                            'Active'
                        ])
                    
                    # Write shipped machines
                    for machine in self.shipped_machines:
                        writer.writerow([
                            machine['estimate_number'],
                            machine['customer_name'],
                            machine.get('service_location', 'Shop'),
                            machine['arrival_date'],
                            machine.get('ship_date', ''),
                            machine['serial_number'],
                            machine['make'],
                            machine['model'],
                            'N/A',  # Est. Sent not applicable for shipped
                            'N/A',  # Est. Complete not applicable for shipped
                            'N/A',  # Parts Ordered not applicable for shipped
                            'Yes' if machine.get('po_received', False) else 'No',
                            machine.get('po_number', ''),
                            'Yes' if machine.get('repair_complete', False) else 'No',
                            'Yes' if machine.get('pdf_attachment') else 'No',
                            'Shipped'
                        ])
                
                messagebox.showinfo("Success", f"Data exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def export_by_customer_csv(self):
        """Export data for a specific customer to CSV (backup option)"""
        customers = self.get_customer_list()
        if not customers:
            return
        
        customer = self.choose_customer_dialog(customers, "Export Customer Data to CSV")
        if not customer:
            return
        
        filename = filedialog.asksaveasfilename(
            title=f"Export {customer} Data to CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"{customer.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        if filename:
            try:
                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # Write header
                    writer.writerow(['Estimate #', 'Customer', 'Field/Shop', 'Arrival Date', 'Ship Date',
                                   'Serial #', 'Make', 'Model', 'Est. Sent', 'Est. Complete', 'Parts Ordered', 'P.O. Received', 'P.O. Number', 'Repair Complete', 'PDF Attachment', 'Status'])
                    
                    # Write customer's active machines
                    for machine in self.active_machines:
                        if machine['customer_name'] == customer:
                            writer.writerow([
                                machine['estimate_number'],
                                machine['customer_name'],
                                machine.get('service_location', 'Shop'),
                                machine['arrival_date'],
                                '',  # No ship date for active
                                machine['serial_number'],
                                machine['make'],
                                machine['model'],
                                'Yes' if machine.get('estimate_sent', False) else 'No',
                                'Yes' if machine.get('estimate_complete', False) else 'No',
                                'Yes' if machine.get('parts_ordered', False) else 'No',
                                'Yes' if machine['po_received'] else 'No',
                                machine.get('po_number', ''),
                                'Yes' if machine.get('repair_complete', False) else 'No',
                                'Yes' if machine.get('pdf_attachment') else 'No',
                                'Active'
                            ])
                    
                    # Write customer's shipped machines
                    for machine in self.shipped_machines:
                        if machine['customer_name'] == customer:
                            writer.writerow([
                                machine['estimate_number'],
                                machine['customer_name'],
                                machine.get('service_location', 'Shop'),
                                machine['arrival_date'],
                                machine.get('ship_date', ''),
                                machine['serial_number'],
                                machine['make'],
                                machine['model'],
                                'N/A',  # Est. Sent not applicable for shipped
                                'N/A',  # Est. Complete not applicable for shipped
                                'N/A',  # Parts Ordered not applicable for shipped
                                'Yes' if machine.get('po_received', False) else 'No',
                                machine.get('po_number', ''),
                                'Yes' if machine.get('repair_complete', False) else 'No',
                                'Yes' if machine.get('pdf_attachment') else 'No',
                                'Shipped'
                            ])
                
                messagebox.showinfo("Success", f"{customer} data exported to CSV: {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def create_json_backup(self):
        """Create a JSON backup of all data"""
        filename = filedialog.asksaveasfilename(
            title="Create JSON Backup",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile=f"repair_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        if filename:
            try:
                backup_data = {
                    'active_machines': self.active_machines,
                    'shipped_machines': self.shipped_machines,
                    'backup_created': datetime.now().isoformat(),
                    'total_active': len(self.active_machines),
                    'total_shipped': len(self.shipped_machines),
                    'total_with_pdf': sum(1 for m in self.active_machines + self.shipped_machines if m.get('pdf_attachment'))
                }
                
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(backup_data, f, indent=2)
                
                messagebox.showinfo("Success", f"Backup created: {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create backup: {str(e)}")
    
    def open_webcam_scanner(self):
        """Open a webcam window to scan barcodes visually"""
        if not WEBCAM_AVAILABLE:
            error_details = '\n'.join(WEBCAM_IMPORT_ERRORS) if WEBCAM_IMPORT_ERRORS else 'Unknown import error'
            messagebox.showerror("Webcam Not Available",
                f"Webcam scanning requires additional packages.\n\n"
                f"Import errors:\n{error_details}\n\n"
                "Install with:\n"
                "pip install opencv-python pyzbar")
            return
        
        # Try to open the default camera
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            messagebox.showerror("Webcam Error", "Could not access webcam.\nMake sure it's connected and not in use by another program.")
            return
        
        # Create window
        scan_win = tk.Toplevel(self.root)
        scan_win.title("Webcam Barcode Scanner")
        scan_win.configure(bg='#2c3e50')
        scan_win.transient(self.root)
        
        # State
        scan_state = {'running': True, 'found': None, 'cap': cap}
        
        # Title
        tk.Label(scan_win, text="📷 Point barcode at camera",
                font=('Arial', 14, 'bold'), fg='white', bg='#2c3e50').pack(pady=10)
        
        # Video display
        video_label = tk.Label(scan_win, bg='black')
        video_label.pack(padx=10, pady=5)
        
        # Status label
        status_label = tk.Label(scan_win, text="Scanning...",
                               font=('Arial', 11), fg='#3498db', bg='#2c3e50')
        status_label.pack(pady=5)
        
        # Camera selection (in case of multiple cameras)
        cam_frame = tk.Frame(scan_win, bg='#2c3e50')
        cam_frame.pack(pady=5)
        tk.Label(cam_frame, text="Camera:", fg='white', bg='#2c3e50').pack(side='left')
        
        cam_var = tk.IntVar(master=self.root, value=0)
        def switch_camera():
            try:
                new_cam = cam_var.get()
                scan_state['cap'].release()
                new_cap = cv2.VideoCapture(new_cam)
                if new_cap.isOpened():
                    scan_state['cap'] = new_cap
                    status_label.config(text=f"Switched to camera {new_cam}")
                else:
                    # Revert
                    scan_state['cap'] = cv2.VideoCapture(0)
                    cam_var.set(0)
                    status_label.config(text=f"Camera {new_cam} not available", fg='#e74c3c')
            except Exception as e:
                status_label.config(text=f"Error: {e}", fg='#e74c3c')
        
        tk.Spinbox(cam_frame, from_=0, to=5, width=3, textvariable=cam_var,
                  command=switch_camera).pack(side='left', padx=5)
        
        def close_scanner():
            scan_state['running'] = False
            try:
                scan_state['cap'].release()
            except:
                pass
            try:
                cv2.destroyAllWindows()
            except:
                pass
            scan_win.destroy()
        
        # Close button
        tk.Button(scan_win, text="Cancel", command=close_scanner,
                 bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'),
                 width=12).pack(pady=10)
        
        scan_win.protocol("WM_DELETE_WINDOW", close_scanner)
        
        # Import PIL for image conversion
        from PIL import Image as PILImage, ImageTk
        
        def scan_frame():
            if not scan_state['running']:
                return
            
            try:
                ret, frame = scan_state['cap'].read()
                if not ret:
                    scan_win.after(100, scan_frame)
                    return
                
                # Resize for display and processing
                display_frame = cv2.resize(frame, (640, 480))
                
                # Try to decode barcodes
                gray = cv2.cvtColor(display_frame, cv2.COLOR_BGR2GRAY)
                
                barcode_data = None
                barcode_rect = None
                
                if PYZBAR_AVAILABLE:
                    # Preferred: pyzbar (supports more formats)
                    barcodes = pyzbar.decode(gray)
                    if barcodes:
                        b = barcodes[0]
                        barcode_data = b.data.decode('utf-8')
                        barcode_rect = b.rect
                else:
                    # Fallback: OpenCV's built-in barcode detector (works with common formats)
                    try:
                        bd = cv2.barcode.BarcodeDetector() if hasattr(cv2, 'barcode') else None
                        if bd:
                            ok, decoded_info, decoded_type, points = bd.detectAndDecode(gray)
                            if ok and decoded_info and decoded_info[0]:
                                barcode_data = decoded_info[0]
                                if points is not None and len(points) > 0:
                                    x_min = int(min(p[0] for p in points[0]))
                                    y_min = int(min(p[1] for p in points[0]))
                                    x_max = int(max(p[0] for p in points[0]))
                                    y_max = int(max(p[1] for p in points[0]))
                                    barcode_rect = (x_min, y_min, x_max - x_min, y_max - y_min)
                        else:
                            # Even simpler fallback: QRCode detector (only QR codes)
                            qr = cv2.QRCodeDetector()
                            data, points, _ = qr.detectAndDecode(gray)
                            if data:
                                barcode_data = data
                                if points is not None and len(points) > 0:
                                    x_min = int(min(p[0] for p in points[0]))
                                    y_min = int(min(p[1] for p in points[0]))
                                    x_max = int(max(p[0] for p in points[0]))
                                    y_max = int(max(p[1] for p in points[0]))
                                    barcode_rect = (x_min, y_min, x_max - x_min, y_max - y_min)
                    except Exception:
                        pass
                
                if barcode_data:
                    # Found a barcode!
                    if barcode_rect:
                        (x, y, w, h) = barcode_rect
                        cv2.rectangle(display_frame, (x, y), (x + w, y + h), (0, 255, 0), 3)
                        cv2.putText(display_frame, barcode_data, (x, max(y - 10, 20)),
                                   cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
                    
                    status_label.config(text=f"✓ Found: {barcode_data}", fg='#27ae60')
                    scan_state['found'] = barcode_data
                    
                    # Update display one last time
                    rgb = cv2.cvtColor(display_frame, cv2.COLOR_BGR2RGB)
                    img = PILImage.fromarray(rgb)
                    imgtk = ImageTk.PhotoImage(image=img)
                    video_label.imgtk = imgtk
                    video_label.config(image=imgtk)
                    
                    # Wait a moment then close and show status
                    scan_win.after(800, lambda: (close_scanner(), self.show_scan_status_popup(barcode_data)))
                    return
                
                # Convert frame for tkinter display
                rgb = cv2.cvtColor(display_frame, cv2.COLOR_BGR2RGB)
                img = PILImage.fromarray(rgb)
                imgtk = ImageTk.PhotoImage(image=img)
                video_label.imgtk = imgtk
                video_label.config(image=imgtk)
                
            except Exception as e:
                print(f"Webcam scan error: {e}")
            
            # Schedule next frame
            if scan_state['running']:
                scan_win.after(30, scan_frame)  # ~30fps
        
        # Start scanning
        scan_frame()
    
    def manual_barcode_lookup(self):
        """Manually enter an estimate number to look up status"""
        from tkinter import simpledialog
        estimate_num = simpledialog.askstring("Barcode Lookup", 
                                              "Enter or scan estimate number:",
                                              parent=self.root)
        if estimate_num and estimate_num.strip():
            self.show_scan_status_popup(estimate_num.strip())
    
    def setup_barcode_scanner_listener(self):
        """Setup global key listener to detect barcode scanner input"""
        import time
        
        # Debug log file to diagnose scanner issues
        import os, sys
        try:
            base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        except:
            base_dir = os.getcwd()
        self._scan_debug_log = os.path.join(base_dir, 'scanner_debug.log')
        
        def debug_log(msg):
            try:
                with open(self._scan_debug_log, 'a') as f:
                    f.write(f"{time.strftime('%H:%M:%S')} - {msg}\n")
            except:
                pass
        
        debug_log("=== Scanner listener started ===")
        
        def on_key_press(event):
            if self._scan_popup_open:
                return
            
            current_time = time.time() * 1000
            time_diff = current_time - self._scan_last_time
            
            # Get character
            char = ''
            if event.char and event.char.isprintable():
                char = event.char
            elif len(event.keysym) == 1 and event.keysym.isprintable():
                char = event.keysym
            
            widget_type = type(event.widget).__name__
            
            # Handle Enter
            if event.keysym in ('Return', 'KP_Enter'):
                scanned = self._scan_buffer.strip()
                debug_log(f"ENTER pressed | buffer='{scanned}' | len={len(scanned)} | widget={widget_type}")
                self._scan_buffer = ""
                self._scan_last_time = 0
                if len(scanned) >= 3:
                    debug_log(f"TRIGGERING POPUP for '{scanned}'")
                    self.root.after(50, lambda s=scanned: self.show_scan_status_popup(s))
                return
            
            if not char:
                debug_log(f"Non-printable key: keysym={event.keysym} keycode={event.keycode} char='{event.char}' widget={widget_type}")
                return
            
            # Reset buffer if too much time between keys
            if time_diff > 500:
                self._scan_buffer = ""
            
            self._scan_last_time = current_time
            self._scan_buffer += char
            
            if len(self._scan_buffer) <= 20:
                debug_log(f"CHAR: '{char}' | buffer='{self._scan_buffer}' | gap={time_diff:.0f}ms | widget={widget_type}")
        
        self.root.bind_all('<KeyPress>', on_key_press)
        debug_log("bind_all '<KeyPress>' registered")
    
    def show_scan_status_popup(self, estimate_num):
        """Show a popup with the current status of a scanned machine with interactive checkboxes"""
        self._scan_popup_open = True
        
        # Search in active machines
        machine = None
        location = None
        for m in self.active_machines:
            if m['estimate_number'] == estimate_num:
                machine = m
                location = "Active Repairs"
                break
        
        # Search in shipped machines if not found
        if not machine:
            for m in self.shipped_machines:
                if m['estimate_number'] == estimate_num:
                    machine = m
                    location = "Shipped / Past Repairs"
                    break
        
        # Create popup window
        popup = tk.Toplevel(self.root)
        popup.title(f"Repair Status - {estimate_num}")
        popup.configure(bg='#2c3e50')
        popup.transient(self.root)
        popup.grab_set()
        
        # Center the popup
        popup.geometry("520x580+%d+%d" % (self.root.winfo_rootx() + 340, self.root.winfo_rooty() + 60))
        
        def close_popup(event=None):
            self._scan_popup_open = False
            popup.destroy()
        
        popup.protocol("WM_DELETE_WINDOW", close_popup)
        popup.bind('<Escape>', close_popup)
        
        if not machine:
            tk.Label(popup, text="❌ Not Found", font=('Arial', 18, 'bold'),
                    fg='#e74c3c', bg='#2c3e50').pack(pady=20)
            tk.Label(popup, text=f"No machine found with Estimate #:\n{estimate_num}",
                    font=('Arial', 14), fg='#ecf0f1', bg='#2c3e50').pack(pady=10)
            tk.Button(popup, text="OK", command=close_popup, bg='#3498db', fg='white',
                     font=('Arial', 12, 'bold'), width=10).pack(pady=20)
            popup.focus_set()
            return
        
        # Header
        tk.Label(popup, text="📋 Repair Status", font=('Arial', 18, 'bold'),
                fg='white', bg='#2c3e50').pack(pady=(15, 5))
        
        # Status location badge
        badge_color = '#27ae60' if location == "Active Repairs" else '#7f8c8d'
        tk.Label(popup, text=f"  {location}  ", font=('Arial', 11, 'bold'),
                fg='white', bg=badge_color).pack(pady=(0, 10))
        
        # Machine info frame
        info_frame = tk.Frame(popup, bg='#34495e', padx=20, pady=10)
        info_frame.pack(fill='x', padx=20, pady=(5, 0))
        
        def add_info_row(parent, label, value, row):
            tk.Label(parent, text=label, font=('Arial', 11, 'bold'), fg='#bdc3c7',
                    bg='#34495e', anchor='e', width=16).grid(row=row, column=0, sticky='e', pady=2)
            tk.Label(parent, text=value, font=('Arial', 11), fg='#ecf0f1',
                    bg='#34495e', anchor='w').grid(row=row, column=1, sticky='w', padx=(10, 0), pady=2)
        
        add_info_row(info_frame, "Estimate #:", machine['estimate_number'], 0)
        add_info_row(info_frame, "Customer:", machine['customer_name'], 1)
        add_info_row(info_frame, "Make / Model:", f"{machine['make']} {machine['model']}", 2)
        add_info_row(info_frame, "Serial #:", machine['serial_number'], 3)
        add_info_row(info_frame, "Arrival Date:", machine['arrival_date'], 4)
        if location == "Shipped / Past Repairs":
            add_info_row(info_frame, "Ship Date:", machine.get('ship_date', 'N/A'), 5)
        
        # Status checkboxes frame
        status_frame = tk.LabelFrame(popup, text="  Status  ", font=('Arial', 12, 'bold'),
                                     fg='#f39c12', bg='#34495e', padx=15, pady=10)
        status_frame.pack(fill='x', padx=20, pady=10)
        
        is_active = location == "Active Repairs"
        
        # Create checkbox variables
        est_sent_var = tk.BooleanVar(value=machine.get('estimate_sent', False))
        est_complete_var = tk.BooleanVar(value=machine.get('estimate_complete', False))
        parts_ordered_var = tk.BooleanVar(value=machine.get('parts_ordered', False))
        po_received_var = tk.BooleanVar(value=machine.get('po_received', False))
        repair_complete_var = tk.BooleanVar(value=machine.get('repair_complete', False))
        
        checkbox_style = {'font': ('Arial', 11), 'bg': '#34495e', 'fg': '#ecf0f1',
                         'selectcolor': '#2c3e50', 'activebackground': '#34495e',
                         'activeforeground': '#ecf0f1'}
        
        cb_est_sent = tk.Checkbutton(status_frame, text="Estimate Sent", variable=est_sent_var,
                                     state='normal' if is_active else 'disabled', **checkbox_style)
        cb_est_sent.grid(row=0, column=0, sticky='w', pady=4)
        
        cb_est_complete = tk.Checkbutton(status_frame, text="Estimate Complete", variable=est_complete_var,
                                         state='normal' if is_active else 'disabled', **checkbox_style)
        cb_est_complete.grid(row=1, column=0, sticky='w', pady=4)
        
        cb_parts_ordered = tk.Checkbutton(status_frame, text="Parts Ordered", variable=parts_ordered_var,
                                          state='normal' if is_active else 'disabled', **checkbox_style)
        cb_parts_ordered.grid(row=2, column=0, sticky='w', pady=4)
        
        cb_po_received = tk.Checkbutton(status_frame, text="P.O. Received", variable=po_received_var,
                                        state='normal' if is_active else 'disabled', **checkbox_style)
        cb_po_received.grid(row=3, column=0, sticky='w', pady=4)
        
        # Show PO number if received
        po_num_text = f"  (P.O. #: {machine.get('po_number', '')})" if machine.get('po_number') else ""
        if po_num_text:
            tk.Label(status_frame, text=po_num_text, font=('Arial', 10), fg='#bdc3c7',
                    bg='#34495e').grid(row=3, column=1, sticky='w', padx=(5, 0))
        
        cb_repair_complete = tk.Checkbutton(status_frame, text="Repair Complete", variable=repair_complete_var,
                                            state='normal' if is_active else 'disabled', **checkbox_style)
        cb_repair_complete.grid(row=4, column=0, sticky='w', pady=4)
        
        # PDF status (not a checkbox, just info)
        has_pdf = bool(machine.get('pdf_attachment'))
        pdf_color = '#27ae60' if has_pdf else '#95a5a6'
        tk.Label(status_frame, text=f"PDF Attached: {'Yes' if has_pdf else 'No'}",
                font=('Arial', 11), fg=pdf_color, bg='#34495e').grid(row=5, column=0, sticky='w', pady=(8, 2))
        
        # Buttons
        btn_frame = tk.Frame(popup, bg='#2c3e50')
        btn_frame.pack(pady=15)
        
        def save_and_close():
            if is_active:
                machine['estimate_sent'] = est_sent_var.get()
                machine['estimate_complete'] = est_complete_var.get()
                machine['parts_ordered'] = parts_ordered_var.get()
                machine['po_received'] = po_received_var.get()
                if not po_received_var.get():
                    machine['po_number'] = ''
                machine['repair_complete'] = repair_complete_var.get()
                self.refresh_active_table()
                self.save_data()
            close_popup()
        
        def open_so_and_close():
            save_and_close() if is_active else close_popup()
            self.open_service_order(machine)
        
        if is_active:
            tk.Button(btn_frame, text="Save Changes", command=save_and_close, bg='#27ae60', fg='white',
                     font=('Arial', 12, 'bold'), width=14).pack(side='left', padx=5)
            tk.Button(btn_frame, text="Open Service Order", command=open_so_and_close, bg='#3498db', fg='white',
                     font=('Arial', 12, 'bold'), width=16).pack(side='left', padx=5)
            tk.Button(btn_frame, text="Cancel", command=close_popup, bg='#95a5a6', fg='white',
                     font=('Arial', 12, 'bold'), width=10).pack(side='left', padx=5)
        else:
            tk.Button(btn_frame, text="Open Service Order", command=open_so_and_close, bg='#3498db', fg='white',
                     font=('Arial', 12, 'bold'), width=16).pack(side='left', padx=5)
            tk.Button(btn_frame, text="OK", command=close_popup, bg='#95a5a6', fg='white',
                     font=('Arial', 12, 'bold'), width=10).pack(side='left', padx=5)
        
        popup.focus_set()
    
    def is_frozen(self):
        """True when running as a PyInstaller EXE rather than a .py script."""
        return bool(getattr(sys, 'frozen', False))

    def get_app_path(self):
        """Full path to the file that would be replaced by an update."""
        if self.is_frozen():
            return os.path.abspath(sys.executable)
        return os.path.abspath(sys.argv[0] if sys.argv and sys.argv[0] else __file__)

    def get_app_base_name(self):
        """'RepairManagementSuite' from RepairManagementSuite.exe"""
        return os.path.splitext(os.path.basename(self.get_app_path()))[0]


    def get_update_staging_dir(self):
        """Local scratch folder for downloaded updates (never on the share)."""
        base = os.environ.get('LOCALAPPDATA') or os.path.expanduser("~")
        path = os.path.join(base, "RepairManagementSuite", "updates")
        try:
            os.makedirs(path, exist_ok=True)
        except Exception:
            path = os.path.join(os.path.expanduser("~"), ".repair_management_suite", "updates")
            os.makedirs(path, exist_ok=True)
        return path

    # ---------------- discovery ----------------








    def _confirm_update_launch(self):
        """Run at startup. Clears the probation flag and reports the result.

        The helper script creates a probation flag before starting the new
        build. Deleting it here is how the new build says 'I started fine'.
        """
        try:
            staging = self.get_update_staging_dir()
        except Exception:
            return

        flag = os.path.join(staging, "update_probation.flag")
        pending = os.path.join(staging, "pending_version.json")

        try:
            if os.path.exists(flag):
                os.remove(flag)
        except Exception:
            pass

        if not os.path.exists(pending):
            return

        try:
            with open(pending, 'r') as f:
                data = json.load(f)
            expected = data.get('version', '')
            os.remove(pending)
        except Exception:
            return

        if not expected:
            return

        if version_key(expected) == version_key(APP_VERSION):
            self._update_result_message = f"✅ Updated to version {APP_VERSION}"
            self._update_result_color = '#27ae60'
        else:
            self._update_result_message = (
                f"⚠ Update to {expected} did not take - still running {APP_VERSION}")
            self._update_result_color = '#e74c3c'

    def _show_update_result(self):
        """Put the post-update message in the status bar once the GUI exists."""
        msg = getattr(self, '_update_result_message', None)
        if not msg:
            return
        try:
            if getattr(self, 'status_label', None):
                self.status_label.config(
                    text=msg, fg=getattr(self, '_update_result_color', '#27ae60'))
        except Exception:
            pass
        self._update_result_message = None

    # ---------------- applying on exit ----------------

    def apply_update_on_exit(self):
        """Ask the tech what to do, then launch the helper. Called from on_closing."""
        info = getattr(self, '_pending_update', None)
        if not info or not info.get('staged'):
            return
        if not os.path.exists(info['staged']):
            return

        # A program installed somewhere the tech can't write (Program Files
        # without admin) can never update itself - say so instead of letting
        # the helper spin for 90 seconds and give up silently.
        target = self.get_app_path()
        install_dir = os.path.dirname(target)
        if install_dir and not os.access(install_dir, os.W_OK):
            messagebox.showwarning(
                "Update Cannot Be Installed",
                f"Version {info['version_str']} was downloaded, but this program is "
                f"installed in a folder you don't have permission to write to:\n\n"
                f"{install_dir}\n\n"
                "Move the program somewhere like your Desktop or a user folder, or "
                "run it as administrator, and the update will install on the next exit.")
            return

        choice = self._ask_update_choice(info)
        if choice == 'skip':
            return

        ok, err = self._launch_updater(info, relaunch=(choice == 'restart'))
        if not ok:
            messagebox.showerror(
                "Update Failed to Start",
                f"The update could not be installed:\n\n{err}\n\n"
                "The program will close normally and try again next time.")

    def _ask_update_choice(self, info):
        """Three-way exit prompt. Returns 'restart', 'close' or 'skip'."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Update Available")
        dlg.configure(bg='#2c3e50')
        dlg.resizable(False, False)
        dlg.transient(self.root)

        result = {'choice': 'skip'}

        tk.Label(dlg, text=f"Update to version {info['version_str']}",
                 font=('Arial', 14, 'bold'), fg='#f39c12',
                 bg='#2c3e50').pack(anchor='w', padx=20, pady=(16, 4))
        tk.Label(dlg, text=f"You are running version {APP_VERSION}.",
                 font=('Arial', 10), fg='#bdc3c7', bg='#2c3e50').pack(anchor='w', padx=20)
        tk.Label(dlg,
                 text="The new version has been downloaded and verified.\n"
                      "Your data and settings are not affected.",
                 font=('Arial', 10), fg='#ecf0f1', bg='#2c3e50',
                 justify='left', wraplength=380).pack(anchor='w', padx=20, pady=(10, 14))

        btns = tk.Frame(dlg, bg='#2c3e50')
        btns.pack(fill='x', padx=20, pady=(0, 16))

        def pick(value):
            result['choice'] = value
            dlg.destroy()

        tk.Button(btns, text="Install and Restart", command=lambda: pick('restart'),
                  bg='#27ae60', fg='white', font=('Arial', 10, 'bold'),
                  width=20).pack(fill='x', pady=3)
        tk.Button(btns, text="Install and Stay Closed", command=lambda: pick('close'),
                  bg='#2980b9', fg='white', font=('Arial', 10, 'bold'),
                  width=20).pack(fill='x', pady=3)
        tk.Button(btns, text="Skip This Time", command=lambda: pick('skip'),
                  bg='#7f8c8d', fg='white', font=('Arial', 10, 'bold'),
                  width=20).pack(fill='x', pady=3)

        dlg.protocol("WM_DELETE_WINDOW", lambda: pick('skip'))
        dlg.update_idletasks()
        try:
            px = self.root.winfo_rootx() + (self.root.winfo_width() // 2) - (dlg.winfo_width() // 2)
            py = self.root.winfo_rooty() + 150
            dlg.geometry(f"+{max(px, 0)}+{max(py, 0)}")
        except Exception:
            pass
        dlg.grab_set()
        self.root.wait_window(dlg)
        return result['choice']

    def _launch_updater(self, info, relaunch=True):
        """Write the helper script and start it detached. Returns (ok, error)."""
        staging = self.get_update_staging_dir()
        target = self.get_app_path()
        staged = info['staged']
        backup = target + ".bak"
        flag = os.path.join(staging, "update_probation.flag")
        log = os.path.join(staging, "update_log.txt")

        # Record what we're expecting so the next launch can report the outcome
        try:
            with open(os.path.join(staging, "pending_version.json"), 'w') as f:
                json.dump({'version': info['version_str'],
                           'from': APP_VERSION,
                           'at': datetime.now().isoformat()}, f)
        except Exception:
            pass

        if relaunch:
            try:
                with open(flag, 'w') as f:
                    f.write(info['version_str'])
            except Exception:
                pass

        try:
            if os.name == 'nt':
                script = self._write_windows_updater(
                    staging, target, staged, backup, flag, log,
                    info['size'], relaunch, info['version_str'])
                creationflags = 0x00000008 | 0x08000000  # DETACHED | NO_WINDOW
                subprocess.Popen(['cmd.exe', '/c', script],
                                 creationflags=creationflags,
                                 close_fds=True)
            else:
                script = self._write_posix_updater(
                    staging, target, staged, backup, flag, log,
                    info['size'], relaunch)
                subprocess.Popen(['/bin/sh', script],
                                 start_new_session=True,
                                 close_fds=True)
            return True, None
        except Exception as e:
            try:
                if os.path.exists(flag):
                    os.remove(flag)
            except Exception:
                pass
            return False, str(e)

    def _write_windows_updater(self, staging, target, staged, backup, flag, log,
                               expected_size, relaunch, version_str):
        """Write the .bat that performs the swap after this process exits."""
        exe_name = target.replace('\\', '/').rstrip('/').split('/')[-1]
        launch_cmd = f'start "" "{target}"'
        if not self.is_frozen():
            py = sys.executable or 'python'
            launch_cmd = f'start "" "{py}" "{target}"'

        script = f"""@echo off
setlocal enableextensions
set "TARGET={target}"
set "STAGED={staged}"
set "BACKUP={backup}"
set "FLAG={flag}"
set "LOG={log}"
set "EXENAME={exe_name}"
set "EXPECTSIZE={expected_size}"
set "RELAUNCH={'1' if relaunch else '0'}"

echo ================================================= >>"%LOG%"
echo [%DATE% %TIME%] update to {version_str} starting >>"%LOG%"

rem --- Wait for the app to release its own file. Moving it only
rem --- succeeds once the process has fully exited.
set /a TRIES=0
:trymove
if exist "%BACKUP%" del /f /q "%BACKUP%" >nul 2>&1
move /y "%TARGET%" "%BACKUP%" >nul 2>&1
if not errorlevel 1 goto moved
set /a TRIES+=1
if %TRIES% GEQ 90 goto locked
ping -n 2 127.0.0.1 >nul
goto trymove

:locked
echo [ERROR] could not replace the program after 90s. >>"%LOG%"
echo         Either it is still running, or this folder is not writable. >>"%LOG%"
goto cleanup

:moved
echo [OK] backed up old version >>"%LOG%"

copy /y "%STAGED%" "%TARGET%" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] copy of new version failed >>"%LOG%"
    goto rollback
)

rem --- Verify the copy landed intact
for %%A in ("%TARGET%") do set "GOTSIZE=%%~zA"
if not "%GOTSIZE%"=="%EXPECTSIZE%" (
    echo [ERROR] size mismatch: got %GOTSIZE% expected %EXPECTSIZE% >>"%LOG%"
    goto rollback
)

echo [OK] installed version {version_str} >>"%LOG%"

if not "%RELAUNCH%"=="1" goto done

rem --- Probation: the new build deletes the flag when it starts cleanly
{launch_cmd}
set /a WAIT=0
:probation
if not exist "%FLAG%" goto launched
set /a WAIT+=1
if %WAIT% GEQ 40 goto badlaunch
ping -n 3 127.0.0.1 >nul
goto probation

:launched
echo [OK] new version started successfully >>"%LOG%"
goto done

:badlaunch
echo [ERROR] new version did not start within 80s - rolling back >>"%LOG%"
taskkill /F /IM "%EXENAME%" >nul 2>&1
ping -n 3 127.0.0.1 >nul

:rollback
if exist "%BACKUP%" (
    if exist "%TARGET%" del /f /q "%TARGET%" >nul 2>&1
    move /y "%BACKUP%" "%TARGET%" >nul 2>&1
    echo [OK] rolled back to previous version >>"%LOG%"
    if "%RELAUNCH%"=="1" {launch_cmd}
) else (
    echo [ERROR] no backup to roll back to >>"%LOG%"
)
goto cleanup

:done
if exist "%STAGED%" del /f /q "%STAGED%" >nul 2>&1

:cleanup
if exist "%FLAG%" del /f /q "%FLAG%" >nul 2>&1
echo [%DATE% %TIME%] updater finished >>"%LOG%"
(goto) 2>nul & del /f /q "%~f0"
"""

        path = os.path.join(staging, "apply_update.bat")
        # cmd.exe reads batch files in the console codepage, not UTF-8, so a
        # UTF-8 file breaks on paths containing non-ASCII characters.
        try:
            with open(path, 'w', encoding='mbcs') as f:
                f.write(script)
        except (LookupError, UnicodeEncodeError):
            with open(path, 'w', encoding='utf-8') as f:
                f.write(script)
        return path

    def _write_posix_updater(self, staging, target, staged, backup, flag, log,
                             expected_size, relaunch):
        """Equivalent helper for macOS / Linux."""
        if self.is_frozen():
            launch = f'"{target}" &'
        else:
            launch = f'"{sys.executable or "python3"}" "{target}" &'

        script = f"""#!/bin/sh
TARGET="{target}"
STAGED="{staged}"
BACKUP="{backup}"
FLAG="{flag}"
LOG="{log}"
EXPECTSIZE={expected_size}
RELAUNCH={1 if relaunch else 0}

echo "[$(date)] update starting" >>"$LOG"

# Wait for the old process to exit
i=0
while [ $i -lt 90 ]; do
    if mv -f "$TARGET" "$BACKUP" 2>/dev/null; then
        break
    fi
    i=$((i+1))
    sleep 1
done

if [ ! -f "$BACKUP" ]; then
    echo "[ERROR] could not back up old version" >>"$LOG"
    rm -f "$FLAG"
    exit 1
fi

if ! cp -f "$STAGED" "$TARGET"; then
    echo "[ERROR] copy failed - rolling back" >>"$LOG"
    mv -f "$BACKUP" "$TARGET"
    rm -f "$FLAG"
    exit 1
fi
chmod +x "$TARGET" 2>/dev/null

GOTSIZE=$(wc -c <"$TARGET" | tr -d ' ')
if [ "$GOTSIZE" != "$EXPECTSIZE" ]; then
    echo "[ERROR] size mismatch - rolling back" >>"$LOG"
    rm -f "$TARGET"
    mv -f "$BACKUP" "$TARGET"
    rm -f "$FLAG"
    exit 1
fi

echo "[OK] installed" >>"$LOG"

if [ "$RELAUNCH" = "1" ]; then
    {launch}
    i=0
    while [ $i -lt 40 ]; do
        [ ! -f "$FLAG" ] && break
        i=$((i+1))
        sleep 2
    done
    if [ -f "$FLAG" ]; then
        echo "[ERROR] new version did not start - rolling back" >>"$LOG"
        rm -f "$TARGET"
        mv -f "$BACKUP" "$TARGET"
        rm -f "$FLAG"
        {launch}
        exit 1
    fi
fi

rm -f "$STAGED"
rm -f "$FLAG"
echo "[$(date)] updater finished" >>"$LOG"
rm -f "$0"
"""
        path = os.path.join(staging, "apply_update.sh")
        with open(path, 'w', encoding='utf-8') as f:
            f.write(script)
        try:
            os.chmod(path, 0o755)
        except Exception:
            pass
        return path

    # ==================================================================
    # ==============  AUTO-UPDATE FROM GITHUB RELEASES  ================
    # ==================================================================
    #
    #   1. You bump APP_VERSION, commit, and push a tag (v1.3.0).
    #   2. GitHub Actions builds the Windows EXE and publishes it as a
    #      release asset.
    #   3. Every copy checks the releases API on boot, in the background,
    #      and quietly downloads a newer build to a local staging folder.
    #   4. On exit the tech is asked whether to install it.
    #   5. If the new build never starts, the old one is put back.
    # ==================================================================

    def _set_update_status(self, code, message):
        """Remember why the last check did or didn't find an update, so the
        Check for Updates button can say so instead of 'up to date'."""
        self._update_check_status = (code, message)
        print(f"Update check [{code}]: {message}")

    def fetch_latest_release(self, timeout=12):
        """Ask GitHub for the newest published release. None on failure,
        with the reason stored by _set_update_status."""
        req = urllib.request.Request(
            GITHUB_API_LATEST,
            headers={
                'Accept': 'application/vnd.github+json',
                # GitHub rejects requests without a User-Agent
                'User-Agent': f'RepairManagementSuite/{APP_VERSION}',
            })
        try:
            with urllib.request.urlopen(req, timeout=timeout,
                                        context=https_context()) as resp:
                if resp.status != 200:
                    self._set_update_status(
                        'error', f"GitHub answered with an unexpected status (HTTP {resp.status}).")
                    return None
                return json.loads(resp.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                self._set_update_status(
                    'error', f"No releases have been published at "
                             f"github.com/{GITHUB_OWNER}/{GITHUB_REPO} yet.")
            elif e.code in (403, 429) and str(e.headers.get('X-RateLimit-Remaining', '')) == '0':
                when = ''
                try:
                    reset = int(e.headers.get('X-RateLimit-Reset', '0'))
                    when = " It resets at " + datetime.fromtimestamp(reset).strftime('%I:%M %p').lstrip('0') + "."
                except Exception:
                    pass
                self._set_update_status(
                    'error', "GitHub's limit on update checks from this network has been "
                             f"reached (60 per hour).{when}")
            else:
                self._set_update_status(
                    'error', f"GitHub refused the request (HTTP {e.code} {e.reason}).")
            return None
        except Exception as e:
            # No internet, DNS blocked, proxy, firewall, certificate problems
            self._set_update_status('error', describe_network_error(e))
            return None

    def find_available_update(self):
        """Return details of a newer release, or None.

        {'version_str', 'tag', 'url', 'size', 'filename', 'notes'}
        The reason for a None is in self._update_check_status.
        """
        self._update_check_status = ('error', "The update check didn't finish.")
        release = self.fetch_latest_release()
        if not release:
            return None
        if release.get('draft') or release.get('prerelease'):
            self._set_update_status('current', "The newest release is a draft or pre-release.")
            return None

        tag = str(release.get('tag_name', '') or '')
        ver = parse_version(tag.lstrip('vV'))
        if not ver:
            self._set_update_status(
                'error', f"The newest release on GitHub is tagged '{tag}', which isn't a "
                         "version number. Tags should look like v1.3.4.")
            return None

        if version_key(ver) <= version_key(APP_VERSION):
            self._set_update_status(
                'current', f"You're on the newest version. The latest release on GitHub "
                           f"is {format_version(ver)}.")
            return None

        # Find the Windows EXE attached to the release
        want = '.exe' if self.is_frozen() else '.py'
        asset = None
        for a in release.get('assets', []):
            name = str(a.get('name', ''))
            if name.lower().endswith(want) and a.get('browser_download_url'):
                asset = a
                break
        if not asset:
            self._set_update_status(
                'building', f"Version {format_version(ver)} has been published, but its "
                            "program file isn't attached yet. GitHub is usually still "
                            "building it - try again in about 5 minutes.")
            return None
        self._set_update_status('available', f"Version {format_version(ver)} is available.")

        return {
            'version_str': format_version(ver),
            'tag': tag,
            'url': asset['browser_download_url'],
            'size': int(asset.get('size') or 0),
            'filename': asset.get('name', f'RepairManagementSuite_{tag}{want}'),
            'notes': (release.get('body') or '').strip(),
        }

    def _download_update(self, info, progress=None):
        """Download the release asset to the local staging folder.

        Returns the downloaded path, or None on failure.
        """
        staging = self.get_update_staging_dir()
        self._download_error = ''

        # Clear anything left from a previous run
        try:
            for old in os.listdir(staging):
                if old.lower().endswith(('.exe', '.py', '.part')):
                    try:
                        os.remove(os.path.join(staging, old))
                    except Exception:
                        pass
        except Exception:
            pass

        dest = os.path.join(staging, info['filename'])
        part = dest + '.part'

        req = urllib.request.Request(
            info['url'],
            headers={'User-Agent': f'RepairManagementSuite/{APP_VERSION}',
                     'Accept': 'application/octet-stream'})
        try:
            with urllib.request.urlopen(req, timeout=60, context=https_context()) as resp, \
                    open(part, 'wb') as f:
                total = info.get('size') or int(resp.headers.get('Content-Length') or 0)
                got = 0
                while True:
                    chunk = resp.read(256 * 1024)
                    if not chunk:
                        break
                    f.write(chunk)
                    got += len(chunk)
                    if progress and total:
                        try:
                            progress(got, total)
                        except Exception:
                            pass
        except Exception as e:
            print(f"Update download failed: {e}")
            self._download_error = describe_network_error(e)
            try:
                if os.path.exists(part):
                    os.remove(part)
            except Exception:
                pass
            return None

        # A truncated download must never be installed
        try:
            actual = os.path.getsize(part)
            if info.get('size') and actual != info['size']:
                print(f"Update download incomplete: {actual} of {info['size']} bytes")
                self._download_error = (f"The download stopped early ({actual:,} of "
                                        f"{info['size']:,} bytes), so it was discarded.")
                os.remove(part)
                return None
            if actual <= 0:
                os.remove(part)
                return None
            if os.path.exists(dest):
                os.remove(dest)
            os.rename(part, dest)
        except Exception as e:
            print(f"Update verification failed: {e}")
            self._download_error = f"The downloaded file couldn't be saved.\n\nDetail: {e}"
            return None

        return dest

    # ---------------- startup check ----------------

    def check_for_updates_async(self):
        """Check and download in the background so boot is never delayed."""
        def worker():
            try:
                info = self.find_available_update()
                if not info:
                    code, _msg = getattr(self, '_update_check_status', ('', ''))
                    banner = None
                    if code == 'error':
                        banner = ("⚠ Couldn't check for updates - click Check for Updates "
                                  "to see why", '#e67e22')
                    elif code == 'building':
                        banner = ("⏳ A new version is still being built - check again "
                                  "in a few minutes", '#95a5a6')
                    if banner:
                        try:
                            self.root.after(0, lambda: self._set_update_banner(*banner))
                        except Exception:
                            pass
                    return
                try:
                    self.root.after(0, lambda: self._set_update_banner(
                        f"⬇ Downloading update {info['version_str']}...", '#3498db'))
                except Exception:
                    pass

                path = self._download_update(info)
                if not path:
                    try:
                        self.root.after(0, lambda: self._set_update_banner(
                            "⚠ Update download failed - will retry next launch", '#e74c3c'))
                    except Exception:
                        pass
                    return

                info['staged'] = path
                try:
                    self.root.after(0, lambda: self._on_update_ready(info))
                except Exception:
                    pass
            except Exception as e:
                print(f"Update check failed: {e}")

        try:
            threading.Thread(target=worker, daemon=True, name="update-check").start()
        except Exception as e:
            print(f"Could not start update check: {e}")

    def _set_update_banner(self, text, color='#f39c12'):
        try:
            if getattr(self, 'update_label', None):
                self.update_label.config(text=text, fg=color)
                self.update_label.pack(side='left', padx=15)
        except Exception:
            pass

    def _on_update_ready(self, info):
        """A newer build is downloaded and waiting to install on exit."""
        self._pending_update = info
        self._set_update_banner(
            f"⬆ Update {info['version_str']} ready - installs when you close", '#f39c12')
        print(f"Update {info['version_str']} staged at {info['staged']}")

    def check_for_updates_manual(self):
        """'Check for Updates' button."""
        if getattr(self, '_pending_update', None):
            info = self._pending_update
            messagebox.showinfo(
                "Update Ready",
                f"Version {info['version_str']} has already been downloaded.\n\n"
                f"You are running {APP_VERSION}.\n\n"
                "It will install when you close the program.")
            return

        if getattr(self, 'status_label', None):
            self.status_label.config(text="Checking GitHub for updates...", fg='#f39c12')
        self.root.update_idletasks()

        info = self.find_available_update()
        if not info:
            code, message = getattr(self, '_update_check_status',
                                    ('error', "The update check didn't finish."))
            status = getattr(self, 'status_label', None)
            if code == 'current':
                messagebox.showinfo("Up To Date",
                                    f"You are running version {APP_VERSION}.\n\n{message}")
                if status:
                    status.config(text=f"Up to date (v{APP_VERSION})", fg='#27ae60')
            elif code == 'building':
                messagebox.showinfo("Update Not Ready Yet",
                                    f"You are running version {APP_VERSION}.\n\n{message}")
                if status:
                    status.config(text="Update is still being built - try again shortly",
                                  fg='#f39c12')
            else:
                messagebox.showwarning(
                    "Couldn't Check for Updates",
                    f"You are running version {APP_VERSION}.\n\n"
                    "The update check failed, so it's not known whether a newer "
                    f"version exists.\n\n{message}\n\n"
                    f"You can also download the newest version yourself from:\n"
                    f"{GITHUB_RELEASES_PAGE}")
                if status:
                    status.config(text="Update check failed", fg='#e74c3c')
            return

        notes = info.get('notes', '')
        if len(notes) > 400:
            notes = notes[:400] + "..."

        if not messagebox.askyesno(
                "Update Available",
                f"Version {info['version_str']} is available "
                f"(you have {APP_VERSION}).\n\n"
                f"Download size: {info['size'] / (1024 * 1024):.1f} MB\n"
                + (f"\nWhat's new:\n{notes}\n" if notes else "")
                + "\nDownload it now? It will install when you close the program."):
            return

        if getattr(self, 'status_label', None):
            self.status_label.config(
                text=f"Downloading update {info['version_str']}...", fg='#f39c12')
        self.root.update_idletasks()

        def on_progress(got, total):
            pct = int(got * 100 / total) if total else 0
            try:
                if getattr(self, 'status_label', None):
                    self.status_label.config(
                        text=f"Downloading update {info['version_str']}... {pct}%",
                        fg='#f39c12')
                self.root.update_idletasks()
            except Exception:
                pass

        path = self._download_update(info, progress=on_progress)
        if not path:
            messagebox.showerror(
                "Download Failed",
                "The update could not be downloaded.\n\n"
                + (getattr(self, '_download_error', '') or
                   "Check this computer's internet connection and try again."))
            if getattr(self, 'status_label', None):
                self.status_label.config(text="Update download failed", fg='#e74c3c')
            return

        info['staged'] = path
        self._on_update_ready(info)
        messagebox.showinfo(
            "Update Downloaded",
            f"Version {info['version_str']} is ready.\n\n"
            "It will be installed when you close the program.")

    def on_closing(self):
        """Handle application closing"""
        try:
            # Save data one final time
            self.save_data()
            self.save_settings()
        except:
            pass
        
        # If a newer build was downloaded, offer to install it now
        try:
            self.apply_update_on_exit()
        except Exception as e:
            print(f"Update on exit failed: {e}")
        
        # Close the application
        self.root.destroy()


class ServiceOrderWindow:
    def __init__(self, parent_app, machine):
        self.parent_app = parent_app
        self.machine = machine
        self.linked_estimate_number = machine.get('estimate_number', '')
        self.record_id = str(machine.get('record_id', '') or '')
        
        # Create Toplevel window
        self.window = tk.Toplevel(parent_app.root)
        self.root = self.window  # Alias for compatibility
        company = parent_app.company_name or "Service Order"
        self.root.title(f"{company} - Service Order - {self.linked_estimate_number}")
        self.root.geometry("1000x900")
        self.root.configure(bg='#1a252f')
        
        # Use parent app's settings
        self.price_list_folder = parent_app.price_list_folder
        self.column_mappings = parent_app.column_mappings
        self.price_lists = parent_app.price_lists
        self.company_name = parent_app.company_name
        
        # Service order data
        self.current_service_order = None
        self.parts_entries = []
        self.so_status_label = None
        self.service_location = machine.get('service_location', 'Shop')
        
        # Service orders directory
        self.service_orders_dir = parent_app.service_orders_dir

        # Another computer may have changed the shared labor rates
        try:
            parent_app.load_labor_rates()
        except Exception as e:
            print(f"Could not refresh labor rates: {e}")

        # Create GUI
        self.create_widgets()
        
        # Load price lists if not already loaded by parent
        if self.price_list_folder and not self.price_lists:
            self.load_price_lists()
        
        # Push up anything saved offline previously, then load
        try:
            self.sync_offline_service_orders(silent=True)
        except Exception as e:
            print(f"Service order sync on open failed: {e}")
        
        # Load existing service order or import machine data
        self.import_machine_data(machine)
        
        # Handle window close
        def on_close():
            try:
                parent_app._service_order_windows.remove(self)
            except:
                pass
            self.window.destroy()
        self.window.protocol("WM_DELETE_WINDOW", on_close)
    
    def save_settings(self):
        """Save settings back to parent app"""
        self.parent_app.price_list_folder = self.price_list_folder
        self.parent_app.column_mappings = self.column_mappings
        self.parent_app.save_settings()
    
    def create_widgets(self):
        """Create the main GUI"""
        # Menu bar along the top of the window. Unlike the toolbar it never
        # gets cut off when the window is narrower than the form.
        menubar = tk.Menu(self.window)
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Labor Rates...", command=self.configure_labor_rates)
        settings_menu.add_separator()
        settings_menu.add_command(label="Set Price Lists Folder...",
                                  command=self.set_price_list_folder)
        settings_menu.add_command(label="Configure Price List Columns...",
                                  command=self.configure_price_list_columns)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        self.window.config(menu=menubar)

        # Create main canvas with scrollbar for the whole form
        main_canvas = tk.Canvas(self.root, bg='#1a252f', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        self.scrollable_frame = tk.Frame(main_canvas, bg='#1a252f')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        # Enable mouse wheel scrolling (only for this window)
        def _on_mousewheel(event):
            # Only scroll if this window is focused
            try:
                if event.widget.winfo_toplevel() == self.window:
                    main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        self.window.bind("<MouseWheel>", _on_mousewheel)
        
        scrollbar.pack(side="right", fill="y")
        main_canvas.pack(side="left", fill="both", expand=True)
        
        # Header
        header_frame = tk.Frame(self.scrollable_frame, bg='#2c3e50', height=80)
        header_frame.pack(fill='x', padx=10, pady=5)
        
        header_left = tk.Frame(header_frame, bg='#2c3e50')
        header_left.pack(side='left', padx=20, pady=10)
        
        tk.Label(header_left, text=self.company_name if self.company_name else "Service Order",
                font=('Arial', 24, 'bold'), fg='#3498db', bg='#2c3e50').pack(anchor='w')

        # Kept up here on the left so it's visible at any window width
        # (the right end of the toolbar is cut off on narrower screens)
        tk.Button(header_frame, text="⚙ Labor Rates", command=self.configure_labor_rates,
                  bg='#d35400', fg='white', activebackground='#e67e22',
                  font=('Arial', 11, 'bold'), padx=12, pady=4,
                  cursor='hand2').pack(side='left', padx=10, pady=10)

        header_right = tk.Frame(header_frame, bg='#2c3e50')
        header_right.pack(side='right', padx=20, pady=10)
        
        tk.Label(header_right, text="Service Order", font=('Arial', 24, 'bold'), 
                fg='white', bg='#2c3e50').pack(anchor='e')
        
        # Toolbar
        toolbar = tk.Frame(self.scrollable_frame, bg='#34495e')
        toolbar.pack(fill='x', padx=10, pady=5)
        
        tk.Button(toolbar, text="💾 Save Service Order", command=self.save_service_order,
                 bg='#27ae60', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5, pady=5)
        
        tk.Button(toolbar, text="📄 Export to PDF", command=self.export_to_pdf,
                 bg='#9b59b6', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5, pady=5)

        tk.Button(toolbar, text="📧 Customer Copy PDF", command=self.export_customer_copy_pdf,
                 bg='#e67e22', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5, pady=5)

        tk.Button(toolbar, text="📊 Export to QuickBooks", command=self.export_to_quickbooks_iif,
                 bg='#2980b9', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5, pady=5)
        
        tk.Button(toolbar, text="🗑️ Clear Form", command=self.clear_form,
                 bg='#e74c3c', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5, pady=5)
        
        tk.Button(toolbar, text="📋 Set Price Lists Folder", command=self.set_price_list_folder,
                 bg='#16a085', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5, pady=5)
        
        tk.Button(toolbar, text="🔧 Configure Columns", command=self.configure_price_list_columns,
                 bg='#8e44ad', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5, pady=5)
        
        # Linked machine display
        self.linked_label = tk.Label(toolbar, text="No machine selected", 
                                    font=('Arial', 10, 'italic'), fg='#bdc3c7', bg='#34495e')
        self.linked_label.pack(side='right', padx=20, pady=5)
        
        # Save status line - shows where the last save actually landed
        so_status_frame = tk.Frame(self.scrollable_frame, bg='#1a252f')
        so_status_frame.pack(fill='x', padx=10)
        self.so_status_label = tk.Label(so_status_frame, text="", font=('Arial', 10, 'bold'),
                                       fg='#ecf0f1', bg='#1a252f')
        self.so_status_label.pack(side='left', pady=2)
        
        # Main form area
        form_frame = tk.Frame(self.scrollable_frame, bg='#1a252f')
        form_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # === TOP SECTION: Dates and Order Numbers ===
        top_section = tk.Frame(form_frame, bg='#2c3e50', relief='ridge', bd=2)
        top_section.pack(fill='x', pady=5)
        
        # Left side - Dates
        dates_frame = tk.Frame(top_section, bg='#2c3e50')
        dates_frame.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        
        tk.Label(dates_frame, text="Estimate Completed Date:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=0, sticky='e', pady=5)
        self.estimate_date_var = tk.StringVar(master=self.root)
        tk.Entry(dates_frame, textvariable=self.estimate_date_var, width=20, font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(dates_frame, text="Repair Completed Date:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=1, column=0, sticky='e', pady=5)
        self.repair_date_var = tk.StringVar(master=self.root, value=datetime.now().strftime('%m-%d-%Y'))
        tk.Entry(dates_frame, textvariable=self.repair_date_var, width=20, font=('Arial', 10)).grid(row=1, column=1, padx=5, pady=5)
        
        # Right side - PO and Order numbers
        orders_frame = tk.Frame(top_section, bg='#2c3e50')
        orders_frame.pack(side='right', fill='both', expand=True, padx=10, pady=10)
        
        tk.Label(orders_frame, text="PO #:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=0, sticky='e', pady=5)
        self.po_number_var = tk.StringVar(master=self.root)
        tk.Entry(orders_frame, textvariable=self.po_number_var, width=20, font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(orders_frame, text="Order #:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=1, column=0, sticky='e', pady=5)
        self.order_number_var = tk.StringVar(master=self.root)
        tk.Entry(orders_frame, textvariable=self.order_number_var, width=20, font=('Arial', 10)).grid(row=1, column=1, padx=5, pady=5)
        
        # === CUSTOMER SECTION ===
        customer_section = tk.Frame(form_frame, bg='#2c3e50', relief='ridge', bd=2)
        customer_section.pack(fill='x', pady=5)
        
        # Bill to / Ship to
        cust_left = tk.Frame(customer_section, bg='#2c3e50')
        cust_left.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        
        tk.Label(cust_left, text="Bill to Customer Name:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=0, sticky='e', pady=5)
        self.bill_to_var = tk.StringVar(master=self.root)
        tk.Entry(cust_left, textvariable=self.bill_to_var, width=35, font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(cust_left, text="Ship to Customer Name:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=1, column=0, sticky='e', pady=5)
        self.ship_to_var = tk.StringVar(master=self.root)
        tk.Entry(cust_left, textvariable=self.ship_to_var, width=35, font=('Arial', 10)).grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(cust_left, text="Address:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=2, column=0, sticky='e', pady=5)
        self.address_var = tk.StringVar(master=self.root)
        tk.Entry(cust_left, textvariable=self.address_var, width=35, font=('Arial', 10)).grid(row=2, column=1, padx=5, pady=5)
        
        tk.Label(cust_left, text="Contact:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=3, column=0, sticky='e', pady=5)
        self.contact_var = tk.StringVar(master=self.root)
        tk.Entry(cust_left, textvariable=self.contact_var, width=35, font=('Arial', 10)).grid(row=3, column=1, padx=5, pady=5)
        
        # Right side - Date, Tag, Phone
        cust_right = tk.Frame(customer_section, bg='#2c3e50')
        cust_right.pack(side='right', fill='both', expand=True, padx=10, pady=10)
        
        tk.Label(cust_right, text="DATE:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=0, sticky='e', pady=5)
        self.date_var = tk.StringVar(master=self.root, value=datetime.now().strftime('%m-%d-%Y'))
        tk.Entry(cust_right, textvariable=self.date_var, width=20, font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(cust_right, text="TAG #:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=1, column=0, sticky='e', pady=5)
        self.tag_var = tk.StringVar(master=self.root)
        tk.Entry(cust_right, textvariable=self.tag_var, width=20, font=('Arial', 10)).grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(cust_right, text="Phone #:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=2, column=0, sticky='e', pady=5)
        self.phone_var = tk.StringVar(master=self.root)
        tk.Entry(cust_right, textvariable=self.phone_var, width=20, font=('Arial', 10)).grid(row=2, column=1, padx=5, pady=5)
        
        # === MACHINE SECTION ===
        machine_section = tk.Frame(form_frame, bg='#2c3e50', relief='ridge', bd=2)
        machine_section.pack(fill='x', pady=5)
        
        tk.Label(machine_section, text="MACHINE INFORMATION", font=('Arial', 12, 'bold'),
                fg='#3498db', bg='#2c3e50').pack(anchor='w', padx=10, pady=5)
        
        machine_grid = tk.Frame(machine_section, bg='#2c3e50')
        machine_grid.pack(fill='x', padx=10, pady=5)
        
        # Left machine (original?)
        tk.Label(machine_grid, text="MFD:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=0, sticky='e', pady=3)
        self.mfd1_var = tk.StringVar(master=self.root)
        tk.Entry(machine_grid, textvariable=self.mfd1_var, width=20, font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=3)
        
        tk.Label(machine_grid, text="Model:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=1, column=0, sticky='e', pady=3)
        self.model1_var = tk.StringVar(master=self.root)
        tk.Entry(machine_grid, textvariable=self.model1_var, width=20, font=('Arial', 10)).grid(row=1, column=1, padx=5, pady=3)
        
        tk.Label(machine_grid, text="S/N #:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=2, column=0, sticky='e', pady=3)
        self.sn1_var = tk.StringVar(master=self.root)
        tk.Entry(machine_grid, textvariable=self.sn1_var, width=20, font=('Arial', 10)).grid(row=2, column=1, padx=5, pady=3)
        
        # Spacer
        tk.Label(machine_grid, text="    ", bg='#2c3e50').grid(row=0, column=2)
        
        # Right machine (secondary?)
        tk.Label(machine_grid, text="MFD:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=3, sticky='e', pady=3)
        self.mfd2_var = tk.StringVar(master=self.root)
        tk.Entry(machine_grid, textvariable=self.mfd2_var, width=20, font=('Arial', 10)).grid(row=0, column=4, padx=5, pady=3)
        
        tk.Label(machine_grid, text="Model:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=1, column=3, sticky='e', pady=3)
        self.model2_var = tk.StringVar(master=self.root)
        tk.Entry(machine_grid, textvariable=self.model2_var, width=20, font=('Arial', 10)).grid(row=1, column=4, padx=5, pady=3)
        
        tk.Label(machine_grid, text="S/N #:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=2, column=3, sticky='e', pady=3)
        self.sn2_var = tk.StringVar(master=self.root)
        tk.Entry(machine_grid, textvariable=self.sn2_var, width=20, font=('Arial', 10)).grid(row=2, column=4, padx=5, pady=3)
        
        # Input Power and Plug
        power_frame = tk.Frame(machine_section, bg='#2c3e50')
        power_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(power_frame, text="Input Power Setting:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').pack(side='left')
        self.power_var = tk.StringVar(master=self.root)
        tk.Entry(power_frame, textvariable=self.power_var, width=15, font=('Arial', 10)).pack(side='left', padx=5)
        
        tk.Label(power_frame, text="    Plug Included With Machine?", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').pack(side='left', padx=20)
        self.plug_var = tk.StringVar(master=self.root, value="NO")
        tk.Radiobutton(power_frame, text="YES", variable=self.plug_var, value="YES",
                      bg='#2c3e50', fg='white', selectcolor='#27ae60', font=('Arial', 10)).pack(side='left')
        tk.Radiobutton(power_frame, text="NO", variable=self.plug_var, value="NO",
                      bg='#2c3e50', fg='white', selectcolor='#e74c3c', font=('Arial', 10)).pack(side='left')
        
        # === SERVICE SECTION ===
        service_section = tk.Frame(form_frame, bg='#2c3e50', relief='ridge', bd=2)
        service_section.pack(fill='x', pady=5)
        
        # Reason for Service
        reason_frame = tk.Frame(service_section, bg='#2c3e50')
        reason_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(reason_frame, text="Reason For Service / Complaint:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').pack(anchor='w')
        self.reason_text = tk.Text(reason_frame, height=3, width=80, font=('Arial', 10))
        self.reason_text.pack(fill='x', pady=5)
        
        # Pre-power check
        prepower_frame = tk.Frame(service_section, bg='#2c3e50')
        prepower_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(prepower_frame, text="Pre-Power Check Completed?", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').pack(side='left')
        self.prepower_var = tk.StringVar(master=self.root, value="NO")
        tk.Radiobutton(prepower_frame, text="YES", variable=self.prepower_var, value="YES",
                      bg='#2c3e50', fg='white', selectcolor='#27ae60', font=('Arial', 10)).pack(side='left', padx=10)
        tk.Radiobutton(prepower_frame, text="NO", variable=self.prepower_var, value="NO",
                      bg='#2c3e50', fg='white', selectcolor='#e74c3c', font=('Arial', 10)).pack(side='left')
        
        # Description of Work
        work_frame = tk.Frame(service_section, bg='#2c3e50')
        work_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(work_frame, text="Description of Work:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').pack(anchor='w')
        self.work_text = tk.Text(work_frame, height=5, width=80, font=('Arial', 10))
        self.work_text.pack(fill='x', pady=5)
        
        # === PARTS SECTION ===
        parts_section = tk.Frame(form_frame, bg='#2c3e50', relief='ridge', bd=2)
        parts_section.pack(fill='x', pady=5)
        
        tk.Label(parts_section, text="PARTS", font=('Arial', 12, 'bold'),
                fg='#3498db', bg='#2c3e50').pack(anchor='w', padx=10, pady=5)
        
        # Parts header - using grid for precise alignment
        parts_header = tk.Frame(parts_section, bg='#34495e')
        parts_header.pack(fill='x', padx=10)
        
        tk.Label(parts_header, text="Part #", font=('Arial', 10, 'bold'), width=16,
                fg='white', bg='#34495e', anchor='w').grid(row=0, column=0, padx=1)
        tk.Label(parts_header, text="Description", font=('Arial', 10, 'bold'), width=35,
                fg='white', bg='#34495e', anchor='w').grid(row=0, column=1, padx=1)
        tk.Label(parts_header, text="Qty", font=('Arial', 10, 'bold'), width=8,
                fg='white', bg='#34495e', anchor='w').grid(row=0, column=2, padx=1)
        tk.Label(parts_header, text="Each $", font=('Arial', 10, 'bold'), width=10,
                fg='white', bg='#34495e', anchor='w').grid(row=0, column=3, padx=1)
        tk.Label(parts_header, text="Total $", font=('Arial', 10, 'bold'), width=10,
                fg='white', bg='#34495e', anchor='w').grid(row=0, column=4, padx=1)
        
        # Parts rows (14 rows like the Excel) - using grid
        self.parts_entries = []
        for i in range(14):
            row_frame = tk.Frame(parts_section, bg='#2c3e50')
            row_frame.pack(fill='x', padx=10, pady=1)
            
            part_num = tk.StringVar(master=self.root)
            desc = tk.StringVar(master=self.root)
            qty = tk.StringVar(master=self.root)
            each = tk.StringVar(master=self.root)
            total = tk.StringVar(master=self.root, value="0.00")
            
            # Part number entry with autocomplete
            part_entry = tk.Entry(row_frame, textvariable=part_num, width=16, font=('Arial', 9))
            part_entry.grid(row=0, column=0, padx=1)
            
            desc_entry = tk.Entry(row_frame, textvariable=desc, width=35, font=('Arial', 9))
            desc_entry.grid(row=0, column=1, padx=1)
            
            qty_entry = tk.Entry(row_frame, textvariable=qty, width=8, font=('Arial', 9))
            qty_entry.grid(row=0, column=2, padx=1)
            
            each_entry = tk.Entry(row_frame, textvariable=each, width=10, font=('Arial', 9))
            each_entry.grid(row=0, column=3, padx=1)
            
            total_label = tk.Label(row_frame, textvariable=total, width=10, font=('Arial', 9),
                                  bg='#ecf0f1', relief='sunken', anchor='e')
            total_label.grid(row=0, column=4, padx=1)
            
            # Bind autocomplete to part number entry
            self.setup_part_autocomplete(part_entry, part_num, desc, each)
            
            # Bind calculation
            try:
                qty.trace_add('write', lambda *args, q=qty, e=each, t=total: self.calculate_line_total(q, e, t))
            except AttributeError:
                qty.trace('w', lambda *args, q=qty, e=each, t=total: self.calculate_line_total(q, e, t))
            try:
                each.trace_add('write', lambda *args, q=qty, e=each, t=total: self.calculate_line_total(q, e, t))
            except AttributeError:
                each.trace('w', lambda *args, q=qty, e=each, t=total: self.calculate_line_total(q, e, t))
            
            self.parts_entries.append({
                'part_num': part_num,
                'desc': desc,
                'qty': qty,
                'each': each,
                'total': total,
                'part_entry': part_entry,
                'desc_entry': desc_entry,
                'each_entry': each_entry
            })
        
        # Additional parts needed?
        additional_frame = tk.Frame(parts_section, bg='#2c3e50')
        additional_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(additional_frame, text="Repair May Require Additional Parts and Labor?", 
                font=('Arial', 10, 'bold'), fg='white', bg='#2c3e50').pack(side='left')
        self.additional_var = tk.StringVar(master=self.root, value="NO")
        tk.Radiobutton(additional_frame, text="YES", variable=self.additional_var, value="YES",
                      bg='#2c3e50', fg='white', selectcolor='#27ae60', font=('Arial', 10)).pack(side='left', padx=10)
        tk.Radiobutton(additional_frame, text="NO", variable=self.additional_var, value="NO",
                      bg='#2c3e50', fg='white', selectcolor='#e74c3c', font=('Arial', 10)).pack(side='left')
        
        # === TOTALS SECTION ===
        totals_section = tk.Frame(form_frame, bg='#2c3e50', relief='ridge', bd=2)
        totals_section.pack(fill='x', pady=5)
        
        # Add TOTALS header
        tk.Label(totals_section, text="TOTALS", font=('Arial', 12, 'bold'),
                fg='#3498db', bg='#2c3e50').pack(anchor='w', padx=10, pady=5)
        
        totals_grid = tk.Frame(totals_section, bg='#2c3e50')
        totals_grid.pack(fill='x', padx=20, pady=10)
        
        # Parts Total
        tk.Label(totals_grid, text="Parts Total:", font=('Arial', 12, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=0, sticky='e', pady=5, padx=5)
        self.parts_total_var = tk.StringVar(master=self.root, value="$0.00")
        tk.Label(totals_grid, textvariable=self.parts_total_var, font=('Arial', 12, 'bold'),
                fg='#27ae60', bg='#ecf0f1', width=12, anchor='e', relief='sunken').grid(row=0, column=1, pady=5, padx=5, sticky='w')
        
        # Labor Qty
        tk.Label(totals_grid, text="Labor Qty (hrs):", font=('Arial', 12, 'bold'),
                fg='white', bg='#2c3e50').grid(row=1, column=0, sticky='e', pady=5, padx=5)
        self.labor_qty_var = tk.StringVar(master=self.root, value="0")
        labor_entry = tk.Entry(totals_grid, textvariable=self.labor_qty_var, width=12, font=('Arial', 12))
        labor_entry.grid(row=1, column=1, pady=5, padx=5, sticky='w')
        
        # Labor Rate - picked from the configured rates (Settings > Labor Rates), not typed.
        # labor_rate_var holds the price used everywhere else; the dropdown
        # shows the rate's name.
        tk.Label(totals_grid, text="Labor Rate:", font=('Arial', 12, 'bold'),
                fg='white', bg='#2c3e50').grid(row=2, column=0, sticky='e', pady=5, padx=5)
        rate_cell = tk.Frame(totals_grid, bg='#2c3e50')
        rate_cell.grid(row=2, column=1, pady=5, padx=5, sticky='w')
        self.labor_rate_var = tk.StringVar(master=self.root, value="0.00")
        self.labor_rate_name_var = tk.StringVar(master=self.root, value="")
        self.labor_rate_combo = ttk.Combobox(rate_cell, textvariable=self.labor_rate_name_var,
                                             state='readonly', width=24, font=('Arial', 12))
        self.labor_rate_combo.pack(side='left')
        self.labor_rate_combo.bind('<<ComboboxSelected>>', self._on_labor_rate_selected)
        tk.Button(rate_cell, text="⚙ Edit Rates", command=self.configure_labor_rates,
                  bg='#7f8c8d', fg='white', font=('Arial', 9, 'bold'),
                  cursor='hand2').pack(side='left', padx=(8, 0))
        self.labor_rate_display = tk.Label(rate_cell, text="", font=('Arial', 12, 'bold'),
                                           fg='#ecf0f1', bg='#2c3e50')
        self.labor_rate_display.pack(side='left', padx=10)
        
        # Labor Total
        tk.Label(totals_grid, text="Labor Total:", font=('Arial', 12, 'bold'),
                fg='white', bg='#2c3e50').grid(row=3, column=0, sticky='e', pady=5, padx=5)
        self.labor_total_var = tk.StringVar(master=self.root, value="$0.00")
        tk.Label(totals_grid, textvariable=self.labor_total_var, font=('Arial', 12, 'bold'),
                fg='#27ae60', bg='#ecf0f1', width=12, anchor='e', relief='sunken').grid(row=3, column=1, pady=5, padx=5, sticky='w')
        
        # Grand Total
        tk.Label(totals_grid, text="GRAND TOTAL:", font=('Arial', 14, 'bold'),
                fg='#f39c12', bg='#2c3e50').grid(row=4, column=0, sticky='e', pady=10, padx=5)
        self.grand_total_var = tk.StringVar(master=self.root, value="$0.00")
        tk.Label(totals_grid, textvariable=self.grand_total_var, font=('Arial', 14, 'bold'),
                fg='#f39c12', bg='#ecf0f1', width=12, anchor='e', relief='sunken').grid(row=4, column=1, pady=10, padx=5, sticky='w')
        
        # Bind labor calculations
        try:
            self.labor_qty_var.trace_add('write', lambda *args: self.calculate_totals())
        except AttributeError:
            self.labor_qty_var.trace('w', lambda *args: self.calculate_totals())
        try:
            self.labor_rate_var.trace_add('write', lambda *args: self.calculate_totals())
        except AttributeError:
            self.labor_rate_var.trace('w', lambda *args: self.calculate_totals())

        # New orders start on the first configured rate
        self.select_default_labor_rate()

        # === SIGNATURE SECTION ===
        sig_section = tk.Frame(form_frame, bg='#2c3e50', relief='ridge', bd=2)
        sig_section.pack(fill='x', pady=5)
        
        sig_grid = tk.Frame(sig_section, bg='#2c3e50')
        sig_grid.pack(fill='x', padx=10, pady=10)
        
        tk.Label(sig_grid, text="Service Engineer:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=0, sticky='e', pady=5)
        self.engineer_var = tk.StringVar(master=self.root)
        tk.Entry(sig_grid, textvariable=self.engineer_var, width=30, font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(sig_grid, text="Date:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').grid(row=0, column=2, sticky='e', pady=5, padx=20)
        self.sig_date_var = tk.StringVar(master=self.root, value=datetime.now().strftime('%m-%d-%Y'))
        tk.Entry(sig_grid, textvariable=self.sig_date_var, width=15, font=('Arial', 10)).grid(row=0, column=3, padx=5, pady=5)
        
        # === CUSTOMER SIGNATURE PAD ===
        cust_sig_section = tk.LabelFrame(form_frame, text="  Customer Acknowledgment  ",
                                          font=('Arial', 11, 'bold'), fg='#f39c12',
                                          bg='#2c3e50', bd=2)
        cust_sig_section.pack(fill='x', pady=5)
        
        # Agreement text
        tk.Label(cust_sig_section, 
                text="I acknowledge that the work described above has been completed to my satisfaction.",
                font=('Arial', 10, 'italic'), fg='#ecf0f1', bg='#2c3e50',
                wraplength=900, justify='left').pack(anchor='w', padx=10, pady=(8, 2))
        
        # Customer name and date row
        cust_info = tk.Frame(cust_sig_section, bg='#2c3e50')
        cust_info.pack(fill='x', padx=10, pady=2)
        
        tk.Label(cust_info, text="Customer Name:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').pack(side='left', padx=(0, 5))
        self.customer_sig_name_var = tk.StringVar(master=self.root)
        tk.Entry(cust_info, textvariable=self.customer_sig_name_var, width=30,
                font=('Arial', 10)).pack(side='left', padx=5)
        
        tk.Label(cust_info, text="Date:", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').pack(side='left', padx=(20, 5))
        self.customer_sig_date_var = tk.StringVar(master=self.root, value=datetime.now().strftime('%m-%d-%Y'))
        tk.Entry(cust_info, textvariable=self.customer_sig_date_var, width=15,
                font=('Arial', 10)).pack(side='left', padx=5)
        
        # Customer email - where the customer copy PDF gets sent
        email_row = tk.Frame(cust_sig_section, bg='#2c3e50')
        email_row.pack(fill='x', padx=10, pady=(6, 2))

        tk.Label(email_row, text="Send Copy To (Email):", font=('Arial', 10, 'bold'),
                fg='white', bg='#2c3e50').pack(side='left', padx=(0, 5))
        self.customer_email_var = tk.StringVar(master=self.root)
        tk.Entry(email_row, textvariable=self.customer_email_var, width=40,
                font=('Arial', 10)).pack(side='left', padx=5)

        self.email_status_label = tk.Label(email_row, text="", font=('Arial', 9, 'italic'),
                                           fg='#95a5a6', bg='#2c3e50')
        self.email_status_label.pack(side='left', padx=10)

        def _validate_email_display(*_args):
            addr = self.customer_email_var.get().strip()
            if not addr:
                self.email_status_label.config(text="", fg='#95a5a6')
            elif self._is_valid_email(addr):
                self.email_status_label.config(text="✔ valid", fg='#2ecc71')
            else:
                self.email_status_label.config(text="⚠ check address", fg='#f39c12')

        try:
            self.customer_email_var.trace_add('write', _validate_email_display)
        except AttributeError:
            self.customer_email_var.trace('w', _validate_email_display)

        tk.Label(cust_sig_section,
                text="The customer copy PDF will be attached to a new Outlook message addressed here.",
                font=('Arial', 8, 'italic'), fg='#95a5a6', bg='#2c3e50').pack(anchor='w', padx=10)

        # Signature canvas
        sig_canvas_frame = tk.Frame(cust_sig_section, bg='#2c3e50')
        sig_canvas_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(sig_canvas_frame, text="Signature (sign below with stylus, finger, or mouse):",
                font=('Arial', 9, 'italic'), fg='#bdc3c7', bg='#2c3e50').pack(anchor='w')
        
        # White canvas for signature
        self.signature_canvas = tk.Canvas(sig_canvas_frame, bg='white', height=140,
                                          highlightthickness=2, highlightbackground='#7f8c8d',
                                          cursor='pencil')
        self.signature_canvas.pack(fill='x', pady=3)
        
        # Store strokes as list of point lists (for save/load/PDF)
        self._signature_strokes = []
        self._current_stroke = []
        self._last_sig_pos = None
        # Canvas size at the moment of signing - needed to scale the signature
        # correctly into a PDF even if the window is later resized.
        self._signature_canvas_size = None

        def on_sig_press(event):
            self._last_sig_pos = (event.x, event.y)
            self._current_stroke = [(event.x, event.y)]
            # Capture the canvas size the first time the customer touches the pad
            if not self._signature_canvas_size:
                try:
                    w = self.signature_canvas.winfo_width()
                    h = self.signature_canvas.winfo_height()
                    if w > 1 and h > 1:
                        self._signature_canvas_size = [w, h]
                except Exception:
                    pass
        
        def on_sig_drag(event):
            if self._last_sig_pos:
                self.signature_canvas.create_line(
                    self._last_sig_pos[0], self._last_sig_pos[1],
                    event.x, event.y,
                    width=2, fill='black', capstyle='round', smooth=True,
                    tags='signature'
                )
                self._current_stroke.append((event.x, event.y))
                self._last_sig_pos = (event.x, event.y)
        
        def on_sig_release(event):
            if self._current_stroke:
                self._signature_strokes.append(self._current_stroke)
                self._current_stroke = []
            self._last_sig_pos = None
        
        # Bind mouse/stylus events
        self.signature_canvas.bind('<Button-1>', on_sig_press)
        self.signature_canvas.bind('<B1-Motion>', on_sig_drag)
        self.signature_canvas.bind('<ButtonRelease-1>', on_sig_release)
        
        # Buttons
        sig_btn_frame = tk.Frame(cust_sig_section, bg='#2c3e50')
        sig_btn_frame.pack(fill='x', padx=10, pady=(2, 10))
        
        def clear_signature():
            self.signature_canvas.delete('signature')
            self._signature_strokes = []
            self._current_stroke = []
            self._last_sig_pos = None
            self._signature_canvas_size = None

        tk.Button(sig_btn_frame, text="🗑 Clear Signature", command=clear_signature,
                 bg='#e74c3c', fg='white', font=('Arial', 10, 'bold')).pack(side='right')
        
        tk.Label(sig_btn_frame, text="Tip: Sign with stylus for best results on tablets",
                font=('Arial', 9, 'italic'), fg='#95a5a6', bg='#2c3e50').pack(side='left')
    
    def setup_part_autocomplete(self, entry_widget, part_var, desc_var, each_var):
        """Setup autocomplete for a part number entry field"""
        # Store reference to current autocomplete window
        self._current_autocomplete = None
        self._autocomplete_listbox = None
        self._autocomplete_results = []
        
        def close_autocomplete():
            if self._current_autocomplete:
                try:
                    self._current_autocomplete.destroy()
                except:
                    pass
                self._current_autocomplete = None
                self._autocomplete_listbox = None
                self._autocomplete_results = []
        
        def on_select_item(event=None):
            """Handle selection from the autocomplete list"""
            if self._autocomplete_listbox and self._autocomplete_results:
                selection = self._autocomplete_listbox.curselection()
                if selection:
                    idx = selection[0]
                    if idx < len(self._autocomplete_results):
                        selected = self._autocomplete_results[idx]
                        part_var.set(selected['part_num'])
                        desc_var.set(selected['desc'])
                        each_var.set(f"{selected['price']:.2f}")
            close_autocomplete()
            entry_widget.focus_set()
        
        def on_key_release(event):
            # Ignore navigation keys
            if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'Escape'):
                return
            
            # Get current text
            text = part_var.get().strip()
            
            # Close existing autocomplete list
            close_autocomplete()
            
            if len(text) < 2:
                return
            
            # Get vendor from MFD field
            vendor = self.mfd1_var.get() if hasattr(self, 'mfd1_var') else None
            
            # Search for matching parts
            results = self.search_part_in_price_list(text, vendor)
            
            if not results:
                return
            
            self._autocomplete_results = results
            
            # Create autocomplete dropdown
            self._current_autocomplete = tk.Toplevel(self.root)
            self._current_autocomplete.wm_overrideredirect(True)
            self._current_autocomplete.configure(bg='white', relief='solid', borderwidth=1)
            
            # Position below the entry widget
            x = entry_widget.winfo_rootx()
            y = entry_widget.winfo_rooty() + entry_widget.winfo_height()
            self._current_autocomplete.geometry(f"+{x}+{y}")
            
            # Create listbox with results
            self._autocomplete_listbox = tk.Listbox(
                self._current_autocomplete, 
                width=80, 
                height=min(len(results), 10),
                font=('Arial', 9), 
                selectbackground='#3498db',
                selectforeground='white',
                activestyle='none'
            )
            self._autocomplete_listbox.pack()
            
            for r in results:
                display = f"{r['part_num']} | {r['desc'][:40]} | ${r['price']:.2f}"
                self._autocomplete_listbox.insert(tk.END, display)
            
            # Select first item by default
            if results:
                self._autocomplete_listbox.selection_set(0)
            
            # Bind events - single click and double click both select
            self._autocomplete_listbox.bind('<ButtonRelease-1>', on_select_item)
            self._autocomplete_listbox.bind('<Double-1>', on_select_item)
            self._autocomplete_listbox.bind('<Return>', on_select_item)
            self._autocomplete_listbox.bind('<Escape>', lambda e: close_autocomplete())
            
            # Keep the autocomplete window on top
            self._current_autocomplete.lift()
            self._current_autocomplete.attributes('-topmost', True)
        
        def on_entry_key(event):
            """Handle special keys in the entry field"""
            if event.keysym == 'Escape':
                close_autocomplete()
                return
            
            if event.keysym == 'Down' and self._autocomplete_listbox:
                # Move focus to listbox
                self._autocomplete_listbox.focus_set()
                if not self._autocomplete_listbox.curselection():
                    self._autocomplete_listbox.selection_set(0)
                return "break"
            
            if event.keysym == 'Return' and self._autocomplete_listbox:
                on_select_item()
                return "break"
        
        def on_tab(event):
            """When Tab is pressed, try to auto-fill from exact match"""
            close_autocomplete()
            text = part_var.get().strip()
            if text:
                vendor = self.mfd1_var.get() if hasattr(self, 'mfd1_var') else None
                exact = self.get_exact_part(text, vendor)
                if exact:
                    part_var.set(exact['part_num'])
                    desc_var.set(exact['desc'])
                    each_var.set(f"{exact['price']:.2f}")
        
        def on_focus_out(event):
            # Delay closing to allow click on listbox to register
            self.root.after(250, lambda: close_autocomplete() if self._current_autocomplete and 
                           self.root.focus_get() != self._autocomplete_listbox else None)
        
        entry_widget.bind('<KeyRelease>', on_key_release)
        entry_widget.bind('<Key>', on_entry_key)
        entry_widget.bind('<Tab>', on_tab)
        entry_widget.bind('<FocusOut>', on_focus_out)

    def calculate_line_total(self, qty_var, each_var, total_var):
        """Calculate total for a single parts line"""
        try:
            qty = float(qty_var.get()) if qty_var.get() else 0
            each = float(each_var.get()) if each_var.get() else 0
            total = qty * each
            total_var.set(f"{total:.2f}")
        except ValueError:
            total_var.set("0.00")
        
        self.calculate_totals()
    
    def calculate_totals(self):
        """Calculate all totals"""
        # Parts total
        parts_total = 0
        for entry in self.parts_entries:
            try:
                parts_total += float(entry['total'].get())
            except ValueError:
                pass
        
        self.parts_total_var.set(f"${parts_total:.2f}")
        
        # Labor total
        try:
            labor_qty = float(self.labor_qty_var.get()) if self.labor_qty_var.get() else 0
            labor_rate = float(self.labor_rate_var.get()) if self.labor_rate_var.get() else 0
            labor_total = labor_qty * labor_rate
        except ValueError:
            labor_total = 0
        
        self.labor_total_var.set(f"${labor_total:.2f}")
        
        # Grand total
        grand_total = parts_total + labor_total
        self.grand_total_var.set(f"${grand_total:.2f}")

    # ---------------- Labor rate dropdown ----------------

    def _configured_labor_rates(self):
        return list(getattr(self.parent_app, 'labor_rates', []) or [])

    def _labor_line_label(self):
        """'Labor - Field Service' on PDFs, or just 'Labor' if unnamed."""
        name = self.labor_rate_name_var.get().strip()
        return f"Labor - {name}" if name else "Labor"

    def _find_labor_rate(self, name):
        for r in self._configured_labor_rates():
            if r['name'] == name:
                return r
        return None

    def _update_labor_rate_display(self):
        """Show the actual $/hr next to the dropdown."""
        rates = self._configured_labor_rates()
        price = parse_money(self.labor_rate_var.get())
        if not rates and not self.labor_rate_name_var.get():
            self.labor_rate_display.config(
                text="No labor rates set up - click ⚙ Edit Rates",
                fg='#e67e22')
            return
        text = f"${price:.2f}/hr" if price is not None else ""
        current = self._find_labor_rate(self.labor_rate_name_var.get())
        if price is not None and (current is None or parse_money(current['rate']) != price):
            # Saved order whose rate has since changed, been removed, or
            # predates named rates - it keeps the price it was saved with
            text += "  (saved on this order)"
            self.labor_rate_display.config(text=text, fg='#f39c12')
        else:
            self.labor_rate_display.config(text=text, fg='#ecf0f1')

    def refresh_labor_rate_choices(self):
        """Reload the dropdown from the configured rates. Called after the
        rates are edited. If the selected rate's price changed, it follows."""
        rates = self._configured_labor_rates()
        self.labor_rate_combo['values'] = [r['name'] for r in rates]
        current = self._find_labor_rate(self.labor_rate_name_var.get())
        if current:
            self.labor_rate_var.set(current['rate'])
        elif not self.labor_rate_name_var.get() and not parse_money(self.labor_rate_var.get()):
            # Nothing chosen yet (e.g. rates were empty before) - use the default
            self.select_default_labor_rate()
            return
        self._update_labor_rate_display()

    def select_default_labor_rate(self):
        rates = self._configured_labor_rates()
        self.labor_rate_combo['values'] = [r['name'] for r in rates]
        if rates:
            self.labor_rate_name_var.set(rates[0]['name'])
            self.labor_rate_var.set(rates[0]['rate'])
        else:
            self.labor_rate_name_var.set('')
            self.labor_rate_var.set('0.00')
        self._update_labor_rate_display()

    def _on_labor_rate_selected(self, event=None):
        r = self._find_labor_rate(self.labor_rate_name_var.get())
        if r:
            self.labor_rate_var.set(r['rate'])
        self._update_labor_rate_display()

    def set_saved_labor_rate(self, name, rate):
        """Show the rate a saved order was priced with. The saved price is
        kept even if that rate has been changed or removed since."""
        self.labor_rate_combo['values'] = [r['name'] for r in self._configured_labor_rates()]
        price = parse_money(rate)
        self.labor_rate_var.set(f"{price:.2f}" if price is not None else '0.00')
        name = str(name or '').strip()
        if not name:
            # Order saved before named rates: show the matching name if the
            # price lines up with one, otherwise just the saved price
            for r in self._configured_labor_rates():
                if price is not None and parse_money(r['rate']) == price:
                    name = r['name']
                    break
        self.labor_rate_name_var.set(name)
        self._update_labor_rate_display()

    def configure_labor_rates(self):
        """Set up to MAX_LABOR_RATES named labor rates for the dropdown."""
        dlg = tk.Toplevel(self.window)
        dlg.title("Labor Rates")
        dlg.configure(bg='#2c3e50')
        dlg.transient(self.window)
        dlg.resizable(False, False)

        tk.Label(dlg, text="Labor Rates", font=('Arial', 16, 'bold'),
                 fg='white', bg='#2c3e50').pack(anchor='w', padx=20, pady=(15, 2))
        tk.Label(dlg, justify='left', fg='#bdc3c7', bg='#2c3e50', font=('Arial', 10),
                 text=(f"Enter up to {MAX_LABOR_RATES} rates. Only rows with a name appear "
                       "in the Labor Rate dropdown.\nThe first one is the default for new "
                       "service orders. Shared with every computer.")
                 ).pack(anchor='w', padx=20, pady=(0, 10))

        grid = tk.Frame(dlg, bg='#2c3e50')
        grid.pack(padx=20)
        for col, text in enumerate(("", "Name (shown in dropdown)", "Rate ($/hr)")):
            tk.Label(grid, text=text, font=('Arial', 10, 'bold'), fg='#ecf0f1',
                     bg='#2c3e50').grid(row=0, column=col, sticky='w', padx=4, pady=(0, 4))

        current = self._configured_labor_rates()
        rows = []
        for i in range(MAX_LABOR_RATES):
            existing = current[i] if i < len(current) else {'name': '', 'rate': ''}
            name_var = tk.StringVar(master=dlg, value=existing['name'])
            rate_var = tk.StringVar(master=dlg, value=existing['rate'])
            tk.Label(grid, text=f"{i + 1}.", fg='#95a5a6', bg='#2c3e50',
                     font=('Arial', 10)).grid(row=i + 1, column=0, sticky='e', padx=4)
            name_entry = tk.Entry(grid, textvariable=name_var, width=30, font=('Arial', 11))
            name_entry.grid(row=i + 1, column=1, padx=4, pady=2)
            tk.Entry(grid, textvariable=rate_var, width=12, font=('Arial', 11),
                     justify='right').grid(row=i + 1, column=2, padx=4, pady=2)
            rows.append((name_var, rate_var))
            if i == 0:
                name_entry.focus_set()

        def save():
            rates, names = [], set()
            for i, (name_var, rate_var) in enumerate(rows, start=1):
                name = ' '.join(name_var.get().split())
                rate_text = rate_var.get().strip()
                if not name and not rate_text:
                    continue
                if not name:
                    messagebox.showerror("Labor Rates",
                                         f"Row {i}: give the ${rate_text} rate a name, "
                                         "or clear it.", parent=dlg)
                    return
                price = parse_money(rate_text)
                if price is None or price < 0:
                    messagebox.showerror("Labor Rates",
                                         f"Row {i}: enter a rate in dollars per hour "
                                         f"for '{name}' (for example 150.00).", parent=dlg)
                    return
                if name.lower() in names:
                    messagebox.showerror("Labor Rates",
                                         f"Row {i}: '{name}' is used twice. Each rate "
                                         "needs its own name.", parent=dlg)
                    return
                names.add(name.lower())
                rates.append({'name': name, 'rate': f"{price:.2f}"})

            shared_ok = self.parent_app.save_labor_rates(rates)
            dlg.destroy()
            if not shared_ok:
                messagebox.showwarning(
                    "Labor Rates",
                    "Saved on this computer, but the shared data folder couldn't be "
                    "reached.\n\nThe rates will be shared with the other computers "
                    "next time this program starts with the folder available.",
                    parent=self.window)

        buttons = tk.Frame(dlg, bg='#2c3e50')
        buttons.pack(fill='x', padx=20, pady=15)
        tk.Button(buttons, text="Save", command=save, bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'), width=10).pack(side='right')
        tk.Button(buttons, text="Cancel", command=dlg.destroy, bg='#7f8c8d', fg='white',
                  font=('Arial', 11), width=10).pack(side='right', padx=8)
        dlg.bind('<Return>', lambda e: save())
        dlg.bind('<Escape>', lambda e: dlg.destroy())

        dlg.update_idletasks()
        x = self.window.winfo_rootx() + max(0, (self.window.winfo_width() - dlg.winfo_width()) // 2)
        y = self.window.winfo_rooty() + 80
        dlg.geometry(f"+{x}+{y}")
        try:
            dlg.grab_set()
        except tk.TclError:
            pass

    def set_price_list_folder(self):
        """Set the folder containing vendor price list Excel files"""
        folder = filedialog.askdirectory(
            title="Select Price List Folder",
            initialdir=self.price_list_folder if self.price_list_folder else None
        )
        if folder:
            self.price_list_folder = folder
            self.save_settings()
            self.load_price_lists()
            
            # Show summary if no errors were shown
            if self.price_lists and not any(self.price_lists):
                pass  # Error message already shown in load_price_lists
            elif self.price_lists:
                # Build success message
                msg = f"Price list folder set to:\n{folder}\n\n"
                msg += f"Successfully loaded {len(self.price_lists)} vendor list(s):\n"
                for vendor, parts in self.price_lists.items():
                    msg += f"  • {vendor}: {len(parts)} parts\n"
                messagebox.showinfo("Success", msg)
    
    def scan_spreadsheet_headers(self, filepath):
        """Scan a spreadsheet and return all column headers found"""
        headers = []
        
        if filepath.lower().endswith('.xlsx'):
            if not OPENPYXL_AVAILABLE:
                return headers
            try:
                wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(min_row=1, max_row=30, values_only=True):
                        if row is None:
                            continue
                        for col_num, cell in enumerate(row):
                            if cell and isinstance(cell, str) and cell.strip():
                                header_text = cell.strip()
                                if header_text not in [h[1] for h in headers]:
                                    # Store as (column_index, header_text, sheet_title)
                                    headers.append((col_num, header_text, sheet.title))
                wb.close()
            except Exception as e:
                print(f"Error scanning headers from {filepath}: {e}")
        
        elif filepath.lower().endswith('.xls'):
            if not XLRD_AVAILABLE:
                return headers
            try:
                wb = xlrd.open_workbook(filepath)
                for sheet_idx in range(wb.nsheets):
                    sheet = wb.sheet_by_index(sheet_idx)
                    for row_num in range(min(30, sheet.nrows)):
                        for col_num in range(sheet.ncols):
                            cell = sheet.cell_value(row_num, col_num)
                            if cell and isinstance(cell, str) and cell.strip():
                                header_text = cell.strip()
                                if header_text not in [h[1] for h in headers]:
                                    headers.append((col_num, header_text, sheet.name))
            except Exception as e:
                print(f"Error scanning headers from {filepath}: {e}")
        
        return headers
    
    def configure_price_list_columns(self):
        """Show dialog to configure which columns to use for each price list spreadsheet"""
        if not self.price_list_folder or not os.path.exists(self.price_list_folder):
            messagebox.showwarning("Warning", "Please set the Price Lists Folder first.")
            return
        
        # Find all Excel files
        excel_files = []
        for filename in sorted(os.listdir(self.price_list_folder)):
            if (filename.lower().endswith(('.xlsx', '.xls')) and not filename.startswith('~$')):
                excel_files.append(filename)
        
        if not excel_files:
            messagebox.showwarning("Warning", "No Excel files found in the price list folder.")
            return
        
        # Create config window
        config_win = tk.Toplevel(self.root)
        config_win.title("Configure Price List Columns")
        config_win.configure(bg='#2c3e50')
        config_win.transient(self.root)
        config_win.grab_set()
        
        # Center window
        config_win.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        tk.Label(config_win, text="Configure Price List Columns", 
                font=('Arial', 16, 'bold'), fg='white', bg='#2c3e50').pack(pady=15)
        
        tk.Label(config_win, text="Choose which column to use for Part Number and Price in each spreadsheet.\n"
                                  "Leave as 'Auto-Detect' to use the default automatic detection.",
                font=('Arial', 10), fg='#bdc3c7', bg='#2c3e50').pack(pady=(0, 10))
        
        # Scrollable frame for the file list
        canvas_frame = tk.Frame(config_win, bg='#2c3e50')
        canvas_frame.pack(fill='both', expand=True, padx=15, pady=5)
        
        canvas = tk.Canvas(canvas_frame, bg='#2c3e50', highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scroll_inner = tk.Frame(canvas, bg='#2c3e50')
        
        scroll_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Enable mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Store dropdowns for saving
        file_widgets = {}
        
        for file_idx, filename in enumerate(excel_files):
            filepath = os.path.join(self.price_list_folder, filename)
            
            # File frame
            file_frame = tk.LabelFrame(scroll_inner, text=filename, font=('Arial', 11, 'bold'),
                                       fg='#f39c12', bg='#34495e', padx=10, pady=8)
            file_frame.pack(fill='x', padx=5, pady=5)
            
            # Scan headers
            headers = self.scan_spreadsheet_headers(filepath)
            header_names = ['Auto-Detect'] + [h[1] for h in headers]
            
            if not headers:
                tk.Label(file_frame, text="⚠ Could not read headers from this file",
                        font=('Arial', 10), fg='#e74c3c', bg='#34495e').pack(anchor='w')
                continue
            
            # Show available columns
            col_list = ", ".join([h[1] for h in headers[:8]])
            if len(headers) > 8:
                col_list += f"... (+{len(headers) - 8} more)"
            tk.Label(file_frame, text=f"Found columns: {col_list}",
                    font=('Arial', 8), fg='#95a5a6', bg='#34495e', wraplength=600, justify='left').pack(anchor='w', pady=(0, 5))
            
            # Dropdowns row
            dropdown_frame = tk.Frame(file_frame, bg='#34495e')
            dropdown_frame.pack(fill='x')
            
            # Part number column dropdown
            tk.Label(dropdown_frame, text="Part Number Column:", font=('Arial', 10, 'bold'),
                    fg='#ecf0f1', bg='#34495e').grid(row=0, column=0, sticky='e', padx=(0, 5))
            
            part_var = tk.StringVar(master=self.root, value='Auto-Detect')
            part_combo = ttk.Combobox(dropdown_frame, textvariable=part_var, values=header_names,
                                      width=30, state='readonly')
            part_combo.grid(row=0, column=1, padx=5, pady=3)
            
            # Description column dropdown
            tk.Label(dropdown_frame, text="Description Column:", font=('Arial', 10, 'bold'),
                    fg='#ecf0f1', bg='#34495e').grid(row=1, column=0, sticky='e', padx=(0, 5))
            
            desc_var = tk.StringVar(master=self.root, value='Auto-Detect')
            desc_combo = ttk.Combobox(dropdown_frame, textvariable=desc_var, values=header_names,
                                      width=30, state='readonly')
            desc_combo.grid(row=1, column=1, padx=5, pady=3)
            
            # Price column dropdown
            tk.Label(dropdown_frame, text="Price Column:", font=('Arial', 10, 'bold'),
                    fg='#ecf0f1', bg='#34495e').grid(row=2, column=0, sticky='e', padx=(0, 5))
            
            price_var = tk.StringVar(master=self.root, value='Auto-Detect')
            price_combo = ttk.Combobox(dropdown_frame, textvariable=price_var, values=header_names,
                                       width=30, state='readonly')
            price_combo.grid(row=2, column=1, padx=5, pady=3)
            
            # Load existing mappings if any
            if filename in self.column_mappings:
                saved = self.column_mappings[filename]
                saved_part = saved.get('part_col', 'Auto-Detect')
                saved_desc = saved.get('desc_col', 'Auto-Detect')
                saved_price = saved.get('price_col', 'Auto-Detect')
                if saved_part in header_names:
                    part_var.set(saved_part)
                if saved_desc in header_names:
                    desc_var.set(saved_desc)
                if saved_price in header_names:
                    price_var.set(saved_price)
            
            file_widgets[filename] = {
                'part_var': part_var,
                'desc_var': desc_var,
                'price_var': price_var
            }
        
        # Buttons
        btn_frame = tk.Frame(config_win, bg='#2c3e50')
        btn_frame.pack(pady=15)
        
        def save_and_reload():
            # Save column mappings
            for filename, widgets in file_widgets.items():
                part_choice = widgets['part_var'].get()
                desc_choice = widgets['desc_var'].get()
                price_choice = widgets['price_var'].get()
                
                if part_choice == 'Auto-Detect' and desc_choice == 'Auto-Detect' and price_choice == 'Auto-Detect':
                    # Remove mapping if all are auto
                    if filename in self.column_mappings:
                        del self.column_mappings[filename]
                else:
                    self.column_mappings[filename] = {
                        'part_col': part_choice,
                        'desc_col': desc_choice,
                        'price_col': price_choice
                    }
            
            self.save_settings()
            
            # Unbind mousewheel before destroying
            canvas.unbind_all("<MouseWheel>")
            config_win.destroy()
            
            # Reload price lists with new mappings
            self.load_price_lists()
            
            msg = f"Column settings saved and price lists reloaded.\n\n"
            msg += f"Loaded {len(self.price_lists)} vendor list(s):\n"
            for vendor, parts in self.price_lists.items():
                msg += f"  • {vendor}: {len(parts)} parts\n"
            messagebox.showinfo("Success", msg)
        
        def cancel():
            canvas.unbind_all("<MouseWheel>")
            config_win.destroy()
        
        tk.Button(btn_frame, text="Save & Reload Price Lists", command=save_and_reload,
                 bg='#27ae60', fg='white', font=('Arial', 11, 'bold'), width=25).pack(side='left', padx=5)
        tk.Button(btn_frame, text="Cancel", command=cancel,
                 bg='#95a5a6', fg='white', font=('Arial', 11, 'bold'), width=10).pack(side='left', padx=5)
        
        # Set a reasonable window size
        config_win.update_idletasks()
        height = min(700, 200 + len(excel_files) * 140)
        config_win.geometry(f"700x{height}")
    
    def load_price_lists(self):
        """Load all price lists from the configured folder"""
        if not OPENPYXL_AVAILABLE:
            print(f"Cannot load price lists - openpyxl not available: {OPENPYXL_ERROR}")
            messagebox.showwarning("Warning", f"Cannot load price lists - openpyxl not available.\n\nError: {OPENPYXL_ERROR}")
            return
        
        if not self.price_list_folder or not os.path.exists(self.price_list_folder):
            return
        
        self.price_lists = {}
        loaded_files = []
        failed_files = []
        xls_files_skipped = []
        
        # Find all Excel files in the folder
        for filename in os.listdir(self.price_list_folder):
            if filename.lower().endswith('.xlsx') and not filename.startswith('~$'):
                filepath = os.path.join(self.price_list_folder, filename)
                # Extract vendor name from filename (without extension)
                vendor_name = os.path.splitext(filename)[0].upper()
                
                try:
                    parts = self.load_single_price_list_xlsx(filepath, filename)
                    if parts:
                        self.price_lists[vendor_name] = parts
                        loaded_files.append(f"{vendor_name}: {len(parts)} parts")
                        print(f"Loaded {len(parts)} parts from {vendor_name}")
                    else:
                        failed_files.append(f"{filename}: No parts found (check column headers)")
                        print(f"No parts found in {filename}")
                except Exception as e:
                    failed_files.append(f"{filename}: {str(e)}")
                    print(f"Error loading {filename}: {e}")
            
            elif filename.lower().endswith('.xls') and not filename.startswith('~$'):
                filepath = os.path.join(self.price_list_folder, filename)
                vendor_name = os.path.splitext(filename)[0].upper()
                
                if XLRD_AVAILABLE:
                    try:
                        parts = self.load_single_price_list_xls(filepath, filename)
                        if parts:
                            self.price_lists[vendor_name] = parts
                            loaded_files.append(f"{vendor_name}: {len(parts)} parts")
                            print(f"Loaded {len(parts)} parts from {vendor_name}")
                        else:
                            failed_files.append(f"{filename}: No parts found (check column headers)")
                            print(f"No parts found in {filename}")
                    except Exception as e:
                        failed_files.append(f"{filename}: {str(e)}")
                        print(f"Error loading {filename}: {e}")
                else:
                    xls_files_skipped.append(filename)
                    print(f"Skipped {filename} - xlrd not installed (needed for .xls files)")
        
        self.price_list_loaded = len(self.price_lists) > 0
        print(f"Total vendors loaded: {len(self.price_lists)}")
        
        # Show detailed results
        if failed_files or xls_files_skipped:
            msg = f"Loaded {len(self.price_lists)} price list(s):\n"
            for f in loaded_files:
                msg += f"  ✓ {f}\n"
            
            if failed_files:
                msg += f"\nFailed to load {len(failed_files)} file(s):\n"
                for f in failed_files:
                    msg += f"  ✗ {f}\n"
            
            if xls_files_skipped:
                msg += f"\nSkipped {len(xls_files_skipped)} .xls file(s) (old Excel format):\n"
                for f in xls_files_skipped:
                    msg += f"  ⚠ {f}\n"
                msg += "\nTo load .xls files, either:\n"
                msg += "  • Install xlrd: pip install xlrd\n"
                msg += "  • Or convert files to .xlsx in Excel"
            
            messagebox.showwarning("Price List Loading Results", msg)
    
    def load_single_price_list_xlsx(self, filepath, filename=None):
        """Load a single .xlsx price list file"""
        parts = []
        wb = None
        
        # Check for saved column mappings
        saved_mapping = self.column_mappings.get(filename, {}) if filename else {}
        saved_part_col_name = saved_mapping.get('part_col', 'Auto-Detect')
        saved_desc_col_name = saved_mapping.get('desc_col', 'Auto-Detect')
        saved_price_col_name = saved_mapping.get('price_col', 'Auto-Detect')
        use_saved_part = saved_part_col_name != 'Auto-Detect'
        use_saved_desc = saved_desc_col_name != 'Auto-Detect'
        use_saved_price = saved_price_col_name != 'Auto-Detect'
        
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            
            for sheet in wb.worksheets:
                # Find the header row
                header_row = None
                part_col = None
                desc_col = None
                price_col = None
                
                for row_num, row in enumerate(sheet.iter_rows(min_row=1, max_row=30, values_only=True), 1):
                    if row is None:
                        continue
                    for col_num, cell in enumerate(row):
                        if cell and isinstance(cell, str):
                            cell_stripped = cell.strip()
                            cell_upper = cell_stripped.upper().replace('_', ' ').replace('-', ' ')
                            
                            # Match saved part column by name
                            if use_saved_part and part_col is None:
                                if cell_stripped == saved_part_col_name:
                                    header_row = row_num
                                    part_col = col_num
                            
                            # Match saved desc column by name
                            if use_saved_desc and desc_col is None:
                                if cell_stripped == saved_desc_col_name:
                                    desc_col = col_num
                            
                            # Match saved price column by name
                            if use_saved_price and price_col is None:
                                if cell_stripped == saved_price_col_name:
                                    price_col = col_num
                            
                            # Auto-detect part number column if not using saved
                            if not use_saved_part and part_col is None:
                                if any(x in cell_upper for x in ['PART NO', 'PART NUM', 'PART #', 'PART#', 'PARTNO', 'PARTNUM',
                                                                   'STOCK NUMBER', 'STOCK NUM', 'STOCK NO', 'STOCK#',
                                                                   'ITEM NO', 'ITEM NUM', 'ITEM #', 'ITEM#', 'ITEMNO',
                                                                   'SKU', 'CATALOG', 'CAT NO', 'CAT #', 'CAT NUM',
                                                                   'PRODUCT NO', 'PRODUCT NUM', 'PRODUCT #', 'PRODUCT CODE']):
                                    header_row = row_num
                                    part_col = col_num
                            
                            # Auto-detect description column if not using saved
                            if not use_saved_desc and desc_col is None:
                                if cell_upper in ('DESCRIPTION', 'DESC', 'ITEM DESCRIPTION', 'PART DESCRIPTION', 
                                                  'ITEM', 'PRODUCT', 'PRODUCT NAME', 'NAME', 'ITEM NAME'):
                                    desc_col = col_num
                            
                            # Auto-detect price column if not using saved
                            if not use_saved_price:
                                if any(x in cell_upper for x in ['NET PRICE', 'NET', 'YOUR PRICE', 'DEALER', 'COST',
                                                                  'DISCOUNT', 'AWP', 'SALE', 'ACTUAL']):
                                    price_col = col_num
                                elif price_col is None:
                                    if any(x in cell_upper for x in ['LIST PRICE', 'LIST', 'PRICE', 'MSRP', 'RETAIL', 'UNIT PRICE']):
                                        price_col = col_num
                    
                    if header_row and part_col is not None:
                        break
                
                if not header_row or part_col is None:
                    print(f"  Sheet '{sheet.title}': Could not find part number column")
                    continue
                
                print(f"  Sheet '{sheet.title}': Found header at row {header_row}, part_col={part_col}, desc_col={desc_col}, price_col={price_col}")
                
                # Read data rows
                row_count = 0
                for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
                    if row is None:
                        continue
                    try:
                        part_num = row[part_col] if part_col < len(row) else None
                        desc = row[desc_col] if desc_col is not None and desc_col < len(row) else ''
                        price = row[price_col] if price_col is not None and price_col < len(row) else 0
                        
                        if part_num and str(part_num).strip():
                            part_num_str = str(part_num).strip()
                            desc_str = str(desc).strip() if desc else ''
                            
                            # Parse price
                            try:
                                if isinstance(price, (int, float)):
                                    price_val = float(price)
                                elif price is None:
                                    price_val = 0
                                else:
                                    price_str = str(price).replace('$', '').replace(',', '').strip()
                                    price_val = float(price_str) if price_str else 0
                            except:
                                price_val = 0
                            
                            parts.append({
                                'part_num': part_num_str,
                                'desc': desc_str,
                                'price': price_val
                            })
                            row_count += 1
                    except Exception as e:
                        continue
                
                print(f"  Sheet '{sheet.title}': Loaded {row_count} parts")
        
        finally:
            if wb:
                wb.close()
        
        return parts
    
    def load_single_price_list_xls(self, filepath, filename=None):
        """Load a single .xls (old Excel format) price list file"""
        parts = []
        
        if not XLRD_AVAILABLE:
            raise ImportError("xlrd package required for .xls files")
        
        # Check for saved column mappings
        saved_mapping = self.column_mappings.get(filename, {}) if filename else {}
        saved_part_col_name = saved_mapping.get('part_col', 'Auto-Detect')
        saved_desc_col_name = saved_mapping.get('desc_col', 'Auto-Detect')
        saved_price_col_name = saved_mapping.get('price_col', 'Auto-Detect')
        use_saved_part = saved_part_col_name != 'Auto-Detect'
        use_saved_desc = saved_desc_col_name != 'Auto-Detect'
        use_saved_price = saved_price_col_name != 'Auto-Detect'
        
        wb = xlrd.open_workbook(filepath)
        
        for sheet_idx in range(wb.nsheets):
            sheet = wb.sheet_by_index(sheet_idx)
            
            # Find the header row
            header_row = None
            part_col = None
            desc_col = None
            price_col = None
            
            for row_num in range(min(30, sheet.nrows)):
                for col_num in range(sheet.ncols):
                    cell = sheet.cell_value(row_num, col_num)
                    if cell and isinstance(cell, str):
                        cell_stripped = cell.strip()
                        cell_upper = cell_stripped.upper().replace('_', ' ').replace('-', ' ')
                        
                        # Match saved part column by name
                        if use_saved_part and part_col is None:
                            if cell_stripped == saved_part_col_name:
                                header_row = row_num
                                part_col = col_num
                        
                        # Match saved desc column by name
                        if use_saved_desc and desc_col is None:
                            if cell_stripped == saved_desc_col_name:
                                desc_col = col_num
                        
                        # Match saved price column by name
                        if use_saved_price and price_col is None:
                            if cell_stripped == saved_price_col_name:
                                price_col = col_num
                        
                        # Auto-detect part number column if not using saved
                        if not use_saved_part and part_col is None:
                            if any(x in cell_upper for x in ['PART NO', 'PART NUM', 'PART #', 'PART#', 'PARTNO', 'PARTNUM',
                                                               'STOCK NUMBER', 'STOCK NUM', 'STOCK NO', 'STOCK#',
                                                               'ITEM NO', 'ITEM NUM', 'ITEM #', 'ITEM#', 'ITEMNO',
                                                               'SKU', 'CATALOG', 'CAT NO', 'CAT #', 'CAT NUM',
                                                               'PRODUCT NO', 'PRODUCT NUM', 'PRODUCT #', 'PRODUCT CODE']):
                                header_row = row_num
                                part_col = col_num
                        
                        # Auto-detect description column if not using saved
                        if not use_saved_desc and desc_col is None:
                            if cell_upper in ('DESCRIPTION', 'DESC', 'ITEM DESCRIPTION', 'PART DESCRIPTION',
                                              'ITEM', 'PRODUCT', 'PRODUCT NAME', 'NAME', 'ITEM NAME'):
                                desc_col = col_num
                        
                        # Auto-detect price column if not using saved
                        if not use_saved_price:
                            if any(x in cell_upper for x in ['NET PRICE', 'NET', 'YOUR PRICE', 'DEALER', 'COST',
                                                              'DISCOUNT', 'AWP', 'SALE', 'ACTUAL']):
                                price_col = col_num
                            elif price_col is None:
                                if any(x in cell_upper for x in ['LIST PRICE', 'LIST', 'PRICE', 'MSRP', 'RETAIL', 'UNIT PRICE']):
                                    price_col = col_num
                
                if header_row is not None and part_col is not None:
                    break
            
            if header_row is None or part_col is None:
                print(f"  Sheet '{sheet.name}': Could not find part number column")
                continue
            
            print(f"  Sheet '{sheet.name}': Found header at row {header_row+1}, part_col={part_col}, desc_col={desc_col}, price_col={price_col}")
            
            # Read data rows
            row_count = 0
            for row_num in range(header_row + 1, sheet.nrows):
                try:
                    part_num = sheet.cell_value(row_num, part_col) if part_col < sheet.ncols else None
                    desc = sheet.cell_value(row_num, desc_col) if desc_col is not None and desc_col < sheet.ncols else ''
                    price = sheet.cell_value(row_num, price_col) if price_col is not None and price_col < sheet.ncols else 0
                    
                    if part_num and str(part_num).strip():
                        part_num_str = str(part_num).strip()
                        desc_str = str(desc).strip() if desc else ''
                        
                        # Parse price
                        try:
                            if isinstance(price, (int, float)):
                                price_val = float(price)
                            else:
                                price_str = str(price).replace('$', '').replace(',', '').strip()
                                price_val = float(price_str) if price_str else 0
                        except:
                            price_val = 0
                        
                        parts.append({
                            'part_num': part_num_str,
                            'desc': desc_str,
                            'price': price_val
                        })
                        row_count += 1
                except Exception as e:
                    continue
            
            print(f"  Sheet '{sheet.name}': Loaded {row_count} parts")
        
        return parts
    
    def search_part_in_price_list(self, partial_part_num, vendor_name=None):
        """Search for parts matching a partial part number across ALL price lists"""
        results = []
        partial_upper = partial_part_num.upper().strip()
        
        if not partial_upper:
            return results
        
        # Always search ALL vendors, but put matching vendor results first
        all_vendors = list(self.price_lists.keys())
        
        if vendor_name:
            vendor_upper = vendor_name.upper()
            matching_vendors = [v for v in all_vendors if vendor_upper in v or v in vendor_upper]
            other_vendors = [v for v in all_vendors if v not in matching_vendors]
            vendors_ordered = matching_vendors + other_vendors
        else:
            vendors_ordered = all_vendors
        
        # Search in all vendor lists
        for vendor in vendors_ordered:
            parts = self.price_lists.get(vendor, [])
            for part in parts:
                if partial_upper in part['part_num'].upper():
                    results.append({
                        'vendor': vendor,
                        'part_num': part['part_num'],
                        'desc': part['desc'],
                        'price': part['price']
                    })
                    
                    # Limit results for performance
                    if len(results) >= 50:
                        return results
        
        return results
    
    def get_exact_part(self, part_num, vendor_name=None):
        """Get exact part match from price list - searches ALL lists"""
        part_upper = part_num.upper().strip()
        
        # Always search all vendors
        all_vendors = list(self.price_lists.keys())
        
        if vendor_name:
            vendor_upper = vendor_name.upper()
            matching_vendors = [v for v in all_vendors if vendor_upper in v or v in vendor_upper]
            other_vendors = [v for v in all_vendors if v not in matching_vendors]
            vendors_ordered = matching_vendors + other_vendors
        else:
            vendors_ordered = all_vendors
        
        for vendor in vendors_ordered:
            parts = self.price_lists.get(vendor, [])
            for part in parts:
                if part['part_num'].upper() == part_upper:
                    return {
                        'vendor': vendor,
                        'part_num': part['part_num'],
                        'desc': part['desc'],
                        'price': part['price']
                    }
        
        return None
    
    
    def import_machine_data(self, machine):
        """Import machine data from repair log into service order"""
        # Clear all existing data first
        self.linked_estimate_number = None
        self.linked_label.config(text="No machine selected")
        
        for var in [self.estimate_date_var, self.repair_date_var, self.po_number_var,
                   self.order_number_var, self.bill_to_var, self.ship_to_var,
                   self.address_var, self.contact_var, self.tag_var, self.phone_var,
                   self.mfd1_var, self.model1_var, self.sn1_var,
                   self.mfd2_var, self.model2_var, self.sn2_var, self.power_var,
                   self.engineer_var, self.customer_sig_name_var,
                   self.customer_email_var]:
            var.set('')
        
        self.date_var.set(datetime.now().strftime('%m-%d-%Y'))
        self.sig_date_var.set(datetime.now().strftime('%m-%d-%Y'))
        self.customer_sig_date_var.set(datetime.now().strftime('%m-%d-%Y'))
        self.repair_date_var.set(datetime.now().strftime('%m-%d-%Y'))
        self.plug_var.set('NO')
        self.prepower_var.set('NO')
        self.additional_var.set('NO')
        self.labor_qty_var.set('0')
        self.select_default_labor_rate()

        self.reason_text.delete('1.0', tk.END)
        self.work_text.delete('1.0', tk.END)
        
        # Clear signature
        self.signature_canvas.delete('signature')
        self._signature_strokes = []
        self._signature_canvas_size = None
        
        for entry in self.parts_entries:
            entry['part_num'].set('')
            entry['desc'].set('')
            entry['qty'].set('')
            entry['each'].set('')
        
        self.calculate_totals()
        
        # Now import the new machine data
        self.linked_estimate_number = machine.get('estimate_number', '')
        self.record_id = str(machine.get('record_id', '') or '')
        
        # Update linked label
        self.linked_label.config(text=f"Linked to: {self.linked_estimate_number} - {machine.get('customer_name', '')}")
        
        # Fill in form fields
        self.bill_to_var.set(machine.get('customer_name', ''))
        self.ship_to_var.set(machine.get('customer_name', ''))
        self.tag_var.set(self.linked_estimate_number)
        self.service_location = machine.get('service_location', 'Shop')
        
        # Machine info
        self.mfd1_var.set(machine.get('make', ''))
        self.model1_var.set(machine.get('model', ''))
        self.sn1_var.set(machine.get('serial_number', ''))
        
        # Set estimate date from arrival date
        self.estimate_date_var.set(machine.get('arrival_date', ''))
        
        # Set order number to estimate number
        self.order_number_var.set(self.linked_estimate_number)
        
        # PO number if available
        self.po_number_var.set(machine.get('po_number', ''))
        
        # Check if there's an existing service order for this machine
        self.load_existing_service_order()
    
    def get_service_order_path(self):
        """Path for this service order on the shared drive.

        This used to call os.makedirs() here, which raises when the share is
        offline. The exception escaped the Save button's handler and, in a
        --noconsole build, produced no message at all - the save just looked
        like it did nothing. Directory creation now happens in the save path
        where it can be handled.
        """
        if self.linked_estimate_number and self.service_orders_dir:
            return os.path.join(self.service_orders_dir,
                                f"{self.linked_estimate_number}_service_order.json")
        return None

    def get_local_service_orders_dir(self):
        """Local cache folder that always works, even with the share down."""
        return get_offline_service_orders_dir()

    def get_local_service_order_path(self):
        """Local copy of this service order."""
        if not self.linked_estimate_number:
            return None
        return os.path.join(self.get_local_service_orders_dir(),
                            f"{self.linked_estimate_number}_service_order.json")
    
    @staticmethod
    def _so_saved_stamp(path):
        return read_service_order_stamp(path)

    @staticmethod
    def _so_record_id(path):
        """record_id stored inside a service order file ('' if legacy)."""
        try:
            with open(path, 'r') as f:
                return str(json.load(f).get('record_id', '') or '')
        except Exception:
            return ''

    def _so_belongs_here(self, path):
        """Guard against opening another job's paperwork.

        Service orders are filed by estimate number. If a number was ever
        reused, that alone would hand the wrong customer's signed work to
        the next job, so the record_id inside the file has to agree too.
        Files written before record_id existed are accepted.
        """
        mine = str(getattr(self, 'record_id', '') or '')
        theirs = self._so_record_id(path)
        if not theirs or not mine:
            return True          # legacy file, or no id to check against
        return theirs == mine

    def find_existing_service_order_file(self):
        """Find the newest copy of this service order.

        Checks the shared drive and the local offline cache and returns
        whichever was saved most recently, so a tech who saved offline still
        sees their own work when they reopen the order.
        """
        if not self.linked_estimate_number:
            return None

        candidates = []

        if self.service_orders_dir:
            # Current filename format
            p1 = os.path.join(self.service_orders_dir,
                              f"{self.linked_estimate_number}_service_order.json")
            if os.path.exists(p1) and self._so_belongs_here(p1):
                candidates.append(p1)
            # Old SO_ prefix format (from an interim version)
            p2 = os.path.join(self.service_orders_dir,
                              f"SO_{self.linked_estimate_number}.json")
            if os.path.exists(p2) and self._so_belongs_here(p2):
                candidates.append(p2)

        try:
            local = self.get_local_service_order_path()
            if local and os.path.exists(local) and self._so_belongs_here(local):
                candidates.append(local)
        except Exception:
            pass

        if not candidates:
            return None
        if len(candidates) == 1:
            return candidates[0]

        # Prefer the most recently saved copy; fall back to file mtime
        def sort_key(path):
            stamp = self._so_saved_stamp(path)
            if stamp:
                return (1, stamp)
            try:
                return (0, datetime.fromtimestamp(os.path.getmtime(path)).isoformat())
            except Exception:
                return (0, '')

        return max(candidates, key=sort_key)

    def sync_offline_service_orders(self, silent=True):
        """Upload any locally cached service orders now that the share is up."""
        uploaded, failed = push_offline_service_orders(self.service_orders_dir)
        if uploaded:
            print(f"Synced {uploaded} offline service order(s)")
            if not silent:
                messagebox.showinfo(
                    "Offline Orders Uploaded",
                    f"{uploaded} service order(s) saved while offline have been "
                    "uploaded to the shared drive.")
        if failed:
            print("Offline sync problems: " + "; ".join(failed))
        return uploaded

    def load_existing_service_order(self):
        """Load existing service order if one exists for this machine"""
        path = self.find_existing_service_order_file()
        if not path:
            return
        
        try:
            with open(path, 'r') as f:
                data = json.load(f)
            
            # Load all fields
            self.estimate_date_var.set(data.get('estimate_date', ''))
            self.repair_date_var.set(data.get('repair_date', ''))
            self.po_number_var.set(data.get('po_number', ''))
            self.order_number_var.set(data.get('order_number', ''))
            self.bill_to_var.set(data.get('bill_to', ''))
            self.ship_to_var.set(data.get('ship_to', ''))
            self.address_var.set(data.get('address', ''))
            self.contact_var.set(data.get('contact', ''))
            self.date_var.set(data.get('date', ''))
            self.tag_var.set(data.get('tag', ''))
            self.phone_var.set(data.get('phone', ''))
            self.mfd1_var.set(data.get('mfd1', ''))
            self.model1_var.set(data.get('model1', ''))
            self.sn1_var.set(data.get('sn1', ''))
            self.mfd2_var.set(data.get('mfd2', ''))
            self.model2_var.set(data.get('model2', ''))
            self.sn2_var.set(data.get('sn2', ''))
            self.power_var.set(data.get('power', ''))
            self.plug_var.set(data.get('plug', 'NO'))
            self.prepower_var.set(data.get('prepower', 'NO'))
            self.additional_var.set(data.get('additional', 'NO'))
            self.labor_qty_var.set(data.get('labor_qty', '0'))
            self.set_saved_labor_rate(data.get('labor_rate_name', ''),
                                      data.get('labor_rate', '150.00'))
            self.engineer_var.set(data.get('engineer', ''))
            self.sig_date_var.set(data.get('sig_date', ''))
            
            # Customer signature
            self.customer_sig_name_var.set(data.get('customer_sig_name', ''))
            self.customer_sig_date_var.set(data.get('customer_sig_date', datetime.now().strftime('%m-%d-%Y')))
            self.customer_email_var.set(data.get('customer_email', ''))
            self.service_location = data.get('service_location',
                                             getattr(self, 'service_location', 'Shop'))
            self._signature_strokes = data.get('signature_strokes', [])
            self._signature_canvas_size = data.get('signature_canvas_size') or None
            # Redraw signature on canvas
            self.signature_canvas.delete('signature')
            for stroke in self._signature_strokes:
                for i in range(1, len(stroke)):
                    x1, y1 = stroke[i-1]
                    x2, y2 = stroke[i]
                    self.signature_canvas.create_line(x1, y1, x2, y2,
                        width=2, fill='black', capstyle='round', smooth=True, tags='signature')
            
            # Text fields
            self.reason_text.delete('1.0', tk.END)
            self.reason_text.insert('1.0', data.get('reason', ''))
            self.work_text.delete('1.0', tk.END)
            self.work_text.insert('1.0', data.get('work', ''))
            
            # Parts
            parts = data.get('parts', [])
            for i, entry in enumerate(self.parts_entries):
                if i < len(parts):
                    entry['part_num'].set(parts[i].get('part_num', ''))
                    entry['desc'].set(parts[i].get('desc', ''))
                    entry['qty'].set(parts[i].get('qty', ''))
                    entry['each'].set(parts[i].get('each', ''))
            
            self.calculate_totals()
            
        except Exception as e:
            print(f"Error loading service order: {e}")
    
    def save_service_order(self):
        """Save the service order (never fails silently)."""
        try:
            self._save_service_order_inner()
        except Exception as e:
            # Last-resort net: a --noconsole build shows nothing on an
            # unhandled callback error, which is what made an offline save
            # look like the button did nothing at all.
            import traceback
            traceback.print_exc()
            messagebox.showerror(
                "Save Failed",
                f"Unexpected error while saving the service order:\n\n{e}\n\n"
                "Your work is still on screen. Try again, or export the "
                "customer copy PDF to preserve it.")

    def _set_so_status(self, text, color='#ecf0f1'):
        """Update the service order window's own status line, if it exists."""
        try:
            if getattr(self, 'so_status_label', None):
                self.so_status_label.config(text=text, fg=color)
        except Exception:
            pass

    def _save_service_order_inner(self):
        """Save the service order"""
        if not self.linked_estimate_number:
            messagebox.showwarning("Warning", "Please select a machine from the repair log first.")
            return
        
        path = self.get_service_order_path()
        if not path:
            messagebox.showerror("Error", "Cannot determine save location.")
            return
        
        # Collect all data
        data = {
            'linked_estimate': self.linked_estimate_number,
            'estimate_date': self.estimate_date_var.get(),
            'repair_date': self.repair_date_var.get(),
            'po_number': self.po_number_var.get(),
            'order_number': self.order_number_var.get(),
            'bill_to': self.bill_to_var.get(),
            'ship_to': self.ship_to_var.get(),
            'address': self.address_var.get(),
            'contact': self.contact_var.get(),
            'date': self.date_var.get(),
            'tag': self.tag_var.get(),
            'phone': self.phone_var.get(),
            'mfd1': self.mfd1_var.get(),
            'model1': self.model1_var.get(),
            'sn1': self.sn1_var.get(),
            'mfd2': self.mfd2_var.get(),
            'model2': self.model2_var.get(),
            'sn2': self.sn2_var.get(),
            'power': self.power_var.get(),
            'plug': self.plug_var.get(),
            'prepower': self.prepower_var.get(),
            'reason': self.reason_text.get('1.0', tk.END).strip(),
            'work': self.work_text.get('1.0', tk.END).strip(),
            'additional': self.additional_var.get(),
            'labor_qty': self.labor_qty_var.get(),
            'labor_rate': self.labor_rate_var.get(),
            'labor_rate_name': self.labor_rate_name_var.get(),
            'engineer': self.engineer_var.get(),
            'sig_date': self.sig_date_var.get(),
            'customer_sig_name': self.customer_sig_name_var.get(),
            'customer_sig_date': self.customer_sig_date_var.get(),
            'customer_email': self.customer_email_var.get().strip(),
            'service_location': getattr(self, 'service_location', 'Shop'),
            'record_id': str(getattr(self, 'record_id', '') or ''),
            'signature_strokes': self._signature_strokes,
            'signature_canvas_size': self._get_signature_canvas_size(),
            'parts_total': self.parts_total_var.get(),
            'labor_total': self.labor_total_var.get(),
            'grand_total': self.grand_total_var.get(),
            'last_saved': datetime.now().isoformat(),
            'parts': []
        }
        
        # Collect parts
        for entry in self.parts_entries:
            if entry['part_num'].get() or entry['desc'].get():
                data['parts'].append({
                    'part_num': entry['part_num'].get(),
                    'desc': entry['desc'].get(),
                    'qty': entry['qty'].get(),
                    'each': entry['each'].get(),
                    'total': entry['total'].get()
                })
        
        # --- Always write the local copy first. This never depends on the
        # --- network, so the tech's work is safe even with the share down.
        local_ok = False
        local_err = ""
        try:
            local_path = self.get_local_service_order_path()
            if local_path:
                with open(local_path, 'w') as f:
                    json.dump(data, f, indent=2)
                local_ok = True
        except Exception as e:
            local_err = str(e)
            print(f"Local service order save failed: {e}")

        # --- Then try the shared drive ---
        server_ok = False
        server_err = ""
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, 'w') as f:
                json.dump(data, f, indent=2)
            server_ok = True
        except Exception as e:
            server_err = str(e)
            print(f"Server service order save failed: {e}")

        # --- Update the repair log (also offline-safe on the parent app) ---
        try:
            self.update_repair_log_with_cost()
        except Exception as e:
            print(f"Could not update repair log: {e}")

        # --- Always tell the tech what happened. One message, every time. ---
        if server_ok:
            self._set_so_status(f"Saved to server at {datetime.now().strftime('%H:%M:%S')}",
                                '#27ae60')
            messagebox.showinfo(
                "Saved",
                f"Service order saved for {self.linked_estimate_number}.")
        elif local_ok:
            self._set_so_status(
                f"OFFLINE - saved locally at {datetime.now().strftime('%H:%M:%S')}",
                '#e67e22')
            messagebox.showwarning(
                "Saved Offline",
                f"Service order saved for {self.linked_estimate_number}.\n\n"
                "The shared drive is not reachable right now, so it was saved on "
                "this computer.\n\n"
                "It will upload automatically when the connection comes back. "
                "You can keep working and close the program safely.")
        else:
            self._set_so_status("SAVE FAILED", '#e74c3c')
            messagebox.showerror(
                "Save Failed",
                f"The service order could NOT be saved.\n\n"
                f"Shared drive: {server_err or 'unavailable'}\n"
                f"This computer: {local_err or 'unavailable'}\n\n"
                "Do not close this window. Try again, or export the customer "
                "copy PDF so the work is not lost.")
    
    def update_repair_log_with_cost(self):
        """Update the repair log with service order cost information"""
        try:
            for machine in self.parent_app.active_machines + self.parent_app.shipped_machines:
                if machine.get('estimate_number') == self.linked_estimate_number:
                    machine['service_order_total'] = self.grand_total_var.get()
                    machine['has_service_order'] = True
                    break
            self.parent_app.save_data()
        except Exception as e:
            print(f"Error updating repair log: {e}")
    
    def export_to_pdf(self):
        """Export service order to PDF"""
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Error", f"PDF export requires reportlab.\nError: {REPORTLAB_ERROR}")
            return
        
        if not self.linked_estimate_number:
            messagebox.showwarning("Warning", "Please select a machine first.")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Service Order to PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"ServiceOrder_{self.linked_estimate_number}.pdf"
        )
        
        if not filename:
            return
        
        try:
            c = pdf_canvas.Canvas(filename, pagesize=letter)
            width, height = letter
            
            # Header
            c.setFont("Helvetica-Bold", 24)
            c.drawString(50, height - 50, self.company_name if self.company_name else "Service Order")
            c.setFont("Helvetica-Bold", 18)
            c.drawString(400, height - 50, "Service Order")
            
            c.setFont("Helvetica", 10)
            y = height - 90
            
            # Order info
            c.drawString(50, y, f"Order #: {self.order_number_var.get()}")
            c.drawString(300, y, f"PO #: {self.po_number_var.get()}")
            c.drawString(450, y, f"Date: {self.date_var.get()}")
            y -= 20
            
            c.drawString(50, y, f"Estimate Date: {self.estimate_date_var.get()}")
            c.drawString(300, y, f"Repair Date: {self.repair_date_var.get()}")
            y -= 30
            
            # Customer info
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y, "Customer Information")
            c.setFont("Helvetica", 10)
            y -= 15
            c.drawString(50, y, f"Bill To: {self.bill_to_var.get()}")
            y -= 15
            c.drawString(50, y, f"Ship To: {self.ship_to_var.get()}")
            y -= 15
            c.drawString(50, y, f"Address: {self.address_var.get()}")
            y -= 15
            c.drawString(50, y, f"Contact: {self.contact_var.get()}    Phone: {self.phone_var.get()}")
            y -= 30
            
            # Machine info
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y, "Machine Information")
            c.setFont("Helvetica", 10)
            y -= 15
            c.drawString(50, y, f"MFD: {self.mfd1_var.get()}    Model: {self.model1_var.get()}    S/N: {self.sn1_var.get()}")
            y -= 15
            c.drawString(50, y, f"Input Power: {self.power_var.get()}    Plug Included: {self.plug_var.get()}")
            y -= 30
            
            # Reason for service
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y, "Reason For Service")
            c.setFont("Helvetica", 10)
            y -= 15
            reason = self.reason_text.get('1.0', tk.END).strip()
            for line in reason.split('\n')[:3]:
                c.drawString(50, y, line[:80])
                y -= 12
            y -= 20
            
            # Work performed
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y, "Description of Work")
            c.setFont("Helvetica", 10)
            y -= 15
            work = self.work_text.get('1.0', tk.END).strip()
            for line in work.split('\n')[:5]:
                c.drawString(50, y, line[:80])
                y -= 12
            y -= 20
            
            # Parts table
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y, "Parts")
            y -= 20
            
            c.setFont("Helvetica-Bold", 9)
            c.drawString(50, y, "Part #")
            c.drawString(150, y, "Description")
            c.drawString(380, y, "Qty")
            c.drawString(420, y, "Each")
            c.drawString(480, y, "Total")
            y -= 15
            
            c.setFont("Helvetica", 9)
            for entry in self.parts_entries:
                if entry['part_num'].get() or entry['desc'].get():
                    c.drawString(50, y, entry['part_num'].get()[:15])
                    c.drawString(150, y, entry['desc'].get()[:35])
                    c.drawString(380, y, entry['qty'].get())
                    c.drawString(420, y, f"${entry['each'].get()}")
                    c.drawString(480, y, f"${entry['total'].get()}")
                    y -= 12
                    
                    if y < 150:
                        break
            
            y -= 20
            
            # Totals
            c.setFont("Helvetica-Bold", 11)
            c.drawString(380, y, f"Parts Total: {self.parts_total_var.get()}")
            y -= 15
            rate_name = self.labor_rate_name_var.get().strip()
            if rate_name:
                c.setFont("Helvetica", 10)
                c.drawString(380, y, f"Labor rate: {rate_name}")
                y -= 14
                c.setFont("Helvetica-Bold", 11)
            c.drawString(380, y, f"Labor ({self.labor_qty_var.get()} hrs @ ${self.labor_rate_var.get()}/hr): {self.labor_total_var.get()}")
            y -= 20
            c.setFont("Helvetica-Bold", 14)
            c.drawString(380, y, f"TOTAL: {self.grand_total_var.get()}")
            
            # Signature
            y -= 40
            c.setFont("Helvetica", 10)
            c.drawString(50, y, f"Service Engineer: {self.engineer_var.get()}")
            c.drawString(300, y, f"Date: {self.sig_date_var.get()}")
            
            # Customer signature section
            y -= 40
            c.setFont("Helvetica-Bold", 11)
            c.drawString(50, y, "CUSTOMER ACKNOWLEDGMENT")
            y -= 15
            c.setFont("Helvetica-Oblique", 9)
            c.drawString(50, y, "I acknowledge that the work described above has been completed to my satisfaction.")
            y -= 20
            c.setFont("Helvetica", 10)
            c.drawString(50, y, f"Customer: {self.customer_sig_name_var.get()}")
            c.drawString(300, y, f"Date: {self.customer_sig_date_var.get()}")
            
            # Draw the signature strokes on the PDF
            y -= 15
            c.setFont("Helvetica", 8)
            c.drawString(50, y, "Signature:")
            
            # Draw signature box
            sig_box_top = y - 5
            sig_box_bottom = sig_box_top - 70
            sig_box_left = 50
            sig_box_right = 550
            c.rect(sig_box_left, sig_box_bottom, sig_box_right - sig_box_left, sig_box_top - sig_box_bottom)
            
            # Render signature strokes if present
            if self._signature_strokes:
                # Canvas is roughly 900x140 (based on typical window width)
                # We need to scale to fit in PDF sig box
                canvas_w = self.signature_canvas.winfo_width() or 900
                canvas_h = self.signature_canvas.winfo_height() or 140
                
                pdf_w = sig_box_right - sig_box_left - 10
                pdf_h = (sig_box_top - sig_box_bottom) - 10
                
                scale_x = pdf_w / canvas_w
                scale_y = pdf_h / canvas_h
                # Use smaller scale to fit and preserve aspect
                scale = min(scale_x, scale_y)
                
                c.setLineWidth(1.2)
                for stroke in self._signature_strokes:
                    if len(stroke) < 2:
                        continue
                    for i in range(1, len(stroke)):
                        x1, y1 = stroke[i-1]
                        x2, y2 = stroke[i]
                        # Transform canvas coords -> PDF coords (PDF origin bottom-left, canvas top-left)
                        px1 = sig_box_left + 5 + x1 * scale
                        py1 = sig_box_top - 5 - y1 * scale
                        px2 = sig_box_left + 5 + x2 * scale
                        py2 = sig_box_top - 5 - y2 * scale
                        c.line(px1, py1, px2, py2)
            
            c.save()
            messagebox.showinfo("Exported", f"Service order exported to:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")
    
    # ==================================================================
    # ============  CUSTOMER COPY PDF (signed field report)  ===========
    # ==================================================================

    def _get_signature_canvas_size(self):
        """Return [w, h] of the signature canvas used when signing.

        Falls back to the live widget size, then to a sane default, so the
        signature always scales correctly into the PDF even if the window
        was resized or the order was reloaded from disk.
        """
        size = getattr(self, '_signature_canvas_size', None)
        if size and len(size) == 2 and size[0] > 1 and size[1] > 1:
            return [int(size[0]), int(size[1])]
        try:
            w = self.signature_canvas.winfo_width()
            h = self.signature_canvas.winfo_height()
            if w > 1 and h > 1:
                return [int(w), int(h)]
        except Exception:
            pass
        return [900, 140]

    def has_signature(self):
        """True if the customer has signed this service order."""
        return bool([s for s in (self._signature_strokes or []) if len(s) > 1])

    def _customer_pdf_options(self):
        """Modal dialog: ask what to include on this customer's copy.

        Returns a dict of options, or None if the user cancelled.
        """
        dlg = tk.Toplevel(self.root)
        dlg.title("Customer Copy Options")
        dlg.configure(bg='#2c3e50')
        dlg.resizable(False, False)
        dlg.transient(self.root)

        result = {'ok': False}

        tk.Label(dlg, text="Customer Copy PDF", font=('Arial', 14, 'bold'),
                 fg='#e67e22', bg='#2c3e50').pack(anchor='w', padx=20, pady=(15, 2))
        tk.Label(dlg, text=f"Service Order: {self.linked_estimate_number}",
                 font=('Arial', 10), fg='#bdc3c7', bg='#2c3e50').pack(anchor='w', padx=20)

        # Signature status line
        if self.has_signature():
            sig_txt = "✔ Customer signature captured - it will appear on the PDF."
            sig_fg = '#2ecc71'
        else:
            sig_txt = "⚠ No signature on this order - the PDF will show a blank signature line."
            sig_fg = '#f1c40f'
        tk.Label(dlg, text=sig_txt, font=('Arial', 9, 'italic'),
                 fg=sig_fg, bg='#2c3e50', wraplength=380,
                 justify='left').pack(anchor='w', padx=20, pady=(8, 10))

        opts_frame = tk.Frame(dlg, bg='#2c3e50')
        opts_frame.pack(fill='x', padx=20, pady=5)

        include_pricing = tk.BooleanVar(master=dlg, value=True)
        include_parts = tk.BooleanVar(master=dlg, value=True)
        include_reason = tk.BooleanVar(master=dlg, value=True)

        tk.Checkbutton(opts_frame, text="Include pricing (parts, labor and total)",
                       variable=include_pricing, bg='#2c3e50', fg='white',
                       selectcolor='#1a252f', activebackground='#2c3e50',
                       activeforeground='white', font=('Arial', 10),
                       anchor='w').pack(fill='x', pady=2)
        tk.Checkbutton(opts_frame, text="Include parts / materials list",
                       variable=include_parts, bg='#2c3e50', fg='white',
                       selectcolor='#1a252f', activebackground='#2c3e50',
                       activeforeground='white', font=('Arial', 10),
                       anchor='w').pack(fill='x', pady=2)
        tk.Checkbutton(opts_frame, text="Include reason for service / complaint",
                       variable=include_reason, bg='#2c3e50', fg='white',
                       selectcolor='#1a252f', activebackground='#2c3e50',
                       activeforeground='white', font=('Arial', 10),
                       anchor='w').pack(fill='x', pady=2)

        tk.Label(dlg, text="Uncheck pricing for a work-completed record with no dollar amounts.",
                 font=('Arial', 8, 'italic'), fg='#95a5a6', bg='#2c3e50',
                 wraplength=380, justify='left').pack(anchor='w', padx=20, pady=(4, 8))

        # --- Email destination ---
        email_box = tk.LabelFrame(dlg, text="  Send to customer  ", font=('Arial', 10, 'bold'),
                                  fg='#3498db', bg='#2c3e50', bd=2)
        email_box.pack(fill='x', padx=20, pady=(2, 10))

        email_var = tk.StringVar(master=dlg, value=self.customer_email_var.get().strip())
        send_email = tk.BooleanVar(master=dlg, value=bool(email_var.get()))

        tk.Checkbutton(email_box, text="Open an Outlook draft with the PDF attached",
                       variable=send_email, bg='#2c3e50', fg='white',
                       selectcolor='#1a252f', activebackground='#2c3e50',
                       activeforeground='white', font=('Arial', 10),
                       anchor='w').pack(fill='x', padx=8, pady=(6, 2))

        email_entry_row = tk.Frame(email_box, bg='#2c3e50')
        email_entry_row.pack(fill='x', padx=8, pady=(0, 8))
        tk.Label(email_entry_row, text="To:", font=('Arial', 10, 'bold'),
                 fg='white', bg='#2c3e50').pack(side='left', padx=(0, 6))
        tk.Entry(email_entry_row, textvariable=email_var, width=34,
                 font=('Arial', 10)).pack(side='left', fill='x', expand=True)

        btn_frame = tk.Frame(dlg, bg='#2c3e50')
        btn_frame.pack(fill='x', padx=20, pady=(0, 15))

        def on_ok():
            addr = email_var.get().strip()
            if send_email.get():
                if not addr:
                    messagebox.showwarning("Email Address Needed",
                                           "Enter the customer's email address, or uncheck "
                                           "the Outlook option.", parent=dlg)
                    return
                if not self._is_valid_email(addr):
                    if not messagebox.askyesno(
                            "Check Address",
                            f"'{addr}' doesn't look like a valid email address.\n\n"
                            "Use it anyway?", parent=dlg):
                        return
            result['ok'] = True
            result['pricing'] = include_pricing.get()
            result['parts'] = include_parts.get()
            result['reason'] = include_reason.get()
            result['send_email'] = send_email.get()
            result['email'] = addr
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        tk.Button(btn_frame, text="Continue →", command=on_ok, bg='#e67e22', fg='white',
                  font=('Arial', 10, 'bold'), width=14).pack(side='right', padx=5)
        tk.Button(btn_frame, text="Cancel", command=on_cancel, bg='#7f8c8d', fg='white',
                  font=('Arial', 10, 'bold'), width=10).pack(side='right', padx=5)

        dlg.protocol("WM_DELETE_WINDOW", on_cancel)
        dlg.update_idletasks()
        # Center over the service order window
        try:
            px = self.root.winfo_rootx() + (self.root.winfo_width() // 2) - (dlg.winfo_width() // 2)
            py = self.root.winfo_rooty() + 120
            dlg.geometry(f"+{max(px, 0)}+{max(py, 0)}")
        except Exception:
            pass
        dlg.grab_set()
        self.root.wait_window(dlg)

        return result if result.get('ok') else None

    def export_customer_copy_pdf(self):
        """Generate a clean, customer-facing PDF of the signed service order."""
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Error",
                                 f"PDF export requires reportlab.\nError: {REPORTLAB_ERROR}")
            return

        if not self.linked_estimate_number:
            messagebox.showwarning("Warning", "Please select a machine from the repair log first.")
            return

        # Warn (but allow) if nothing has been signed yet
        if not self.has_signature():
            if not messagebox.askyesno(
                    "No Signature",
                    "This service order has not been signed by the customer.\n\n"
                    "The PDF will be generated with a blank signature line.\n\n"
                    "Continue anyway?"):
                return

        opts = self._customer_pdf_options()
        if not opts:
            return

        safe_customer = "".join(
            ch for ch in (self.bill_to_var.get() or "").strip()
            if ch.isalnum() or ch in (' ', '-', '_')
        ).strip().replace(' ', '_')[:30]

        default_name = f"ServiceOrder_{self.linked_estimate_number}"
        if safe_customer:
            default_name += f"_{safe_customer}"
        default_name += "_CustomerCopy.pdf"

        initial_dir = self.service_orders_dir if self.service_orders_dir else None

        filename = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save Customer Copy PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=default_name,
            initialdir=initial_dir
        )

        if not filename:
            return

        try:
            pages = self._render_customer_copy_pdf(filename, opts)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create customer copy PDF:\n{e}")
            return

        # Keep the address the engineer confirmed at export time on the order
        if opts.get('email') and opts['email'] != self.customer_email_var.get().strip():
            self.customer_email_var.set(opts['email'])

        page_txt = f"{pages} page{'s' if pages != 1 else ''}"

        # Hand off to Outlook when an address was given
        if opts.get('send_email') and opts.get('email'):
            ok, method, detail = self.compose_outlook_email(
                filename, opts['email'], include_pricing=opts.get('pricing', True))

            if ok and method in ('com', 'exe'):
                messagebox.showinfo(
                    "Outlook Draft Ready",
                    f"Customer copy saved ({page_txt}):\n{filename}\n\n"
                    f"{detail}\nAddressed to: {opts['email']}\n\n"
                    "Review the message in Outlook, then send it.")
            elif ok:
                messagebox.showwarning(
                    "Draft Opened - Attach Manually",
                    f"Customer copy saved ({page_txt}):\n{filename}\n\n{detail}")
            else:
                messagebox.showerror("Outlook Unavailable", detail)
                if messagebox.askyesno("Open PDF?",
                                       "Open the customer copy PDF so you can attach it "
                                       "to an email yourself?"):
                    self._open_file_externally(filename)
            return

        # No email requested - just offer to open the file
        if messagebox.askyesno(
                "Customer Copy Created",
                f"Customer copy saved ({page_txt}):\n\n{filename}\n\nOpen it now?"):
            self._open_file_externally(filename)

    def _open_file_externally(self, path):
        """Open a file with the system default application."""
        try:
            os.startfile(path)  # Windows
        except AttributeError:
            try:
                import subprocess
                subprocess.run(['open', path])  # macOS
            except Exception:
                try:
                    import subprocess
                    subprocess.run(['xdg-open', path])  # Linux
                except Exception as e:
                    messagebox.showinfo("Info", f"Could not open the file automatically.\n{e}")
        except Exception as e:
            messagebox.showinfo("Info", f"Could not open the file automatically.\n{e}")

    # ------------------- email helpers / Outlook -------------------

    @staticmethod
    def _is_valid_email(addr):
        """Loose but practical check. Accepts several addresses separated by ; or ,"""
        import re
        addr = (addr or "").strip()
        if not addr or len(addr) > 1000:
            return False
        parts = [p.strip() for p in re.split(r"[;,]", addr) if p.strip()]
        if not parts:
            return False
        pattern = r"^[^@\s,;]+@[^@\s,;]+\.[A-Za-z]{2,}$"
        return all(len(p) <= 254 and re.match(pattern, p) for p in parts)

    @staticmethod
    def _normalize_recipients(addr):
        """Outlook wants recipients separated by semicolons."""
        import re
        parts = [p.strip() for p in re.split(r"[;,]", addr or "") if p.strip()]
        return "; ".join(parts)

    def _build_email_subject(self):
        order_ref = self.order_number_var.get().strip() or self.linked_estimate_number
        bits = [f"Service Order {order_ref}"]
        if self.company_name:
            bits.append(self.company_name)
        if self.repair_date_var.get().strip():
            bits.append(f"Completed {self.repair_date_var.get().strip()}")
        return " - ".join(bits)

    def _build_email_body(self, include_pricing=True):
        """Plain-text draft body. The engineer edits it before sending."""
        contact = self.contact_var.get().strip()
        # Use just the first name if the contact is "Name, Title" or "First Last"
        greeting_name = ""
        if contact:
            greeting_name = contact.split(',')[0].strip().split(' ')[0].strip()

        order_ref = self.order_number_var.get().strip() or self.linked_estimate_number

        unit = self.model1_var.get().strip()
        if self.mfd1_var.get().strip():
            unit = f"{self.mfd1_var.get().strip()} {unit}".strip()
        if self.sn1_var.get().strip():
            unit += f" (S/N {self.sn1_var.get().strip()})"

        lines = []
        lines.append(f"Hello {greeting_name}," if greeting_name else "Hello,")
        lines.append("")

        intro = "Attached is your copy of the completed service order"
        if unit.strip():
            intro += f" for the {unit}"
        if self.repair_date_var.get().strip():
            intro += f", completed on {self.repair_date_var.get().strip()}"
        intro += "."
        lines.append(intro)
        lines.append("")

        lines.append(f"Service Order #: {order_ref}")
        if self.po_number_var.get().strip():
            lines.append(f"PO #: {self.po_number_var.get().strip()}")
        if unit.strip():
            lines.append(f"Equipment: {unit}")
        if include_pricing and self.grand_total_var.get().strip():
            lines.append(f"Total: {self.grand_total_var.get().strip()}")
        lines.append("")

        if self.has_signature() or self.customer_sig_name_var.get().strip():
            lines.append("The attached PDF includes the signed customer acknowledgment "
                         "of the work performed.")
        else:
            lines.append("The attached PDF details the work performed.")

        if self.additional_var.get() == "YES":
            lines.append("")
            lines.append("Please note this repair may require additional parts and labor.")

        lines.append("")
        lines.append("Please let us know if you have any questions.")
        lines.append("")
        lines.append("Thank you,")
        if self.engineer_var.get().strip():
            lines.append(self.engineer_var.get().strip())
        if self.company_name:
            lines.append(self.company_name)

        return "\n".join(lines)

    def _find_outlook_exe(self):
        """Locate OUTLOOK.EXE on Windows. Returns a path or None."""
        try:
            import winreg
        except ImportError:
            return None

        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for sub in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
                        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"):
                try:
                    with winreg.OpenKey(root, sub) as key:
                        path, _ = winreg.QueryValueEx(key, None)
                        if path and os.path.exists(path):
                            return path
                except Exception:
                    continue

        # Fall back to the usual install locations
        candidates = []
        for base in (os.environ.get('ProgramFiles', r'C:\Program Files'),
                     os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)')):
            if not base:
                continue
            candidates.append(os.path.join(base, 'Microsoft Office', 'root', 'Office16', 'OUTLOOK.EXE'))
            for ver in ('Office16', 'Office15', 'Office14'):
                candidates.append(os.path.join(base, 'Microsoft Office', ver, 'OUTLOOK.EXE'))
        for path in candidates:
            if os.path.exists(path):
                return path
        return None

    def compose_outlook_email(self, pdf_path, to_addr, include_pricing=True):
        """Open a new Outlook message with the customer copy attached.

        Tries, in order:
          1. Outlook COM automation (pywin32) - attaches the PDF directly.
          2. OUTLOOK.EXE command line (/c ipm.note /m ... /a ...) - also attaches.
          3. A plain mailto: draft plus opening the PDF's folder to drag it in.

        Returns (ok, method, detail).
        """
        import urllib.parse

        pdf_path = os.path.abspath(pdf_path)
        to_addr = self._normalize_recipients(to_addr)
        subject = self._build_email_subject()
        body = self._build_email_body(include_pricing)
        errors = []

        # ---- 1. COM automation -------------------------------------
        try:
            try:
                import pythoncom
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
            except ImportError:
                pass

            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = to_addr
            mail.Subject = subject
            mail.Body = body
            mail.Attachments.Add(pdf_path)
            mail.Display(False)  # show the draft, don't block
            try:
                outlook.ActiveWindow().Activate()
            except Exception:
                pass
            return True, 'com', 'Outlook draft opened with the PDF attached.'
        except ImportError as e:
            errors.append(f"pywin32 not installed ({e})")
        except Exception as e:
            errors.append(f"Outlook automation: {e}")

        # ---- 2. OUTLOOK.EXE command line ---------------------------
        try:
            exe = self._find_outlook_exe()
            if exe:
                import subprocess
                # /m takes a mailto-style string: comma-separated, no spaces
                cli_to = ",".join(p.strip() for p in to_addr.split(';') if p.strip())
                m_arg = (f"{cli_to}?subject={urllib.parse.quote(subject)}"
                         f"&body={urllib.parse.quote(body)}")
                subprocess.Popen([exe, '/c', 'ipm.note', '/m', m_arg, '/a', pdf_path])
                return True, 'exe', 'Outlook draft opened with the PDF attached.'
            errors.append("OUTLOOK.EXE not found")
        except Exception as e:
            errors.append(f"Outlook command line: {e}")

        # ---- 3. mailto: + open the folder --------------------------
        try:
            # RFC 6068 mailto: takes comma-separated recipients, not semicolons
            mailto_to = ",".join(p.strip() for p in to_addr.split(';') if p.strip())
            mailto = (f"mailto:{mailto_to}?subject={urllib.parse.quote(subject)}"
                      f"&body={urllib.parse.quote(body)}")
            if len(mailto) > 1800:  # some handlers choke on very long mailto URLs
                mailto = f"mailto:{mailto_to}?subject={urllib.parse.quote(subject)}"
            self._open_file_externally(mailto)
            folder = os.path.dirname(pdf_path)
            if folder and os.path.isdir(folder):
                try:
                    self._open_file_externally(folder)
                except Exception:
                    pass
            detail = ("Opened a draft in your default mail program - the PDF could not be "
                      "attached automatically, so drag it in from the folder that just opened."
                      "\n\nWhy: " + "; ".join(errors[:2]))
            return True, 'mailto', detail
        except Exception as e:
            errors.append(f"mailto: {e}")

        return False, 'none', "Could not open an email draft.\n\n" + "\n".join(errors)

    # ---------------- PDF rendering helpers ----------------

    def _render_customer_copy_pdf(self, filename, opts):
        """Render the customer copy. Two passes so footers can say 'Page X of Y'."""
        import io

        # Pass 1: render to memory purely to count pages
        buf = io.BytesIO()
        total = self._draw_customer_copy(buf, opts, total_pages=None)
        # Pass 2: render for real, now that the page count is known
        self._draw_customer_copy(filename, opts, total_pages=total)
        buf.close()
        return total

    def _draw_customer_copy(self, target, opts, total_pages=None):
        """Draw the whole document to `target` (path or file-like). Returns page count."""
        c = pdf_canvas.Canvas(target, pagesize=letter)
        width, height = letter

        LEFT = 50
        RIGHT = width - 50
        CONTENT_W = RIGHT - LEFT
        TOP = height - 50
        BOTTOM = 72  # room for the footer

        state = {'y': TOP, 'page': 1}

        company = self.company_name if self.company_name else "Service Order"
        order_ref = self.order_number_var.get() or self.linked_estimate_number

        # ---- small drawing utilities -------------------------------

        def wrap(text, font, size, max_w):
            """Word-wrap `text` to max_w, honouring existing newlines."""
            out = []
            for raw_line in str(text).split('\n'):
                raw_line = raw_line.rstrip()
                if not raw_line.strip():
                    out.append('')
                    continue
                words = raw_line.split()
                line = ''
                for word in words:
                    trial = f"{line} {word}".strip()
                    if c.stringWidth(trial, font, size) <= max_w:
                        line = trial
                    else:
                        if line:
                            out.append(line)
                        # A single word longer than the line: hard-break it
                        while c.stringWidth(word, font, size) > max_w and len(word) > 1:
                            cut = len(word)
                            while cut > 1 and c.stringWidth(word[:cut], font, size) > max_w:
                                cut -= 1
                            out.append(word[:cut])
                            word = word[cut:]
                        line = word
                if line:
                    out.append(line)
            return out or ['']

        def draw_header():
            y = TOP
            if self.company_name:
                c.setFont("Helvetica-Bold", 20)
                c.setFillColorRGB(0.13, 0.35, 0.55)
                title = self.company_name
                while title and c.stringWidth(title, "Helvetica-Bold", 20) > 330:
                    title = title[:-1]
                c.drawString(LEFT, y, title)
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica-Bold", 16)
            c.drawRightString(RIGHT, y, "SERVICE ORDER")
            c.setFont("Helvetica-Bold", 9)
            c.setFillColorRGB(0.85, 0.42, 0.10)
            c.drawRightString(RIGHT, y - 14, "CUSTOMER COPY")
            c.setFillColorRGB(0, 0, 0)
            if state['page'] > 1:
                c.setFont("Helvetica", 8)
                c.setFillColorRGB(0.45, 0.45, 0.45)
                c.drawString(LEFT, y - 14, f"Service Order {order_ref}  (continued)")
                c.setFillColorRGB(0, 0, 0)
            y -= 26
            c.setLineWidth(1.5)
            c.setStrokeColorRGB(0.13, 0.35, 0.55)
            c.line(LEFT, y, RIGHT, y)
            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(1)
            return y - 20

        def draw_footer():
            c.setFont("Helvetica", 7.5)
            c.setFillColorRGB(0.45, 0.45, 0.45)
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0.8, 0.8, 0.8)
            c.line(LEFT, BOTTOM - 8, RIGHT, BOTTOM - 8)
            c.setStrokeColorRGB(0, 0, 0)

            mid_txt = "Customer Copy - retain for your records"
            mid_w = c.stringWidth(mid_txt, "Helvetica", 7.5)
            mid_x = (LEFT + RIGHT) / 2

            if total_pages:
                page_txt = f"Page {state['page']} of {total_pages}"
            else:
                page_txt = f"Page {state['page']}"

            # Truncate the outer texts so nothing collides with the centre text
            left_txt = company
            left_max = (mid_x - mid_w / 2) - LEFT - 12
            while left_txt and c.stringWidth(left_txt, "Helvetica", 7.5) > left_max:
                left_txt = left_txt[:-1]

            c.drawString(LEFT, BOTTOM - 20, left_txt)
            c.drawCentredString(mid_x, BOTTOM - 20, mid_txt)
            c.drawRightString(RIGHT, BOTTOM - 20, page_txt)
            c.setFillColorRGB(0, 0, 0)
            c.setLineWidth(1)

        def new_page():
            draw_footer()
            c.showPage()
            state['page'] += 1
            state['y'] = draw_header()

        def need(space):
            """Ensure `space` points remain on the page; break if not."""
            if state['y'] - space < BOTTOM:
                new_page()
                return True
            return False

        def section(title):
            need(34)
            c.setFillColorRGB(0.93, 0.94, 0.95)
            c.rect(LEFT, state['y'] - 15, CONTENT_W, 17, stroke=0, fill=1)
            c.setFillColorRGB(0.13, 0.35, 0.55)
            c.setFont("Helvetica-Bold", 10.5)
            c.drawString(LEFT + 5, state['y'] - 11, title.upper())
            c.setFillColorRGB(0, 0, 0)
            state['y'] -= 26

        def field_row(pairs, gap=14):
            """pairs = list of (label, value); laid out evenly across the width."""
            pairs = [(l, v) for l, v in pairs if str(v).strip()]
            if not pairs:
                return
            need(gap + 2)
            col_w = CONTENT_W / len(pairs)
            for i, (label, value) in enumerate(pairs):
                x = LEFT + i * col_w
                c.setFont("Helvetica-Bold", 8.5)
                c.setFillColorRGB(0.4, 0.4, 0.4)
                c.drawString(x, state['y'], f"{label}:")
                lw = c.stringWidth(f"{label}: ", "Helvetica-Bold", 8.5)
                c.setFillColorRGB(0, 0, 0)
                c.setFont("Helvetica", 9.5)
                val = str(value)
                # Trim to fit the column
                while val and c.stringWidth(val, "Helvetica", 9.5) > (col_w - lw - 8):
                    val = val[:-1]
                c.drawString(x + lw, state['y'], val)
            state['y'] -= gap

        def paragraph(text, font="Helvetica", size=9.5, leading=12, indent=2):
            text = (text or "").strip()
            if not text:
                text = "-"
            for line in wrap(text, font, size, CONTENT_W - indent - 4):
                need(leading)
                c.setFont(font, size)
                c.drawString(LEFT + indent, state['y'], line)
                state['y'] -= leading
            state['y'] -= 4

        # ---- gather data -------------------------------------------

        parts = []
        for entry in self.parts_entries:
            pn = entry['part_num'].get().strip()
            desc = entry['desc'].get().strip()
            if pn or desc:
                parts.append({
                    'part_num': pn,
                    'desc': desc,
                    'qty': entry['qty'].get().strip(),
                    'each': entry['each'].get().strip(),
                    'total': entry['total'].get().strip(),
                })

        show_pricing = bool(opts.get('pricing', True))

        # ---- page 1 -------------------------------------------------

        state['y'] = draw_header()

        # Order reference block
        field_row([
            ("Service Order #", order_ref),
            ("PO #", self.po_number_var.get()),
            ("Tag / Estimate #", self.tag_var.get() or self.linked_estimate_number),
        ])
        field_row([
            ("Service Date", self.date_var.get()),
            ("Repair Completed", self.repair_date_var.get()),
            ("Estimate Completed", self.estimate_date_var.get()),
        ])
        field_row([
            ("Service Location", getattr(self, 'service_location', 'Shop')),
        ])
        state['y'] -= 6

        # Customer
        cust_fields = [self.bill_to_var, self.ship_to_var, self.address_var,
                       self.contact_var, self.phone_var]
        if any(v.get().strip() for v in cust_fields):
            section("Customer Information")
            field_row([("Bill To", self.bill_to_var.get()), ("Ship To", self.ship_to_var.get())])
            if self.address_var.get().strip():
                field_row([("Address", self.address_var.get())])
            field_row([("Contact", self.contact_var.get()), ("Phone", self.phone_var.get())])
            state['y'] -= 6

        # Equipment
        equip_fields = [self.mfd1_var, self.model1_var, self.sn1_var,
                        self.mfd2_var, self.model2_var, self.sn2_var, self.power_var]
        if any(v.get().strip() for v in equip_fields):
            section("Equipment Serviced")
            field_row([
                ("Manufacturer", self.mfd1_var.get()),
                ("Model", self.model1_var.get()),
                ("Serial #", self.sn1_var.get()),
            ])
            if (self.mfd2_var.get().strip() or self.model2_var.get().strip()
                    or self.sn2_var.get().strip()):
                field_row([
                    ("Manufacturer", self.mfd2_var.get()),
                    ("Model", self.model2_var.get()),
                    ("Serial #", self.sn2_var.get()),
                ])
            field_row([
                ("Input Power", self.power_var.get()),
                ("Plug Included", self.plug_var.get()),
                ("Pre-Power Check", self.prepower_var.get()),
            ])
            state['y'] -= 6

        # Reason for service
        if opts.get('reason', True) and self.reason_text.get('1.0', tk.END).strip():
            section("Reason For Service / Complaint")
            paragraph(self.reason_text.get('1.0', tk.END))

        # Work performed - the headline of the customer copy
        section("Description of Work Performed")
        paragraph(self.work_text.get('1.0', tk.END))

        # Parts table
        if opts.get('parts', True) and parts:
            section("Parts & Materials")

            if show_pricing:
                cols = [
                    ("Part #", LEFT + 2, 'l', 95),
                    ("Description", LEFT + 100, 'l', 245),
                    ("Qty", LEFT + 355, 'r', 35),
                    ("Each", LEFT + 425, 'r', 60),
                    ("Total", RIGHT - 2, 'r', 65),
                ]
            else:
                cols = [
                    ("Part #", LEFT + 2, 'l', 120),
                    ("Description", LEFT + 130, 'l', 320),
                    ("Qty", RIGHT - 2, 'r', 40),
                ]

            def parts_header():
                need(20)
                c.setFont("Helvetica-Bold", 8.5)
                c.setFillColorRGB(0.25, 0.25, 0.25)
                for title, x, align, _w in cols:
                    if align == 'r':
                        c.drawRightString(x, state['y'], title)
                    else:
                        c.drawString(x, state['y'], title)
                c.setFillColorRGB(0, 0, 0)
                state['y'] -= 4
                c.setLineWidth(0.6)
                c.line(LEFT, state['y'], RIGHT, state['y'])
                state['y'] -= 11

            parts_header()

            def fit(text, max_w, font="Helvetica", size=8.5):
                text = str(text)
                while text and c.stringWidth(text, font, size) > max_w:
                    text = text[:-1]
                return text

            shade = False
            for p in parts:
                if need(13):
                    parts_header()
                if shade:
                    c.setFillColorRGB(0.965, 0.965, 0.965)
                    c.rect(LEFT, state['y'] - 3, CONTENT_W, 12, stroke=0, fill=1)
                    c.setFillColorRGB(0, 0, 0)
                shade = not shade

                c.setFont("Helvetica", 8.5)
                if show_pricing:
                    values = [
                        p['part_num'], p['desc'], p['qty'],
                        f"${p['each']}" if p['each'] else "",
                        f"${p['total']}" if p['total'] else "",
                    ]
                else:
                    values = [p['part_num'], p['desc'], p['qty']]

                for (title, x, align, w), val in zip(cols, values):
                    val = fit(val, w)
                    if align == 'r':
                        c.drawRightString(x, state['y'], val)
                    else:
                        c.drawString(x, state['y'], val)
                state['y'] -= 12

            state['y'] -= 4
            c.setLineWidth(0.6)
            c.line(LEFT, state['y'], RIGHT, state['y'])
            state['y'] -= 12

        # Totals - skipped entirely when there is nothing to charge for
        def _num(s):
            try:
                return float(str(s).replace('$', '').replace(',', '').strip() or 0)
            except Exception:
                return 0.0

        has_charges = bool(parts) or _num(self.labor_qty_var.get()) > 0 \
            or _num(self.grand_total_var.get()) > 0

        if show_pricing and has_charges:
            need(84)
            box_left = RIGHT - 250
            c.setFont("Helvetica", 10)
            c.drawString(box_left, state['y'], "Parts Total:")
            c.drawRightString(RIGHT - 5, state['y'], self.parts_total_var.get() or "$0.00")
            state['y'] -= 15

            labor_label = self._labor_line_label()
            hours = (f" ({self.labor_qty_var.get()} hr @ ${self.labor_rate_var.get()}/hr)"
                     if self.labor_qty_var.get() else "")
            labor_desc = f"{labor_label}{hours}:"
            amount = self.labor_total_var.get() or "$0.00"
            c.setFont("Helvetica", 10)
            room = (RIGHT - 5 - box_left) - c.stringWidth(amount, "Helvetica", 10) - 10
            if c.stringWidth(labor_desc, "Helvetica", 10) > room:
                # Long rate name: put it on its own line above the hours
                c.drawString(box_left, state['y'], labor_label)
                state['y'] -= 13
                labor_desc = f"   {hours.strip()}:" if hours else "   Labor:"
            c.drawString(box_left, state['y'], labor_desc)
            c.drawRightString(RIGHT - 5, state['y'], amount)
            state['y'] -= 8

            c.setLineWidth(0.8)
            c.line(box_left, state['y'], RIGHT, state['y'])
            state['y'] -= 16

            c.setFillColorRGB(0.13, 0.35, 0.55)
            c.setFont("Helvetica-Bold", 13)
            c.drawString(box_left, state['y'], "TOTAL:")
            c.drawRightString(RIGHT - 5, state['y'], self.grand_total_var.get() or "$0.00")
            c.setFillColorRGB(0, 0, 0)
            state['y'] -= 22
        elif opts.get('parts', True) and parts:
            state['y'] -= 4

        # Additional work notice
        if self.additional_var.get() == "YES":
            need(24)
            c.setFillColorRGB(1.0, 0.97, 0.88)
            c.rect(LEFT, state['y'] - 6, CONTENT_W, 18, stroke=0, fill=1)
            c.setFillColorRGB(0.6, 0.35, 0.0)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(LEFT + 6, state['y'],
                         "NOTE: This repair may require additional parts and labor.")
            c.setFillColorRGB(0, 0, 0)
            state['y'] -= 26

        # ---- signature block: keep it whole on one page --------------

        SIG_BLOCK_H = 186
        need(SIG_BLOCK_H)

        section("Customer Acknowledgment")

        # Reference line - keeps this block meaningful if it lands on its own page
        ref_bits = [f"Service Order {order_ref}"]
        if self.model1_var.get().strip():
            unit = self.model1_var.get().strip()
            if self.sn1_var.get().strip():
                unit += f"  S/N {self.sn1_var.get().strip()}"
            ref_bits.append(unit)
        if self.repair_date_var.get().strip():
            ref_bits.append(f"Repair completed {self.repair_date_var.get().strip()}")
        ref_line = "   -   ".join(ref_bits)
        while ref_line and c.stringWidth(ref_line, "Helvetica-Bold", 8.5) > CONTENT_W - 4:
            ref_line = ref_line[:-1]
        c.setFont("Helvetica-Bold", 8.5)
        c.setFillColorRGB(0.35, 0.35, 0.35)
        c.drawString(LEFT + 2, state['y'], ref_line)
        c.setFillColorRGB(0, 0, 0)
        state['y'] -= 15

        c.setFont("Helvetica-Oblique", 9)
        c.drawString(LEFT + 2, state['y'],
                     "I acknowledge that the work described above has been completed "
                     "to my satisfaction.")
        state['y'] -= 20

        # Signature box
        sig_box_h = 72
        sig_box_top = state['y']
        sig_box_bottom = sig_box_top - sig_box_h
        c.setLineWidth(0.8)
        c.setStrokeColorRGB(0.55, 0.55, 0.55)
        c.rect(LEFT, sig_box_bottom, CONTENT_W, sig_box_h)
        c.setStrokeColorRGB(0, 0, 0)

        c.setFont("Helvetica", 7)
        c.setFillColorRGB(0.5, 0.5, 0.5)
        c.drawString(LEFT + 4, sig_box_top - 10, "CUSTOMER SIGNATURE")
        c.setFillColorRGB(0, 0, 0)

        strokes = [s for s in (self._signature_strokes or []) if len(s) > 1]
        if strokes:
            canvas_w, canvas_h = self._get_signature_canvas_size()

            # Tight bounding box of the actual ink, so a small signature in a
            # wide canvas still renders at a readable size.
            xs = [pt[0] for s in strokes for pt in s]
            ys = [pt[1] for s in strokes for pt in s]
            min_x, max_x = min(xs), max(xs)
            min_y, max_y = min(ys), max(ys)
            ink_w = max(max_x - min_x, 1)
            ink_h = max(max_y - min_y, 1)

            pad = 8
            avail_w = CONTENT_W - 2 * pad
            avail_h = sig_box_h - 22 - pad

            # canvas_fit renders the signature at exactly the proportion it was
            # signed at; bbox_fit fills the box. Take the fill, but never blow a
            # small signature up more than 2.5x its on-pad proportion.
            canvas_fit = min(avail_w / canvas_w, avail_h / canvas_h)
            bbox_fit = min(avail_w / ink_w, avail_h / ink_h)
            scale = max(min(bbox_fit, canvas_fit * 2.5), 0.05)

            draw_h = ink_h * scale
            origin_x = LEFT + pad
            origin_y = sig_box_bottom + pad + (avail_h - draw_h) / 2 + draw_h

            c.setLineWidth(1.3)
            c.setLineCap(1)
            c.setLineJoin(1)
            for stroke in strokes:
                path = c.beginPath()
                first = True
                for (sx, sy) in stroke:
                    px = origin_x + (sx - min_x) * scale
                    py = origin_y - (sy - min_y) * scale  # flip: canvas y grows down
                    if first:
                        path.moveTo(px, py)
                        first = False
                    else:
                        path.lineTo(px, py)
                c.drawPath(path, stroke=1, fill=0)
            c.setLineWidth(1)
        else:
            c.setLineWidth(0.5)
            c.setStrokeColorRGB(0.7, 0.7, 0.7)
            c.line(LEFT + 20, sig_box_bottom + 22, RIGHT - 20, sig_box_bottom + 22)
            c.setStrokeColorRGB(0, 0, 0)
            c.setFont("Helvetica-Oblique", 8)
            c.setFillColorRGB(0.6, 0.6, 0.6)
            c.drawCentredString((LEFT + RIGHT) / 2, sig_box_bottom + 10, "(not signed)")
            c.setFillColorRGB(0, 0, 0)

        state['y'] = sig_box_bottom - 16

        # Printed name / date under the box
        c.setFont("Helvetica", 9.5)
        c.drawString(LEFT + 2, state['y'],
                     f"Printed Name: {self.customer_sig_name_var.get()}")
        c.drawString(LEFT + 300, state['y'],
                     f"Date: {self.customer_sig_date_var.get()}")
        state['y'] -= 22

        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.line(LEFT, state['y'] + 8, RIGHT, state['y'] + 8)
        c.setStrokeColorRGB(0, 0, 0)

        c.setFont("Helvetica", 9.5)
        c.drawString(LEFT + 2, state['y'] - 4,
                     f"Service Engineer: {self.engineer_var.get()}")
        c.drawString(LEFT + 300, state['y'] - 4,
                     f"Date: {self.sig_date_var.get()}")

        draw_footer()
        c.save()
        return state['page']

    def export_to_quickbooks_iif(self):
        """Export service order as QuickBooks IIF file for Estimate import - NO PRICING (QB uses its own)"""
        if not self.linked_estimate_number:
            messagebox.showwarning("Warning", "Please select a machine first.")
            return
        
        # Check if there's any data to export
        has_labor = False
        has_parts = False
        
        try:
            labor_qty = float(self.labor_qty_var.get()) if self.labor_qty_var.get() else 0
            if labor_qty > 0:
                has_labor = True
        except:
            pass
        
        for entry in self.parts_entries:
            if entry['part_num'].get() and entry['qty'].get():
                has_parts = True
                break
        
        if not has_labor and not has_parts:
            messagebox.showwarning("Warning", "No labor or parts to export.\nPlease add items before exporting.")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export to QuickBooks IIF",
            defaultextension=".iif",
            filetypes=[("IIF files", "*.iif"), ("All files", "*.*")],
            initialfile=f"QB_Estimate_{self.linked_estimate_number}.iif"
        )
        
        if not filename:
            return
        
        try:
            # Build the IIF content
            # Note: We're NOT including prices - QuickBooks will use its own item prices
            lines = []
            
            # === ESTIMATE HEADER DEFINITION ===
            # Simplified header - no amount/price fields since QB will calculate
            lines.append("!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tDOCNUM\tMEMO\tTOPRINT")
            lines.append("!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tMEMO\tQNTY\tINVITEM")
            lines.append("!ENDTRNS")
            
            # Format date for QuickBooks (MM/DD/YYYY)
            try:
                date_obj = datetime.strptime(self.date_var.get(), '%m-%d-%Y')
                qb_date = date_obj.strftime('%m/%d/%Y')
            except:
                qb_date = datetime.now().strftime('%m/%d/%Y')
            
            # Build line items (part numbers and quantities only)
            line_items = []
            
            # Add parts - NO PRICING
            for entry in self.parts_entries:
                part_num = entry['part_num'].get().strip()
                desc = entry['desc'].get().strip()
                qty_str = entry['qty'].get().strip()
                
                if part_num and qty_str:
                    try:
                        qty = float(qty_str)
                        
                        # Memo includes machine info and description for reference
                        memo = f"{self.mfd1_var.get()} {self.model1_var.get()} S/N:{self.sn1_var.get()}"
                        if desc:
                            memo = f"{desc} - {memo}"
                        
                        line_items.append({
                            'item': part_num,
                            'qty': qty,
                            'memo': memo[:255]  # QuickBooks memo limit
                        })
                    except ValueError:
                        continue
            
            # Add labor (using WLDSVC item) - NO PRICING
            try:
                labor_qty = float(self.labor_qty_var.get()) if self.labor_qty_var.get() else 0
                if labor_qty > 0:
                    labor_memo = f"Labor: {self.mfd1_var.get()} {self.model1_var.get()} S/N:{self.sn1_var.get()}"
                    work_desc = self.work_text.get('1.0', tk.END).strip()
                    if work_desc:
                        # Add first part of work description to memo
                        labor_memo = f"{work_desc[:100]}... - {labor_memo}" if len(work_desc) > 100 else f"{work_desc} - {labor_memo}"
                    
                    line_items.append({
                        'item': 'WLDSVC',
                        'qty': labor_qty,
                        'memo': labor_memo[:255]
                    })
            except ValueError:
                pass
            
            if not line_items:
                messagebox.showwarning("Warning", "No valid line items to export.")
                return
            
            # === BUILD THE ESTIMATE TRANSACTION ===
            # Main transaction line (TRNS) - no amount, QB will calculate
            estimate_memo = f"Service Order {self.linked_estimate_number} - {self.mfd1_var.get()} {self.model1_var.get()}"
            
            # NAME is left blank - user will select customer in QuickBooks
            trns_line = f"TRNS\tESTIMATE\t{qb_date}\tAccounts Receivable\t\t{self.linked_estimate_number}\t{estimate_memo[:255]}\tN"
            lines.append(trns_line)
            
            # Split lines (SPL) for each item - quantities only, no prices
            for item in line_items:
                spl_line = f"SPL\tESTIMATE\t{qb_date}\tSales\t\t{item['memo']}\t{item['qty']:.2f}\t{item['item']}"
                lines.append(spl_line)
            
            # End transaction marker
            lines.append("ENDTRNS")
            
            # Write the IIF file (must be tab-delimited)
            with open(filename, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
            
            # Count items for summary
            parts_count = sum(1 for item in line_items if item['item'] != 'WLDSVC')
            
            # Show success with instructions
            msg = f"QuickBooks IIF file created:\n{filename}\n\n"
            msg += "ITEMS EXPORTED (quantities only - QB will use its own pricing):\n"
            msg += f"• {parts_count} part(s)\n"
            if has_labor:
                msg += f"• Labor: {self.labor_qty_var.get()} hrs (WLDSVC)\n"
            msg += "\nTO IMPORT INTO QUICKBOOKS:\n"
            msg += "1. Open QuickBooks Desktop\n"
            msg += "2. Go to File → Utilities → Import → IIF Files\n"
            msg += "3. Select the file you just created\n"
            msg += "4. Click Open to import\n\n"
            msg += "AFTER IMPORT:\n"
            msg += "• Find the Estimate (it will have no customer assigned)\n"
            msg += "• Open it and assign the correct customer\n"
            msg += "• QuickBooks will automatically apply your item prices\n"
            msg += "• Review and save\n\n"
            msg += f"Estimate #: {self.linked_estimate_number}"
            
            messagebox.showinfo("QuickBooks Export Complete", msg)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export IIF file: {e}")
    
    def clear_form(self):
        """Clear all form fields"""
        if messagebox.askyesno("Clear Form", "Are you sure you want to clear all fields?"):
            self.linked_estimate_number = None
            self.linked_label.config(text="No machine selected")
            
            # Clear all StringVars
            for var in [self.estimate_date_var, self.repair_date_var, self.po_number_var,
                       self.order_number_var, self.bill_to_var, self.ship_to_var,
                       self.address_var, self.contact_var, self.tag_var, self.phone_var,
                       self.mfd1_var, self.model1_var, self.sn1_var,
                       self.mfd2_var, self.model2_var, self.sn2_var, self.power_var,
                       self.engineer_var, self.customer_sig_name_var,
                   self.customer_email_var]:
                var.set('')
            
            self.date_var.set(datetime.now().strftime('%m-%d-%Y'))
            self.sig_date_var.set(datetime.now().strftime('%m-%d-%Y'))
            self.customer_sig_date_var.set(datetime.now().strftime('%m-%d-%Y'))
            self.repair_date_var.set(datetime.now().strftime('%m-%d-%Y'))
            
            # Clear signature
            self.signature_canvas.delete('signature')
            self._signature_strokes = []
            self._signature_canvas_size = None
            self.plug_var.set('NO')
            self.prepower_var.set('NO')
            self.additional_var.set('NO')
            self.labor_qty_var.set('0')
            self.select_default_labor_rate()

            # Clear text fields
            self.reason_text.delete('1.0', tk.END)
            self.work_text.delete('1.0', tk.END)
            
            # Clear parts
            for entry in self.parts_entries:
                entry['part_num'].set('')
                entry['desc'].set('')
                entry['qty'].set('')
                entry['each'].set('')
            
            self.calculate_totals()



# Main application entry point
def main():
    """Main function to run the application"""
    try:
        root = tk.Tk()
        app = MachineRepairApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Error starting application: {e}")
        messagebox.showerror("Startup Error", f"Failed to start application:\n{str(e)}")

if __name__ == "__main__":
    main()
