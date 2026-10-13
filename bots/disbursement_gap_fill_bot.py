import argparse
import logging
import os
import socket
import time

from openpyxl import load_workbook
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

socket.setdefaulttimeout(30)

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S')
logger = logging.getLogger(__name__)

SCREENSHOT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_screenshots")

DISBURSEMENT_SHEET_START_ROW = 5
COL_CLIENT_CODE = 1
COL_CONTRACT_NO = 2
COL_CUSTOMER_NAME = 3
COL_GENDER = 6
COL_SECTOR = 7
COL_SUB_SECTOR = 8

LTV_WORKING_SHEET_NAME = "LTV Working - Disbursement"
LTV_WORKING_START_ROW = 5
LTV_COL_CONTRACT_NO = 1
LTV_COL_CUSTOMER_NAME = 2

GENDER_TITLE_CHECKS = [
    (3, "MR.", "Male"),
    (3, "REV", "Male"),
    (4, "MISS", "Female"),
    (3, "MS.", "Female"),
    (4, "MRS.", "Female"),
]

CLIENT_TYPE_FIELD_IDS = {
    "Individual": {
        "radio": "redIndividualClientType_CUSEClientMain",
        "name": "strClientName_CUSEClientMain",
        "gender": "strGenderINVClnt_CUSEClientMain",
        # Confirmed from live page HTML - note the ERP's own typo ("Mairtal", not "Marital").
        "marital_status": "strMairtalStatusINVClnt_CUSEClientMain",
        "description": "strSectorDescINVClnt_CUSEClientMain",
    },
    "Company": {
        "radio": "redCompanyClientType_CUSEClientMain",  # TODO: confirm
        "name": "strClientName_CUSEClientMain",
        "gender": None,  # Company records have no Gender field
        "marital_status": None,  # Company records have no Marital Status field
        "description": "strSectorDescCOMClnt_CUSEClientMain",  # TODO: confirm
    },
}

# Salutation tokens stripped off the front of an existing Customer Name
# before rebuilding it with the correct title - covers both a name with no
# title at all and one with a wrong/junk title (e.g. a literal "Other" typed
# where "Mrs."/"Miss" belonged).
STRIPPABLE_TITLE_TOKENS = {
    "MR", "MR.", "MRS", "MRS.", "MISS", "MS", "MS.", "REV", "REV.",
    "DR", "DR.", "OTHER", "OTHER.",
}


def _strip_existing_title(customer_name):
    name = str(customer_name or "").strip()
    if not name:
        return ""
    parts = name.split(None, 1)
    if parts[0].upper() in STRIPPABLE_TITLE_TOKENS and len(parts) > 1:
        return parts[1].strip()
    return name


def title_from_gender_marital(gender, marital_status):
    """Male -> Mr. always. Female -> Miss if the ERP's Marital Status field
    reads Single/Unmarried, otherwise Mrs. (Married/Widowed/Divorced/unknown
    all fall back to Mrs., per convention agreed with the user)."""
    if gender == "Male":
        return "Mr."
    if gender == "Female":
        status = (marital_status or "").strip().upper()
        if status in ("SINGLE", "UNMARRIED"):
            return "Miss"
        return "Mrs."
    return None


def build_corrected_title_name(existing_name, gender, marital_status):
    """Rebuild existing_name with a title derived from Gender/Marital Status
    (not from whatever prefix, if any, it currently has). Returns None if
    there's no bare name to attach a title to, or no gender to base a title
    on (e.g. ERP lookup failed)."""
    bare_name = _strip_existing_title(existing_name)
    title = title_from_gender_marital(gender, marital_status)
    if not bare_name or not title:
        return None
    return f"{title} {bare_name}"


def client_type_from_code(client_code):
    """'1...' -> Individual, '2...' -> Company, anything else -> None (unknown/skip)."""
    if not client_code:
        return None
    if client_code.startswith("1"):
        return "Individual"
    if client_code.startswith("2"):
        return "Company"
    return None


def is_blank(value):
    return value is None or str(value).strip() == ""


def normalize_client_code(value):
    """IA Working reads Client Code as a float (e.g. 11158273.0); ERP wants a plain code string."""
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def gender_from_name(customer_name):
    """Mirror the Gender formula's own title-matching logic, so we can tell
    which rows it would evaluate to FALSE for, without needing Excel to have
    calculated anything."""
    if is_blank(customer_name):
        return None
    name = str(customer_name).strip().upper()
    for length, prefix, gender in GENDER_TITLE_CHECKS:
        if name[:length] == prefix:
            return gender
    return None


def find_gap_rows(ws):
    """Rows missing Customer Name, Sector, or Sub-Sector (plus Gender, for
    Individual codes only - Company records have no Gender field), or whose
    Customer Name is present but has no recognized title prefix (missing
    entirely, or a wrong/junk one like "Other") - those get "title" flagged
    so the name's prefix gets rebuilt from ERP Gender/Marital Status. Client
    codes that aren't recognisably Individual ('1...') or Company ('2...')
    are skipped, since we wouldn't know which ERP client type to search."""
    gap_rows = []
    for row in range(DISBURSEMENT_SHEET_START_ROW, ws.max_row + 1):
        raw_code = ws.cell(row=row, column=COL_CLIENT_CODE).value
        client_code = normalize_client_code(raw_code)
        client_type = client_type_from_code(client_code)
        if not client_type:
            continue

        customer_name = ws.cell(row=row, column=COL_CUSTOMER_NAME).value
        sector = ws.cell(row=row, column=COL_SECTOR).value
        sub_sector = ws.cell(row=row, column=COL_SUB_SECTOR).value
        name_missing_title = client_type == "Individual" and gender_from_name(customer_name) is None

        needs = {
            "name": is_blank(customer_name),
            "gender": name_missing_title,
            "title": name_missing_title and not is_blank(customer_name),
            "sector": is_blank(sector),
            "sub_sector": is_blank(sub_sector),
        }
        if any(needs.values()):
            gap_rows.append({
                "row": row,
                "client_code": client_code,
                "client_type": client_type,
                "needs": needs,
            })

    return gap_rows


def backfill_ltv_working_customer_names(wb):
    """Fill blank Customer Name cells in the 'LTV Working - Disbursement'
    sheet using the Disbursement sheet's own (already ERP-backfilled)
    Contract No -> Customer Name data - no new ERP lookups. Every Vehicles
    and Machinery-collateral contract in LTV Working also appears as a row
    in Disbursement (same 3 monthly IA Working files feed both sheets - see
    ltv_working_disbursement.py/disbursement_processor.py), so whatever name
    ERP resolved there can just be copied across by Contract No."""
    if LTV_WORKING_SHEET_NAME not in wb.sheetnames:
        return 0

    ws_disb = wb['Disbursement']
    contract_to_name = {}
    for row in range(DISBURSEMENT_SHEET_START_ROW, ws_disb.max_row + 1):
        contract_no = ws_disb.cell(row=row, column=COL_CONTRACT_NO).value
        name = ws_disb.cell(row=row, column=COL_CUSTOMER_NAME).value
        if contract_no is None or is_blank(name):
            continue
        contract_to_name[str(contract_no).strip()] = name

    ws_ltv = wb[LTV_WORKING_SHEET_NAME]
    updated = 0
    for row in range(LTV_WORKING_START_ROW, ws_ltv.max_row + 1):
        contract_no = ws_ltv.cell(row=row, column=LTV_COL_CONTRACT_NO).value
        if contract_no is None:
            continue
        name_cell = ws_ltv.cell(row=row, column=LTV_COL_CUSTOMER_NAME)
        if not is_blank(name_cell.value):
            continue
        name = contract_to_name.get(str(contract_no).strip())
        if name:
            name_cell.value = name
            updated += 1

    return updated


def _dismiss_stray_alert(driver, client_code):
    """A JS alert() (e.g. the ERP's "Client Blocked" business-rule popup) can
    arrive slightly late - after the record that triggered it has already
    finished fetch_client_details()'s own post-Enter response-wait loop - and
    sit there undismissed when the next record's lookup starts. A native
    alert steals input focus for the whole page, so every following Selenium
    action (radio click, typing into the code field) fails instantly with
    "element not interactable" - which looks like a broken field but is
    really just an unclosed alert left over from a previous search. Call this
    first thing in fetch_client_details(), before touching the radio or code
    field, to catch that leftover case (the post-Enter loop still separately
    catches an alert that appears for the record currently being searched)."""
    try:
        alert = driver.switch_to.alert
    except Exception:
        return
    print(f"  (stray alert left over from a previous search: {alert.text!r} - dismissing before continuing)")
    logger.warning(f"  {client_code}: stray alert left over from a previous search dismissed: {alert.text!r}")
    try:
        alert.accept()
    except Exception:
        pass
    driver.switch_to.default_content()


_FIND_AND_CLOSE_DIALOG_JS = """
const sel = '.modal.show, .modal.in, .modal[style*="display: block"], .swal2-popup, '
          + '.sweet-alert.visible, .bootbox, .ui-dialog, [role="dialog"], [role="alertdialog"]';
const visible = el => el.getClientRects().length > 0 && getComputedStyle(el).visibility !== 'hidden';
const dialogs = Array.from(document.querySelectorAll(sel)).filter(visible);
if (!dialogs.length) return null;
const dlg = dialogs[dialogs.length - 1];
const text = (dlg.innerText || '').trim().replace(/\\s+/g, ' ').slice(0, 300);
const buttons = Array.from(dlg.querySelectorAll('button, input[type=button], input[type=submit], a.btn, .close'))
    .filter(visible);
const label = b => (b.innerText || b.value || b.getAttribute('aria-label') || '').trim();
const preferred = buttons.find(b => /^(ok|close|yes|okay|x|×)$/i.test(label(b))) || buttons[0];
if (preferred) { preferred.click(); }
return {text: text, clicked: preferred ? label(preferred) : null, id: dlg.id, cls: dlg.className};
"""


def _dismiss_error_dialog(driver, client_code):
    """The ERP sometimes pops up an in-page HTML dialog titled "ERROR" (a DOM
    modal, not a native alert(), so _dismiss_stray_alert can't see it). While
    it's open it sits on top of the page and every Selenium click/keystroke
    underneath fails with "element not interactable". Finds the topmost
    visible dialog, logs its text/id/class (so a precise selector can be
    hardcoded later) plus a screenshot, and clicks its OK/Close button.
    Returns True if a dialog was found."""
    try:
        driver.switch_to.default_content()
        info = driver.execute_script(_FIND_AND_CLOSE_DIALOG_JS)
    except Exception as e:
        logger.debug(f"  {client_code}: dialog check failed ({type(e).__name__}: {e})")
        return False
    if not info:
        return False

    print(f"  (in-page dialog found: {info['text']!r} - clicked {info['clicked']!r})")
    # The click above already ran, so this screenshot shows the page after;
    # the logged text/id/class are the record of what the dialog was.
    try:
        os.makedirs(SCREENSHOT_DIR, exist_ok=True)
        driver.save_screenshot(os.path.join(SCREENSHOT_DIR, f"{client_code}_dialog_{int(time.time())}.png"))
    except Exception:
        pass
    logger.warning(f"  {client_code}: in-page dialog dismissed - text={info['text']!r} "
                   f"clicked={info['clicked']!r} id={info['id']!r} class={info['cls']!r}")
    time.sleep(1)
    return True


def _fetch_with_dialog_retry(driver, wait, code, **kwargs):
    """fetch_client_details(), but if it fails because an ERP error dialog is
    covering the page, close the dialog and try that same code once more."""
    try:
        return fetch_client_details(driver, wait, code, **kwargs)
    except WebDriverException as e:
        if _is_connection_error(e) or not _dismiss_error_dialog(driver, code):
            raise
        logger.info(f"  {code}: retrying lookup after closing the ERP error dialog")
        return fetch_client_details(driver, wait, code, **kwargs)


def _is_connection_error(exc):
    """Distinguish a dead browser/ERP session (real problem) from a plain
    'element not there yet' NoSuchElementException (expected while polling)."""
    msg = str(exc).lower()
    return any(s in msg for s in (
        "invalid session id", "chrome not reachable", "disconnected",
        "no such window", "target window already closed", "session deleted",
        "not connected to devtools", "connection refused", "connection reset",
    ))


def _find_description_field(driver, client_type, client_code):
    """CLIENT_TYPE_FIELD_IDS['Company']['description'] (strSectorDescCOMClnt_CUSEClientMain)
    was an unconfirmed guess and doesn't actually exist on the ERP page - every
    Company lookup was finding the client fine (Name/type both resolve) but then
    throwing 'no such element' on this one field, which discarded the whole
    result. Try the configured id first; if it's missing, scan the DOM for the
    real Sector-description input and log what's found so the id can be fixed
    properly in CLIENT_TYPE_FIELD_IDS."""
    configured_id = CLIENT_TYPE_FIELD_IDS[client_type]["description"]
    try:
        return driver.find_element(By.ID, configured_id)
    except NoSuchElementException:
        pass

    logger.warning(f"  {client_code}: configured description field id {configured_id!r} not found on page "
                    f"for client type {client_type!r} - scanning DOM for the real Sector-description field...")
    candidates = driver.find_elements(By.CSS_SELECTOR, "input[id*='SectorDesc']")
    individual_id = CLIENT_TYPE_FIELD_IDS["Individual"]["description"]
    others = [el for el in candidates if el.get_attribute("id") != individual_id]
    for el in candidates:
        logger.info(f"  {client_code}: candidate Sector-description field on page: "
                    f"id={el.get_attribute('id')!r} value={el.get_attribute('value')!r}")

    if len(others) == 1:
        found_id = others[0].get_attribute("id")
        logger.warning(f"  {client_code}: using discovered field id {found_id!r} for this lookup - "
                       f"update CLIENT_TYPE_FIELD_IDS['{client_type}']['description'] to this value "
                       f"to stop relying on DOM discovery")
        return others[0]

    raise NoSuchElementException(
        f"Could not find the {client_type} Sector-description field: configured id {configured_id!r} "
        f"doesn't exist, and DOM scan found {len(others)} other 'SectorDesc' candidate(s) instead of exactly 1"
    )


def _find_marital_status_field(driver, client_type, client_code):
    """Marital Status is a <select> (not <input>, unlike the other fields
    here), confirmed from live page HTML - CLIENT_TYPE_FIELD_IDS' configured
    id should match directly. Kept a DOM-scan fallback anyway, same
    guess-then-scan pattern as _find_description_field, in case the ERP page
    changes; it covers both the ERP's actual "Mairtal" typo and the correctly
    spelled "Marital" in case that ever gets fixed on their end. Returns None
    (not raise) on failure, since a missing Marital Status shouldn't block
    Gender/Name/Sector, which is why this is the only field-finder here where
    an outer 'if it's missing, return None' caller is expected."""
    configured_id = CLIENT_TYPE_FIELD_IDS[client_type]["marital_status"]
    if not configured_id:
        return None
    try:
        return driver.find_element(By.ID, configured_id)
    except NoSuchElementException:
        pass

    logger.warning(f"  {client_code}: configured marital-status field id {configured_id!r} not found on page "
                    f"for client type {client_type!r} - scanning DOM for the real Marital Status field...")
    candidates = driver.find_elements(By.CSS_SELECTOR, "select[id*='Mairtal'], select[id*='Marital']")
    for el in candidates:
        logger.info(f"  {client_code}: candidate Marital Status field on page: "
                    f"id={el.get_attribute('id')!r} value={el.get_attribute('value')!r}")

    if len(candidates) == 1:
        found_id = candidates[0].get_attribute("id")
        logger.warning(f"  {client_code}: using discovered field id {found_id!r} for this lookup - "
                       f"update CLIENT_TYPE_FIELD_IDS['{client_type}']['marital_status'] to this value "
                       f"to stop relying on DOM discovery")
        return candidates[0]

    logger.error(f"  {client_code}: could not find the Marital Status field - configured id {configured_id!r} "
                 f"doesn't exist, and DOM scan found {len(candidates)} 'Marital' candidate(s) instead of "
                 f"exactly 1; title correction for this record will fall back to Gender alone")
    return None


def fetch_client_details(driver, wait, client_code, client_type=None, need_marital_status=True):
    """Search one client code on the ERP Customer Inquiry screen and read
    back Name/Gender/Description. client_type is 'Individual' or 'Company'
    (inferred from the code's leading digit via client_type_from_code() if
    not given). Returns None if the code doesn't resolve to anything (fields
    stay empty).

    need_marital_status: pass False to skip reading Marital Status even for
    an Individual - it's only needed to build a corrected title for a row
    whose Customer Name is already filled in (a blank name is filled from
    ERP's own Name field as-is, no title math needed), and it's never needed
    for Male clients (title is Mr. regardless of marital status) - both
    checks save a field read that isn't otherwise used, which matters across
    a large batch. Defaults to True so the standalone test_client_codes()
    smoke test still reads it, for debugging."""

    start_ts = time.time()
    client_type = client_type or client_type_from_code(client_code)
    if client_type not in CLIENT_TYPE_FIELD_IDS:
        raise ValueError(f"Can't tell client type for code {client_code!r} (expected to start with '1' or '2')")
    field_ids = CLIENT_TYPE_FIELD_IDS[client_type]

    _dismiss_stray_alert(driver, client_code)
    _dismiss_error_dialog(driver, client_code)

    try:
        WebDriverWait(driver, 15).until(EC.invisibility_of_element_located((By.ID, "02626jqueryLoaderDiv")))
    except TimeoutException:
        logger.warning(f"  {client_code}: loading spinner still visible after 15s - proceeding anyway")
        print("  (loading spinner still visible after 15s - proceeding anyway)")

    print(f"  -> Selecting '{client_type}' and searching client code {client_code}...")
    logger.info(f"  {client_code}: elapsed {time.time() - start_ts:.1f}s so far - selecting client type and typing code")

    try:
        type_radio = wait.until(EC.presence_of_element_located((By.ID, field_ids["radio"])))
    except TimeoutException as e:
        logger.error(f"  {client_code}: TIMEOUT waiting for client-type radio button ({field_ids['radio']}) - "
                     f"page may not have loaded, or session/connection was lost. {e}")
        raise
    except WebDriverException as e:
        if _is_connection_error(e):
            logger.error(f"  {client_code}: CONNECTION/SESSION ERROR while locating client-type radio - {e}")
        raise

    if not type_radio.is_selected():
        driver.execute_script("arguments[0].click();", type_radio)
        time.sleep(1)  # give any resulting form reset/reload a moment to settle
    else:
        print(f"  ('{client_type}' already selected - not clicking)")

    try:
        code_input = wait.until(EC.element_to_be_clickable((By.ID, "strClientCode_CUSEClientMain")))
    except TimeoutException as e:
        logger.error(f"  {client_code}: TIMEOUT waiting for client-code input field to become clickable - "
                     f"page may not have loaded, or session/connection was lost. {e}")
        raise
    except WebDriverException as e:
        if _is_connection_error(e):
            logger.error(f"  {client_code}: CONNECTION/SESSION ERROR while locating client-code input - {e}")
        raise

    try:
        code_input.click()
    except Exception:
        driver.execute_script("arguments[0].click();", code_input)

    code_input.send_keys(Keys.END)
    for _ in range(15):
        code_input.send_keys(Keys.BACK_SPACE)
    for ch in str(client_code):
        code_input.send_keys(ch)
        time.sleep(0.08)
    time.sleep(0.3)
    print(f"  (typed code, about to check field value...)")

    typed_value = code_input.get_attribute("value")
    if typed_value.strip() != str(client_code).strip():
        print(f"  (field reads {typed_value!r} after real-keystroke typing - mask may be rejecting this input; leaving as-is)")

    print(f"  (field confirmed, pressing Enter...)")
    code_input.send_keys(Keys.ENTER)
    print(f"  (Enter sent, waiting for a response...)")

    deadline = time.time() + 20
    found = False
    connection_error = None
    while time.time() < deadline:
        try:
            alert = driver.switch_to.alert
            print(f"  (alert appeared: {alert.text!r} - dismissing)")
            logger.info(f"  {client_code}: unexpected alert dismissed: {alert.text!r}")
            alert.accept()
            driver.switch_to.default_content()
        except Exception:
            pass

        try:
            if driver.find_element(By.ID, field_ids["name"]).get_attribute("value").strip():
                found = True
                break
        except WebDriverException as e:
            if _is_connection_error(e):
                # Browser/session died mid-poll - no point continuing to poll for 20s.
                connection_error = e
                break
        except Exception:
            pass
        time.sleep(0.5)

    if connection_error is not None:
        logger.error(f"  {client_code}: CONNECTION/SESSION ERROR while waiting for ERP response "
                     f"(after {time.time() - start_ts:.1f}s) - {connection_error}")
        raise connection_error

    if not found:
        elapsed = time.time() - start_ts
        print(f"  -> {client_code}: NOT FOUND (or no response) in ERP")
        logger.info(f"  {client_code}: not found in ERP (waited {elapsed:.1f}s total - if this is a TIMEOUT "
                     f"rather than a real 'no such client', check whether the ERP session has expired: "
                     f"long unattended runs can time out mid-batch and every code afterwards will show as "
                     f"NOT FOUND even though it exists)")
        try:
            os.makedirs(SCREENSHOT_DIR, exist_ok=True)
            shot_path = os.path.join(SCREENSHOT_DIR, f"{client_code}_{int(time.time())}.png")
            driver.save_screenshot(shot_path)
            print(f"  (saved screenshot for inspection: {shot_path})")
            logger.info(f"  {client_code}: diagnostic screenshot saved to {shot_path}")
        except Exception as shot_err:
            print(f"  (could not save diagnostic screenshot: {shot_err})")
            logger.warning(f"  {client_code}: could not save diagnostic screenshot ({type(shot_err).__name__}: {shot_err})")
        return None

    logger.info(f"  {client_code}: found after {time.time() - start_ts:.1f}s - now sleeping 100s before reading fields")
    time.sleep(10)  # give the ERP a moment to finish any remaining AJAX updates

    name = driver.find_element(By.ID, field_ids["name"]).get_attribute("value").strip()

    gender = None
    if field_ids["gender"]:
        gender_raw = driver.find_element(By.ID, field_ids["gender"]).get_attribute("value").strip().upper()
        gender = {"MALE": "Male", "FEMALE": "Female"}.get(gender_raw)

    marital_status = None
    if field_ids["marital_status"] and need_marital_status and gender == "Female":
        marital_field = _find_marital_status_field(driver, client_type, client_code)
        if marital_field is not None:
            marital_status = marital_field.get_attribute("value").strip() or None

    description = _find_description_field(driver, client_type, client_code).get_attribute("value").strip()

    sector, sub_sector = "", ""
    if description:
        parts = description.split("-", 1)
        sector = parts[0].strip()
        sub_sector = parts[1].strip() if len(parts) > 1 else ""

    result = {"name": name, "gender": gender, "marital_status": marital_status, "sector": sector, "sub_sector": sub_sector}
    print(f"  -> {client_code} ({client_type}): FOUND - Name={name!r} Gender={gender!r} "
          f"MaritalStatus={marital_status!r} Sector={sector!r} Sub-Sector={sub_sector!r}")
    logger.info(f"  {client_code}: lookup complete in {time.time() - start_ts:.1f}s")
    return result


def navigate_to_customer_inquiry(driver, wait):
    """Login already done - refresh (needed for the nav menu to load, same
    as the already-working client_title_extraction_bot.py flow), open Client
    Management, the Customer Creation tile, then Inquiry."""

    print("Refreshing page so the navigation menu loads...")
    driver.refresh()
    time.sleep(5)

    print("Opening Client Management...")
    logger.info("Opening Client Management...")
    client_mgmt_link = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//a[@title='Customer Evaluation']//span[contains(text(),'Client Management')]"
    )))
    driver.execute_script("arguments[0].scrollIntoView(true);", client_mgmt_link)
    client_mgmt_link.click()

    wait.until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(3)

    print("Opening Customer Creation tile...")
    logger.info("Opening Customer Creation tile...")
    tile_xpaths = [
        "//div[contains(@class,'metro-tile-body')]//div[contains(text(),'Customer Creation')]",
        "//div[@class='metro-content-hide' and contains(text(),'Customer Creation')]",
    ]
    customer_tile = None
    for path in tile_xpaths:
        try:
            customer_tile = wait.until(EC.element_to_be_clickable((By.XPATH, path)))
            break
        except Exception:
            continue
    if customer_tile is None:
        raise RuntimeError("Customer Creation tile not found")

    driver.execute_script("arguments[0].scrollIntoView(true);", customer_tile)
    from selenium.webdriver.common.action_chains import ActionChains
    ActionChains(driver).move_to_element(customer_tile).double_click().perform()
    time.sleep(3)

    print("Clicking Inquiry...")
    logger.info("Clicking Inquiry...")
    inquiry_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='btn_02626_Inquiry']")))
    inquiry_button.click()
    time.sleep(3)

    _switch_to_frame_containing(driver, By.ID, "redIndividualClientType_CUSEClientMain")
    print("On Customer Inquiry screen - ready to search client codes.\n")


def _switch_to_frame_containing(driver, by, value, timeout=10):
    """This ERP's tile-launched screens commonly render inside an <iframe>,
    which Selenium can't see into unless explicitly switched. Try the
    default (top-level) content first, then each iframe on the page, until
    the target element turns up - and print what was found either way, so
    a failure here gives real evidence instead of a bare timeout."""
    driver.switch_to.default_content()
    try:
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        print(f"  (found {value!r} in the top-level page - no iframe switch needed)")
        return
    except TimeoutException:
        pass

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    print(f"  {value!r} not in top-level page; checking {len(iframes)} iframe(s)...")
    for i, iframe in enumerate(iframes):
        frame_id = iframe.get_attribute("id") or "(no id)"
        frame_src = iframe.get_attribute("src") or "(no src)"
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(iframe)
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
            print(f"  -> found in iframe #{i} (id={frame_id}, src={frame_src})")
            return
        except TimeoutException:
            print(f"  -> not in iframe #{i} (id={frame_id}, src={frame_src})")
            continue

    driver.switch_to.default_content()
    raise TimeoutException(f"{value!r} not found in the top-level page or any of {len(iframes)} iframe(s)")


def _load_credentials():
    """Matches the credential scheme actually configured on this machine:
    FERNET_KEY as an OS environment variable, ENC_USERNAME/ENC_PASSWORD in
    the project root .env - same scheme as bots/client_title_extraction_bot.py.
    NOT report_automations/erp_login.py's scheme (ERP_SECRET_KEY/
    ERP_USERNAME_ENC/ERP_PASSWORD_ENC in a report_automations/.env), which
    isn't set up in this project."""
    from dotenv import load_dotenv
    from cryptography.fernet import Fernet

    load_dotenv()  # searches upward from CWD, so it finds the project-root .env regardless of where this script is run from

    enc_username = os.getenv("ENC_USERNAME")
    enc_password = os.getenv("ENC_PASSWORD")
    if not enc_username or not enc_password:
        raise RuntimeError("ENC_USERNAME/ENC_PASSWORD not found - check the project root .env file")

    key_str = os.environ.get("FERNET_KEY", "").strip()
    if not key_str:
        raise RuntimeError("FERNET_KEY environment variable is not set")

    fernet = Fernet(key_str.encode())
    username = fernet.decrypt(enc_username.encode()).decode()
    password = fernet.decrypt(enc_password.encode()).decode()
    return username, password


def _erp_login():
    """Self-contained login (mirrors bots/client_title_extraction_bot.py's
    proven sequence) rather than report_automations/erp_login.py, since that
    module expects a credential scheme this project doesn't actually have
    configured."""
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager

    username, password = _load_credentials()

    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_experimental_option("prefs", {
        "profile.default_content_setting_values.geolocation": 1
    })

    print("Opening Chrome...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 25)

    driver.execute_cdp_cmd("Browser.grantPermissions", {
        "origin": "https://erp.assetline.lk",
        "permissions": ["geolocation"],
    })
    driver.execute_cdp_cmd("Page.setGeolocationOverride", {
        "latitude": 6.9271,
        "longitude": 79.8612,
        "accuracy": 100,
    })

    print("Opening ERP login page...")
    driver.get("https://erp.assetline.lk/")
    time.sleep(3)

    wait.until(EC.presence_of_element_located((By.ID, "strUserName"))).send_keys(username)
    wait.until(EC.presence_of_element_located((By.ID, "strPassword"))).send_keys(password)
    wait.until(EC.presence_of_element_located((By.ID, "idSignIn"))).click()
    print("Logged into ERP system")
    time.sleep(3)

    try:
        select_button = wait.until(
            EC.presence_of_element_located((By.XPATH, "//button[@id='idSignIn' and @type='submit']"))
        )
        select_button.click()
        print("Default location selected")
        time.sleep(3)
    except Exception:
        print("Default location screen not found - continuing...")

    return driver, wait


def login_and_navigate():
    print("Opening Chrome and logging into the ERP...")
    driver, wait = _erp_login()
    navigate_to_customer_inquiry(driver, wait)
    return driver, wait


def run(report_file):
    wb = load_workbook(report_file)
    if 'Disbursement' not in wb.sheetnames:
        logger.warning("'Disbursement' sheet not found - nothing to backfill")
        return

    ws = wb['Disbursement']
    gap_rows = find_gap_rows(ws)

    erp_error = None
    if not gap_rows:
        logger.info("No Individual-client gaps found in Disbursement sheet - skipping ERP lookup")
    else:
        try:
            _fill_disbursement_gaps_from_erp(ws, gap_rows)
        except Exception as e:
            # Don't let an ERP/browser failure (login, navigation, session
            # drop, etc.) skip the LTV Working backfill below - that step
            # needs no ERP access at all, so it should still run and save
            # whatever it can. Re-raised after saving so the caller still
            # sees this run as failed (rc != 0), same as before.
            erp_error = e
            logger.error(f"ERP gap-fill for Disbursement sheet failed ({type(e).__name__}: {e}); "
                         f"Disbursement sheet may be only partially updated. Continuing to backfill "
                         f"'{LTV_WORKING_SHEET_NAME}' from whatever Customer Name data is already "
                         f"present, and saving either way.")

    ltv_updated = backfill_ltv_working_customer_names(wb)
    if ltv_updated:
        logger.info(f"Backfilled {ltv_updated} Customer Name cell(s) in '{LTV_WORKING_SHEET_NAME}' sheet "
                    f"from the Disbursement sheet's own Contract No -> Customer Name data (no new ERP lookups)")

    wb.save(report_file)
    logger.info(f"Saved: {report_file}")

    if erp_error is not None:
        raise erp_error


def _fill_disbursement_gaps_from_erp(ws, gap_rows):
    distinct_codes = sorted({row["client_code"] for row in gap_rows})
    # Marital Status is only useful for codes with a "title" gap (Customer
    # Name already filled, just missing/wrong prefix) - skip the read for
    # every other code (blank-name and sector/sub_sector-only gaps).
    codes_needing_title_fix = {row["client_code"] for row in gap_rows if row["needs"]["title"]}
    logger.info(f"{len(gap_rows)} row(s) with gaps across {len(distinct_codes)} distinct client code(s)")

    driver, wait = login_and_navigate()
    results = {}
    batch_start = time.time()
    consecutive_not_found = 0
    try:
        for idx, code in enumerate(distinct_codes, start=1):
            record_start = time.time()
            elapsed_total = record_start - batch_start
            print(f"[{idx}/{len(distinct_codes)}]", end=" ")
            logger.info(f"[{idx}/{len(distinct_codes)}] Looking up {code}... "
                        f"(batch elapsed {elapsed_total/60:.1f} min)")
            try:
                results[code] = _fetch_with_dialog_retry(
                    driver, wait, code, need_marital_status=code in codes_needing_title_fix
                )
            except TimeoutException as e:
                logger.error(f"[{idx}/{len(distinct_codes)}] {code}: TIMEOUT ({e}); leaving blank")
                results[code] = None
            except WebDriverException as e:
                if _is_connection_error(e):
                    logger.error(f"[{idx}/{len(distinct_codes)}] {code}: CONNECTION/SESSION ERROR - browser or "
                                 f"ERP session appears to have dropped after {elapsed_total/60:.1f} min "
                                 f"({idx - 1} record(s) processed so far); leaving blank. {e}")
                else:
                    logger.error(f"[{idx}/{len(distinct_codes)}] {code}: WebDriver error ({e}); leaving blank")
                results[code] = None
            except socket.timeout as e:
                logger.error(f"[{idx}/{len(distinct_codes)}] {code}: NETWORK SOCKET TIMEOUT ({e}); leaving blank")
                results[code] = None
            except Exception as e:
                logger.error(f"[{idx}/{len(distinct_codes)}] {code}: lookup failed "
                             f"({type(e).__name__}: {e}); leaving blank")
                results[code] = None

            record_elapsed = time.time() - record_start
            logger.info(f"[{idx}/{len(distinct_codes)}] {code}: record took {record_elapsed:.1f}s")

            if results[code] is None:
                consecutive_not_found += 1
                if consecutive_not_found == 3:
                    logger.warning(
                        f"{consecutive_not_found} client codes in a row have come back blank/not-found "
                        f"(batch elapsed {(time.time() - batch_start)/60:.1f} min). If earlier codes were "
                        f"found fine, this streak likely means the ERP session/login expired mid-batch "
                        f"rather than these specific codes being missing - check debug_screenshots/ for "
                        f"a login or session-timeout page instead of the Customer Inquiry screen."
                    )
            else:
                consecutive_not_found = 0
    finally:
        driver.quit()

    updated = 0
    not_found = 0
    for gap in gap_rows:
        result = results.get(gap["client_code"])
        if result is None:
            not_found += 1
            continue

        row = gap["row"]
        needs = gap["needs"]
        row_updated = False
        if needs["name"] and result["name"]:
            ws.cell(row=row, column=COL_CUSTOMER_NAME).value = result["name"]
            row_updated = True
        if needs["title"]:
            # Rebuild the sheet's existing name with a title derived from
            # ERP Gender/Marital Status - not with result["name"] (ERP's own
            # Name field), since the sheet's name is otherwise correct and
            # only its title prefix is missing/wrong.
            existing_name = ws.cell(row=row, column=COL_CUSTOMER_NAME).value
            corrected_name = build_corrected_title_name(existing_name, result.get("gender"), result.get("marital_status"))
            if corrected_name:
                ws.cell(row=row, column=COL_CUSTOMER_NAME).value = corrected_name
                row_updated = True
        if needs["gender"] and result["gender"]:
            ws.cell(row=row, column=COL_GENDER).value = result["gender"]
            row_updated = True
        if needs["sector"] and result["sector"]:
            ws.cell(row=row, column=COL_SECTOR).value = result["sector"]
            row_updated = True
        if needs["sub_sector"] and result["sub_sector"]:
            ws.cell(row=row, column=COL_SUB_SECTOR).value = result["sub_sector"]
            row_updated = True
        if row_updated:
            updated += 1

    logger.info(f"Updated {updated} row(s); {not_found} client code(s) not found in ERP (left blank)")


def test_client_codes(codes):
    """Selenium-only smoke test: log in, navigate to Customer Inquiry, and
    search each given code - no xlsx file is read or written. Use this to
    watch the browser and confirm a specific client code (e.g. one still
    blank in an output file) actually resolves in the ERP or not."""

    driver, wait = login_and_navigate()
    try:
        for idx, code in enumerate(codes, start=1):
            print(f"[{idx}/{len(codes)}]", end=" ")
            try:
                fetch_client_details(driver, wait, code)
            except (TimeoutException, WebDriverException, socket.timeout) as e:
                print(f"  -> {code}: lookup failed ({e}); moving to next code")
    finally:
        input("\nDone. Press Enter to close the browser...")
        driver.quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Backfill Disbursement sheet gaps for Individual clients from ERP")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--report-file", help="Path to the Quarter Classification Loans & Advance xlsx")
    group.add_argument("--client-codes", help="Comma-separated client codes to test in the browser, no file involved")
    args = parser.parse_args()

    if args.client_codes:
        test_client_codes([c.strip() for c in args.client_codes.split(",") if c.strip()])
    else:
        run(args.report_file)