"""
ERP lookups for the MT workbook (report_automations/mt_report_consolidator.py).

fetch_genders(client_codes)
    Client Management -> Customer Creation -> Inquiry -> Client Code -> Gender.
    Reuses the Customer Inquiry code of bots/disbursement_gap_fill_bot.py.

fetch_gl_opening_balances(dates, sub_ledger_code="1846")
    Finance -> GL Account View -> Inquiry -> Sub Ledger Code + Transaction Date
    (same date in both date fields) -> Search -> Opening Balance.
    e.g. the balance as at 30-Sep-2025 is the opening balance on 01/10/2025.

Run directly to try them:
    python bots/mt_erp_lookups.py --genders 11071803 10854724
    python bots/mt_erp_lookups.py --gl-dates 01/07/2025 01/10/2025
"""
import os
import re
import sys
import time
import logging
import argparse

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

logger = logging.getLogger(__name__)

FINANCE_URL = "https://erp.assetline.lk/Application/Home/FINANCE"
GL_INQUIRY_BTN_ID = "btn_02366_Inquiry"
GL_SUB_LEDGER_ID = "strSubLedgerCode_FINGLAccountView"
GL_FROM_DATE_ID = "dtmTransactionDate_FINGLAccountView"
GL_TO_DATE_ID = "dtmTransactionToDate_FINGLAccountView"
GL_SEARCH_BTN_ID = "btn_02366_BtnSearch_FINGLAccountView"
GL_RESULT_TIMEOUT = 60  # seconds to wait for the Opening Balance after Search
# The Opening Balance box: tried by id first, then as the first input after
# the on-screen "Opening Balance" label (its id doesn't contain the words).
GL_OPENING_BALANCE_XPATHS = [
    "//input[contains(@id,'_FINGLAccountView') and "
    "contains(translate(@id,'OPENINGBALNC','openingbalnc'),'openingbalance')]",
    "//*[normalize-space(text())='Opening Balance']/following::input[1]",
]


def fetch_genders(client_codes):
    """{client code: 'Male' / 'Female' / None} for Individual client codes."""
    from bots.disbursement_gap_fill_bot import login_and_navigate, _fetch_with_dialog_retry

    results = {}
    codes = sorted({str(c).strip() for c in client_codes if c})
    if not codes:
        return results
    logger.info(f"ERP gender lookup for {len(codes)} client code(s)")
    driver, wait = login_and_navigate()
    try:
        for i, code in enumerate(codes, start=1):
            try:
                found = _fetch_with_dialog_retry(driver, wait, code, client_type="Individual",
                                                 need_marital_status=False)
                results[code] = found["gender"] if found else None
            except Exception as e:
                logger.warning(f"[{i}/{len(codes)}] {code}: ERP lookup failed ({type(e).__name__}: {e})")
                results[code] = None
            logger.info(f"[{i}/{len(codes)}] {code}: gender = {results[code]}")
    finally:
        driver.quit()
    return results


def _parse_amount(text):
    """'1,819,501,992.51' / '(1,234.00)' / '1,234.00 Dr' -> float; None if empty."""
    if text is None or not str(text).strip():
        return None
    raw = str(text).strip()
    negative = raw.startswith('(') and raw.endswith(')') or raw.startswith('-')
    number = re.sub(r'[^0-9.]', '', raw)
    if not number:
        return None
    value = float(number)
    return -value if negative else value


def _set_input(driver, element_id, value):
    el = driver.find_element(By.ID, element_id)
    driver.execute_script(
        "arguments[0].removeAttribute('readonly');"
        "arguments[0].value = arguments[1];"
        "arguments[0].dispatchEvent(new Event('input',  {bubbles:true}));"
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));"
        "arguments[0].dispatchEvent(new Event('blur',   {bubbles:true}));",
        el, value)
    return el.get_attribute("value")


def _open_gl_account_view(driver, wait, wait_for_loader):
    driver.get(FINANCE_URL)
    wait_for_loader(driver)
    time.sleep(3)
    for el in driver.find_elements(By.XPATH, "//*[normalize-space(text())='GL Account View']"):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            driver.execute_script("arguments[0].click();", el)
            wait_for_loader(driver)
            break
        except Exception:
            continue
    else:
        raise RuntimeError("'GL Account View' tile not found on the Finance page")
    time.sleep(2)
    inquiry = wait.until(EC.presence_of_element_located((By.ID, GL_INQUIRY_BTN_ID)))
    driver.execute_script("arguments[0].click();", inquiry)
    wait_for_loader(driver)
    wait.until(EC.presence_of_element_located((By.ID, GL_SUB_LEDGER_ID)))


def fetch_gl_opening_balances(dates, sub_ledger_code="1846"):
    """{'dd/mm/yyyy': opening balance} from GL Account View, one search per date
    (the date goes into both the From and To transaction date fields)."""
    from bots.erp_login import erp_login, wait_for_loader
    from bots.reschedule_contract_details_bot import _select_premises

    driver, wait = erp_login()
    balances = {}
    try:
        _select_premises(driver)
        _open_gl_account_view(driver, wait, wait_for_loader)
        for date in dates:
            sub_ledger = driver.find_element(By.ID, GL_SUB_LEDGER_ID)
            sub_ledger.clear()
            sub_ledger.send_keys(sub_ledger_code)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", sub_ledger)
            wait_for_loader(driver)
            shown_from = _set_input(driver, GL_FROM_DATE_ID, date)
            shown_to = _set_input(driver, GL_TO_DATE_ID, date)
            logger.info(f"GL {sub_ledger_code}: dates set to {shown_from} - {shown_to}")

            field = next((found[0] for xpath in GL_OPENING_BALANCE_XPATHS
                          if (found := driver.find_elements(By.XPATH, xpath))), None)
            if field is None:
                raise RuntimeError("Opening Balance field not found on GL Account View")
            # Empty the box first, so the value read below is this search's
            # result - not the previous date's, and not read before it loads.
            driver.execute_script("arguments[0].value = '';", field)

            search = wait.until(EC.element_to_be_clickable((By.ID, GL_SEARCH_BTN_ID)))
            driver.execute_script("arguments[0].click();", search)
            wait_for_loader(driver)

            raw = ''
            deadline = time.time() + GL_RESULT_TIMEOUT
            while time.time() < deadline:
                raw = (driver.execute_script("return arguments[0].value;", field) or '').strip()
                if raw:
                    break
                time.sleep(0.5)
            if not raw:
                logger.warning(f"GL {sub_ledger_code}: Opening Balance still empty {GL_RESULT_TIMEOUT}s "
                               f"after searching {date}")
            balances[date] = _parse_amount(raw)
            logger.info(f"GL {sub_ledger_code} opening balance on {date} "
                        f"(field id {field.get_attribute('id')!r}): {raw!r} -> {balances[date]}")
    except Exception:
        try:
            driver.save_screenshot(os.path.join(CURRENT_DIR, "mt_gl_balance_error.png"))
        except Exception:
            pass
        raise
    finally:
        driver.quit()
    return balances


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s',
                        datefmt='%H:%M:%S', handlers=[logging.StreamHandler(sys.stdout)], force=True)
    parser = argparse.ArgumentParser(description="Try the MT ERP lookups")
    parser.add_argument("--genders", nargs="*", help="Individual client (ICAM) codes")
    parser.add_argument("--gl-dates", nargs="*", help="dd/mm/yyyy dates for the GL 1846 opening balance")
    args = parser.parse_args()
    if args.genders:
        print(fetch_genders(args.genders))
    if args.gl_dates:
        print(fetch_gl_opening_balances(args.gl_dates))
