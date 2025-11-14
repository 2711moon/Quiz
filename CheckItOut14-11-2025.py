#report is shown at the end in the terminal
import time
import json
import openpyxl 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    ElementClickInterceptedException, NoAlertPresentException
)

# ---CONFIG 
QUIZ_URL = "https://training.kisna.com/"
DEPARTMENT_TO_SELECT = "IT"
MAX_WAIT = 15
EXCEL_FILE = "users.xlsx"  #user details: Name and respective Mobile number
PROGRESS_FILE = "quiz_progress.json"  #tracking users who have already completed the quiz
VERIFY_DONE_USERS = True  #it reverifies users marked as done, if need no such reverification, then set it to false
# --- END CONFIG


def safe_click(driver, element, retries=2):
    """Try to click an element safely with retries."""
    for attempt in range(retries):
        try:
            element.click()
            return True
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].scrollIntoView();", element)
            time.sleep(0.5)
        except Exception:
            time.sleep(0.5)
    return False


def load_users_from_excel(filename):
    """Read Name and Mobile from users.xlsx"""
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active
    users = []
    for row in sheet.iter_rows(min_row=2, values_only=True):  # skip header row
        name, mobile = row
        if name and mobile:
            users.append((str(name).strip(), str(mobile).strip()))
    return users


def load_progress():
    """Load finished users from JSON file."""
    try:
        with open(PROGRESS_FILE, "r") as f:
            return set(json.load(f))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()


def save_progress(done_users):
    """Save finished users to JSON file."""
    with open(PROGRESS_FILE, "w") as f:
        json.dump(list(done_users), f, indent=2)


def create_driver():
    """Create a fresh Chrome driver."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    return driver


def run_quiz_for_user(driver, wait, name, mobile, chosen_index, ans_word):
    print(f"\n🚀 Running quiz for {name} ({mobile})")

    try:
        #Back Office Login
        back_btn = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//a[contains(., 'BACK OFFICE LOGIN') or contains(., 'Back Office')]")))
        safe_click(driver, back_btn)
        print("Clicked BACK OFFICE LOGIN")

        time.sleep(1)

        #Department - IT Selection
        dept_select_el = wait.until(EC.presence_of_element_located((By.TAG_NAME, "select")))
        select = Select(dept_select_el)
        try:
            select.select_by_visible_text(DEPARTMENT_TO_SELECT)
        except Exception:
            select.select_by_index(0)
        print("Department selected:", DEPARTMENT_TO_SELECT)

        #Credentials
        def find_input(xpaths):
            for xp in xpaths:
                try:
                    return driver.find_element(By.XPATH, xp)
                except NoSuchElementException:
                    continue
            return None

        name_input = find_input([
            "//input[@name='name']", "//input[@id='name']",
            "//input[contains(@placeholder, 'Name')]"
        ])
        phone_input = find_input([
            "//input[@name='mobile']", "//input[@id='mobile']",
            "//input[contains(@placeholder, 'Mobile')]"
        ])

        if name_input:
            name_input.clear()
            name_input.send_keys(name)
        if phone_input:
            phone_input.clear()
            phone_input.send_keys(mobile)
        print("Filled credentials.")

        #Continue
        try:
            cont = wait.until(EC.element_to_be_clickable((By.XPATH,
                "//button[contains(., 'Continue') or contains(., 'Proceed')]")))
            safe_click(driver, cont)
        except Exception:
            print("Continue button not found.")
        time.sleep(2)

        #Quiz selection
        selects = driver.find_elements(By.TAG_NAME, "select")
        quiz_select = None
        for sel in selects:
            s = Select(sel)
            if any("select" not in opt.text.lower() for opt in s.options):
                quiz_select = s
                break

        if not quiz_select:
            print("❌ No quiz dropdown found.")
            return "failed"

        for opt in quiz_select.options:
            if "select" not in opt.text.lower():
                quiz_select.select_by_visible_text(opt.text)
                print("Selected quiz:", opt.text.strip())
                break

        time.sleep(1)

        #Check for any popup alert
        try:
            alert = driver.switch_to.alert
            alert_text = alert.text
            print(f"⚠️ Popup for {name}: {alert_text}")
            alert.accept()

            try:
                logout_btn = driver.find_element(By.XPATH, "//a[contains(., 'Logout')] | //button[contains(., 'Logout')]")
                safe_click(driver, logout_btn)
                print(f"⏭️ Skipped {name} due to popup.")
            except Exception:
                print("Logout not found after popup.")
            return "already_done"

        except NoAlertPresentException:
            pass

        #Start the Quiz
        start_btn = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//*[contains(text(),'Start Quiz') or contains(text(),'Start')]")))
        safe_click(driver, start_btn)
        print("Started quiz.")

        time.sleep(2)
        qn = 0
        while True:
            try:
                option_elements = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".option_list .option"))
                )
                if not option_elements:
                    break

                if chosen_index < len(option_elements):
                    chosen = option_elements[chosen_index]
                else:
                    chosen = option_elements[-1]
                    print(f"⚠️ Only {len(option_elements)} options found. Using last option.")

                safe_click(driver, chosen)
                print(f"Q{qn+1}: clicked {ans_word} option")

                # next or skip
                try:
                    nav_btns = WebDriverWait(driver, 5).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".next_btn, .skip_btn"))
                    )
                    btn_to_click = None
                    for b in nav_btns:
                        if "next" in b.get_attribute("class"):
                            btn_to_click = b
                            break
                        elif "skip" in b.get_attribute("class"):
                            btn_to_click = b
                    if not btn_to_click or not safe_click(driver, btn_to_click, retries=3):
                        print("⚠️ Couldn’t click Next/Skip. Exiting loop.")
                        break
                    print("➡️ Moved to next question.")
                except TimeoutException:
                    print("⚠️ No Next/Skip found. Probably last question.")
                    break

                qn += 1
                time.sleep(1)
            except TimeoutException:
                print("❌ Timeout waiting for question.")
                break

        #Logout
        try:
            logout_btn = driver.find_element(By.XPATH, "//a[contains(., 'Logout')] | //button[contains(., 'Logout')]")
            safe_click(driver, logout_btn)
            print("Logged out.")
        except Exception:
            print("Logout not found.")

        print(f"✅ Quiz finished for {name}. Questions answered: {qn}")
        time.sleep(2)
        return "completed"

    except Exception as e:
        print(f"❌ Unexpected error in run_quiz_for_user: {e}")
        return "failed"


def main():
    ans_word = input("Type your answer for all questions (first/second/third/fourth): ").strip().lower()
    mapping = {"first": 1, "second": 2, "third": 3, "fourth": 4}
    chosen_index = mapping.get(ans_word, 1) - 1
    print(f"✅ Using '{ans_word}' option for all questions.\n")

    users = load_users_from_excel(EXCEL_FILE)
    done_users = load_progress()
    results = {"completed": [], "already_done": [], "failed": []}

    print(f"📘 Found {len(users)} users in Excel. Already completed: {len(done_users)}.\n")

    driver = create_driver()
    wait = WebDriverWait(driver, MAX_WAIT)

    for name, mobile in users:
        if mobile in done_users:
            if VERIFY_DONE_USERS:
                print(f"🔍Verifying if any previous completition - {name} ({mobile})...")
                try:
                    driver.get(QUIZ_URL)
                    status = run_quiz_for_user(driver, wait, name, mobile, chosen_index, ans_word)
                    if status == "already_done":
                        print(f"⏭️ Confirmed from site: {name} has already completed earlier.")
                        results["already_done"].append(name)
                    elif status == "completed":
                        print(f"✅ REVERIFYING - {name} has already completed.")
                        results["completed"].append(name)
                    else:
                        print(f"⚠️ {name} marked done locally but failed on recheck.")
                        results["failed"].append(name)
                    continue
                except Exception as e:
                    print(f"❌ Verification error for {name}: {e}")
                    results["failed"].append(name)
                    continue
            else:
                print(f"⏩ Skipping {name} ({mobile}) — already marked done (no verification).")
                results["already_done"].append(name)
                continue

        try:
            driver.get(QUIZ_URL)
            status = run_quiz_for_user(driver, wait, name, mobile, chosen_index, ans_word)

            if status == "completed":
                results["completed"].append(name)
                done_users.add(mobile)
            elif status == "already_done":
                results["already_done"].append(name)
                done_users.add(mobile)
            else:
                results["failed"].append(name)

            save_progress(done_users)

        except Exception as e:
            print(f"❌ Error for {name}: {e}")
            results["failed"].append(name)
            #Restarting the browser
            try:
                driver.quit()
            except Exception:
                pass
            print("🔄 Restarting Chrome...")
            driver = create_driver()
            wait = WebDriverWait(driver, MAX_WAIT)

    driver.quit()

    total = len(users)
    completed = len(results["completed"])
    already_done = len(results["already_done"])
    failed = len(results["failed"])

    print("\n📊 Summary:")
    print("---------------------------------")
    print(f"Total users: {total}")
    print(f"✅ Completed via automation: {completed}")
    print(f"⏭️ Already completed earlier: {already_done}")
    print(f"❌ Failed due to error: {failed}")
    print("---------------------------------")
    print(f"Sum check: {completed + already_done + failed} = {total}\n")

    if results["failed"]:
        print("🚫 Failed Users:")
        for name in results["failed"]:
            print(" -", name)


if __name__ == "__main__":
    main()
