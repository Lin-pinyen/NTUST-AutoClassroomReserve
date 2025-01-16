import json
import datetime
import math
import time
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.common.exceptions import NoSuchElementException

# =============== Utility Functions ===============

def load_settings(file_path):
    """Load parameters from the setting.txt file."""
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)

def get_week_of_month(year, month, day):
    """Calculate the week of the month for a given date."""
    first_day = datetime.date(year, month, 1)
    target_day = datetime.date(year, month, day)
    first_day_weekday = first_day.weekday()
    days_difference = (target_day - first_day).days
    return math.ceil((days_difference + first_day_weekday + 1) / 7)

def element_exists(driver, css_selector):
    """Check if element with given CSS selector exists."""
    try:
        driver.find_element(By.CSS_SELECTOR, css_selector)
        return True
    except NoSuchElementException:
        return False

def select_month_on_datepicker(driver, target_month_name):
    """Navigate datepicker to the target month."""
    while True:
        current_month_name = driver.execute_script(
            "return document.querySelector('#ui-datepicker-div .ui-datepicker-month').textContent;"
        )
        if current_month_name == target_month_name:
            break
        next_button = driver.find_element(By.CSS_SELECTOR, "#ui-datepicker-div .ui-datepicker-next")
        driver.execute_script("arguments[0].click();", next_button)
        time.sleep(1)

def set_time_slider(driver, target_hour, target_minute):
    # ----------------- Hour Slider -----------------
    min_hour = 9   # Earliest hour displayed in your system (adjust as needed)
    max_hour = 16  # Latest hour. If your system is 9–16, total range is 7
    hour_range = max_hour - min_hour

    # Calculate left percentage for the hour
    left_percentage_per_hour = 100 / hour_range
    left_hour_value = (target_hour - min_hour) * left_percentage_per_hour

    # 1) Directly set style.left
    slider_handle = driver.find_element(
        By.CSS_SELECTOR,
        "#ui-datepicker-div > div.ui-timepicker-div > dl > dd.ui_tpicker_hour > "
        "div.ui-slider.ui-slider-horizontal.ui-widget.ui-widget-content.ui-corner-all > "
        "a.ui-slider-handle.ui-state-default.ui-corner-all"
    )
    driver.execute_script(
        """
        arguments[0].style.left = arguments[1] + '%';
        arguments[0].setAttribute('aria-valuenow', arguments[1]);
        arguments[0].dispatchEvent(new Event('input'));
        arguments[0].dispatchEvent(new Event('change'));
        """,
        slider_handle,
        left_hour_value
    )

    # 2) Simulate mouse drag to ensure the UI updates fully
    hour_slider = driver.find_element(
        By.CSS_SELECTOR,
        "#ui-datepicker-div > div.ui-timepicker-div > dl > dd.ui_tpicker_hour > div.ui_tpicker_hour_slider"
    )
    hour_slider_handle = driver.find_element(
        By.CSS_SELECTOR,
        "#ui-datepicker-div > div.ui-timepicker-div > dl > dd.ui_tpicker_hour > div.ui_tpicker_hour_slider > a.ui-slider-handle"
    )

    driver.execute_script(
        """
        const slider = arguments[0];
        const handle = arguments[1];
        const targetLeft = arguments[2];

        // Trigger mousedown
        const mousedownEvent = new MouseEvent('mousedown', {
            bubbles: true,
            cancelable: true,
            clientX: slider.getBoundingClientRect().left + (targetLeft / 100) * slider.offsetWidth,
            clientY: slider.getBoundingClientRect().top
        });
        handle.dispatchEvent(mousedownEvent);

        // Trigger mousemove
        const mousemoveEvent = new MouseEvent('mousemove', {
            bubbles: true,
            cancelable: true,
            clientX: slider.getBoundingClientRect().left + (targetLeft / 100) * slider.offsetWidth,
            clientY: slider.getBoundingClientRect().top
        });
        slider.dispatchEvent(mousemoveEvent);

        // Trigger mouseup
        const mouseupEvent = new MouseEvent('mouseup', {
            bubbles: true,
            cancelable: true
        });
        handle.dispatchEvent(mouseupEvent);
        """,
        hour_slider,
        hour_slider_handle,
        left_hour_value
    )
    time.sleep(2)

    # ----------------- Minute Slider -----------------
    # For your system, you used minute_range = 50, but normally we’d expect 0–59 or 0–60.
    # Adjust as needed. (If your slider has discrete steps, it might top out at 50.)
    min_minute = 0
    max_minute = 50  # or 60, depending on the timepicker’s step
    minute_range = max_minute - min_minute

    left_percentage_per_minute = 100 / minute_range
    left_minute_value = (target_minute - min_minute) * left_percentage_per_minute

    # 1) Directly set style.left
    minute_slider_handle = driver.find_element(
        By.CSS_SELECTOR,
        "#ui-datepicker-div > div.ui-timepicker-div > dl > dd.ui_tpicker_minute > "
        "div.ui-slider-horizontal > a.ui-slider-handle"
    )
    driver.execute_script(
        """
        arguments[0].style.left = arguments[1] + '%';
        arguments[0].setAttribute('aria-valuenow', arguments[1]);
        arguments[0].dispatchEvent(new Event('input'));
        arguments[0].dispatchEvent(new Event('change'));
        """,
        minute_slider_handle,
        left_minute_value
    )

    # 2) Simulate mouse drag
    minute_slider = driver.find_element(
        By.CSS_SELECTOR,
        "#ui-datepicker-div > div.ui-timepicker-div > dl > dd.ui_tpicker_minute > div.ui-slider-horizontal"
    )
    minute_slider_handle = driver.find_element(
        By.CSS_SELECTOR,
        "#ui-datepicker-div > div.ui-timepicker-div > dl > dd.ui_tpicker_minute > div.ui-slider-horizontal > a.ui-slider-handle"
    )

    driver.execute_script(
        """
        const slider = arguments[0];
        const handle = arguments[1];
        const targetLeft = arguments[2];

        // Trigger mousedown
        const mousedownEvent = new MouseEvent('mousedown', {
            bubbles: true,
            cancelable: true,
            clientX: slider.getBoundingClientRect().left + (targetLeft / 100) * slider.offsetWidth,
            clientY: slider.getBoundingClientRect().top
        });
        handle.dispatchEvent(mousedownEvent);

        // Trigger mousemove
        const mousemoveEvent = new MouseEvent('mousemove', {
            bubbles: true,
            cancelable: true,
            clientX: slider.getBoundingClientRect().left + (targetLeft / 100) * slider.offsetWidth,
            clientY: slider.getBoundingClientRect().top
        });
        slider.dispatchEvent(mousemoveEvent);

        // Trigger mouseup
        const mouseupEvent = new MouseEvent('mouseup', {
            bubbles: true,
            cancelable: true
        });
        handle.dispatchEvent(mouseupEvent);
        """,
        minute_slider,
        minute_slider_handle,
        left_minute_value
    )
    time.sleep(2)

def borrow_room_for_date(driver, day_css_selector, reason):
    """Reserve a room for the selected date."""
    day_element = driver.find_element(By.CSS_SELECTOR, day_css_selector)
    driver.execute_script("arguments[0].click();", day_element)
    time.sleep(2)

    set_time_slider(driver, settings["start_time_hour"], settings["start_time_minute"])
    time.sleep(1)

    expiration_button = driver.find_element(By.CSS_SELECTOR, "#LoanExpirationDateBtn")
    driver.execute_script("arguments[0].click();", expiration_button)
    time.sleep(1)

    set_time_slider(driver, settings["end_time_hour"], settings["end_time_minute"])

    actions = ActionChains(driver)
    actions.move_by_offset(1, 1).click().perform()

    borrow_reason = driver.find_element(By.CSS_SELECTOR, "#LoanExtra")
    borrow_reason.clear()
    borrow_reason.send_keys(reason)
    time.sleep(1)

    checkbox = driver.find_element(By.CSS_SELECTOR, "#LoanIsAgree")
    if not checkbox.is_selected():
        checkbox.click()
    time.sleep(1)

    borrow_button = driver.find_element(By.CSS_SELECTOR, "#LoanRoomReserveForm > button.btn.btn-success")
    driver.execute_script("arguments[0].click();", borrow_button)
    time.sleep(2)

def go_to_target_room_page(driver, target_url):
    """Navigate to the target room reservation page."""
    driver.get(target_url)
    time.sleep(2)

# =============== Main Script ===============

def main():
    global settings
    settings = load_settings("setting.txt")

    service = Service(EdgeChromiumDriverManager().install())
    options = webdriver.EdgeOptions()
    driver = webdriver.Edge(service=service, options=options)

    driver.get("https://yo-1.ct.ntust.edu.tw/rms/")
    time.sleep(2)

    login_button_element = driver.find_element(By.CSS_SELECTOR, "#wrapper > nav.navbar.navbar-inverse.navbar-fixed-bottom > div.collapse.navbar-collapse.navbar-ex1-collapse > ul > li:nth-child(3) > a")
    driver.execute_script("arguments[0].click();", login_button_element)
    time.sleep(1)

    driver.find_element(By.CSS_SELECTOR, "#UserEmail").send_keys(settings["username"])
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, "#UserPassword").send_keys(settings["password"])
    driver.find_element(By.CSS_SELECTOR, "#UserLoginForm > p > button").click()
    time.sleep(2)

    target_url = settings["target_room_url"]
    go_to_target_room_page(driver, target_url)

    max_days_to_borrow = settings["max_days_to_borrow"]
    borrow_count = 0
    week_of_month = get_week_of_month(*settings["start_date"])

    for month_name in settings["desired_months"]:
        datepicker_btn = driver.find_element(By.CSS_SELECTOR, "#LoanStartDateBtn")
        driver.execute_script("arguments[0].click();", datepicker_btn)
        time.sleep(1)

        select_month_on_datepicker(driver, month_name)
        time.sleep(1)

        iter_range = range(week_of_month, 7) if borrow_count == 0 else range(1, 7)
        for row in iter_range:
            css_for_day = f"#ui-datepicker-div > table > tbody > tr:nth-child({row}) > td:nth-child({settings['weekday'] + 1}) > a"
            if element_exists(driver, css_for_day):
                try:
                    borrow_room_for_date(driver, css_for_day, settings["reason"])
                    borrow_count += 1

                    if borrow_count < max_days_to_borrow:
                        go_to_target_room_page(driver, target_url)
                    else:
                        print(f"Reached maximum of {max_days_to_borrow} reservations.")
                        driver.quit()
                        return
                    time.sleep(3)
                    datepicker_btn = driver.find_element(By.CSS_SELECTOR, "#LoanStartDateBtn")
                    driver.execute_script("arguments[0].click();", datepicker_btn)
                    time.sleep(1)
                    select_month_on_datepicker(driver, month_name)
                    time.sleep(1)

                except Exception as e:
                    print(f"Error borrowing date at row={row}, col={settings['weekday'] + 1}: {e}")

    print("Finished borrowing logic.")
    driver.quit()

if __name__ == "__main__":
    main()
