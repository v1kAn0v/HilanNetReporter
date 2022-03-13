# from hilannet import HilanNet
from UI import SimpleForm
import datetime
import os
from time import sleep
import wget
from zipfile import ZipFile
from pathlib import Path
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from win32com.client import Dispatch
from selenium import webdriver


class InstallDriver:
    __version = None
    __CHROME_PATH = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                     r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    __WIN_CHROME_URL = "https://chromedriver.storage.googleapis.com/{0}/{1}"

    def __init__(self):

        self.__os_name = os.name
        if InstallDriver.__chrome_version() and self.host_os() == "nt":
            self.__url = self.__WIN_CHROME_URL.format(InstallDriver.__chrome_version(), "chromedriver_win32.zip")
            InstallDriver.__install_driver(self.__url, str(Path.home()))
            # InstallDriver.__home_path()

    @staticmethod
    def __home_path():
        env_var = "PATH"
        env_val = r"{0};%PATH%".format(str(Path.home()))
        os.system('SETX {0} "{1}" /M'.format(env_var, env_val))

    def host_os(self):
        return str.lower(self.__os_name)

    @classmethod
    def __chrome_version(cls):
        return list(filter(None, [cls.get_version(p) for p in cls.__CHROME_PATH]))[0]

    @staticmethod
    def get_version(filename):
        parser = Dispatch("Scripting.FileSystemObject")
        try:
            version = parser.GetFileVersion(filename)
        except Exception:
            return None
        return version

    @classmethod
    def __install_driver(cls, url, driver_dir=None):
        cls.request_driver(url)
        if not driver_dir:
            driver_dir = str(Path.home())
        with ZipFile(str(Path.home()) + r"\chromedriver_win32.zip", 'r') as zip_ref:
            zip_ref.extractall(driver_dir)

    @staticmethod
    def request_driver(url, tempdir=None):
        if not tempdir:
            tempdir = str(Path.home())
        wget.download(url, tempdir)

    def update_chrome(self):
        pass


class Robot(webdriver.Chrome):

    def _close(self):
        super().close()

    def feed_user(self, attr_name=None, attr_val=None, username=None):
        return super().find_element(by=attr_name, value=attr_val).send_keys(username)

    def feed_pass(self, attr_name=None, attr_val=None, password=None):
        return super().find_element(by=attr_name, value=attr_val).send_keys(password)

    def click_button(self, attr_name=None, attr_val=None):
        super().find_element(by=attr_name, value=attr_val).click()

    def get_inner_text(self, attr_name=None, attr_val=None):
        return super().find_element(by=attr_name, value=attr_val).text

    def locate(self, attr_name=None, attr_val=None, selector=None, text=None):
        if attr_name == "id":
            try:
                Robot.wait_for_id(self, attr_val)
            except TimeoutException:
                print("ERROR\tLocator with {0} = {1} has not loaded yet. Try again.".format(attr_name, attr_val))
            finally:
                return super().find_element(by=attr_name, value=attr_val)
        if attr_name == "class name":
            try:
                Robot.wait_for_class(self, attr_val)
            except TimeoutException:
                print("ERROR\tLocator with {0} = {1} has not loaded yet. Try again.".format(attr_name, attr_val))
            finally:
                return super().find_element(by=attr_name, value=attr_val)
        if attr_name == "link text":
            try:
                Robot.wait_for_link(self, attr_val)
            except TimeoutException:
                print("ERROR\tLocator with {0} = {1} has not loaded yet. Try again.".format(attr_name, attr_val))
            finally:
                return super().find_element(by=attr_name, value=attr_val)
        if attr_name == "xpath":
            try:
                Robot.wait_for_xpath(self, attr_val)
            except TimeoutException:
                print("ERROR\tLocator with {0} = {1} has not loaded yet. Try again.".format(attr_name, attr_val))
            finally:
                return super().find_element(by=attr_name, value=attr_val)
        if attr_name == "text":
            try:
                Robot.wait_for_text(self, selector, text)
            except TimeoutException:
                print(
                    f"ERROR\tLocator with xpath `//{selector}[contains(text(), '{text}')]` has not loaded yet.")
            finally:
                super().find_element("xpath", f"{selector}[contains(text(), '{text}')]")

    def wait_for_link(self, attr_vat=None):
        return WebDriverWait(self, 5).until(EC.presence_of_element_located((By.LINK_TEXT, attr_vat)))

    def wait_for_xpath(self, attr_val=None):
        return WebDriverWait(self, 5).until(EC.element_to_be_clickable((By.XPATH, attr_val)))

    def wait_for_class(self, attr_val=None):
        return WebDriverWait(self, 5).until(EC.presence_of_element_located((By.CLASS_NAME, attr_val)))

    def wait_for_css(self, attr_val=None):
        return WebDriverWait(self, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, attr_val)))

    def wait_for_text(self, selector="*", text=None):
        return WebDriverWait(self, 5).until(
            EC.presence_of_element_located((By.XPATH, f"//{selector}[contains(text(), '{text}')]")))

    def wait_for_id(self, attr_val=None):
        return WebDriverWait(self, 5).until(EC.presence_of_element_located((By.ID, attr_val)))

    def wait_for_staleness(self, element):
        return WebDriverWait(self, 5).until(EC.staleness_of(element))

    def wait_for_option(self, css_val=None):
        try:
            Robot.wait_for_css(self, "option[value='{0}']".format(css_val))
        except TimeoutException:
            print("ERROR\tCSS_SELECTOR {0} = {1} has not loaded yet. Try again.".format("option", css_val))

    def select(self, attr_name, attr_val, option_val=None, staleness=False):
        if staleness:
            Robot.wait_for_staleness(self, Robot.locate(self, attr_name, attr_val))
        else:
            try:
                Robot.wait_for_option(self, option_val)
            except TimeoutException:
                print("ERROR\tCSS_SELECTOR {0} = {1} has not loaded yet. Try again.".format("option", option_val))
            finally:
                Select(Robot.locate(self, attr_name, attr_val)).select_by_value(option_val)

    @staticmethod
    def current_date():
        return datetime.date.today().day


class BadDateException(BaseException):
    pass


class HilanNet(Robot):

    CUSTOMERS = {55: "משרד הבריאות - יחידת המיכון",
                 3109:   "מטריקס בי איי בע''מ",
                 994: "מטריקס אי.טי. בע''מ",
                 14: "מטריקס - תחום מטה מטריקס",
                 408: "בנק הפועלים",
                 3980: "2B SECURE LTD"}

    ORDERS = {416716: "פיתוח מערכת ביג דאטא ומודל AI - הזמנה 4501893248",
              422135: "פיתוח מערכת ביג דאטא ומודל AI - הזמנה :",
              419813: "הקמת מאגר גנומי קליני (4) - הזמנה :",
              411434: "ביג דאטא-תהליכי הכנת המידע למחקר - הזמנה :",
              12918: "תקורה MatrixBI"}

    MISSIONS = {416733: "ש''ע שרשרת הדבקה",
                416766: "ש''ע קורונה לייק תמנע",
                431891: "מנתח מערכות /מיישם מוצר",
                431954: "מפתח בכיר BIG DATA",
                441841: "תחזוקה -ש''ע מנתח מערכות/ מיישם מוצר",
                419907: "ש''ע מנתח מערכות/ מיישם מוצר",
                411506: "טעינות",
                433791: "טעינות - מוקפא",
                94570: "תקורה מטריקס BI",
                452221: "המתנה לפתיחת פעילות - אלביט"}

    def payslip(self, attr):
        try:
            super().wait_for_id(attr).click()
        except TimeoutException:
            print("ERROR:\tInternal Error. Last payslip locator not found.")

    def submit_shift(self):
        super().locate("xpath", "//input[contains(@id, 'btnSave')]").click()

    def submit_modal(self):
        super().locate("id", "alertFrame")
        self.switch_to.frame(0)
        # super().locate("text", text="הנתונים נשמרו בהצלחה", selector="span")
        super().click_button("xpath", "//td[contains(text(), 'אישור')]")
        # self.switch_to.window(current_window)

    @staticmethod
    def parse_time(enter_, exit_):
        enter_ = ":" + enter_.split(":")[0] + enter_.split(":")[1]
        exit_ = exit_.replace(":", "")
        return enter_, exit_

    def attendance(self, workday=None, customer=None, order=None, mission=None, enter_=":1000", exit_="1900",
                   msg_="Code review", predefined=False):
        enter_, exit_ = self.parse_time(enter_, exit_)
        if not workday:
            workday = super().current_date()
        # elif workday < super().current_date():
        #     raise BadDateException("Run update attendance")
        elif workday > super().current_date():
            raise Exception("Can not report attendance for future date")
        super().locate("id", "tabItem_9_3_SpanBackground").click()
        super().locate("link text", "דיווח ועדכון").click()
        work_days = super().find_elements("class name", "dTS")

        for day in work_days:
            if int(day.text) == super().current_date():
                day.click()
                break

        for day in work_days:
            if int(day.text) == workday:
                ac = ActionChains(self)
                ac.double_click(day).perform()
                break

        if not predefined:

            self.select("xpath", "//select[contains(@id,'SymbolId_EmployeeReports_row_0_0')]", "0", True)
            self.select("xpath", "//select[contains(@id,'Step1_EmployeeReports_row_0_0')]", customer)
            self.select("xpath", "//select[contains(@id,'Step2_EmployeeReports_row_0_0')]", order)
            self.select("xpath", "//select[contains(@id,'Step3_EmployeeReports_row_0_0')]", mission)
            super().locate("xpath", "//input[contains(@id,'Entry_EmployeeReports_row_0_0')]").send_keys(enter_)
            super().locate("xpath", "//input[contains(@id,'Exit_EmployeeReports_row_0_0')]").send_keys(exit_)
            super().locate("xpath", "//input[contains(@id,'Comment_EmployeeReports_row_0_0')]").send_keys(msg_)
            self.submit_shift()
            self.submit_modal()


class Reporter(HilanNet):

    def __init__(self):
        super().__init__()
        # __input = input("Enter attendance request: ")

        simple_form.submit()
        # self.__workday = int(__input.split("\t")[0])
        # self.__customer = __input.split("\t")[1]
        # self.__order = __input.split("\t")[2]
        # self.__mission = __input.split("\t")[3]
        # self.__enter = __input.split("\t")[4]
        # self.__exit = __input.split("\t")[5]
        # self.__msg = __input.split("\t")[6]

        self.__workday = int(simple_form.values[0].split("\t")[0])
        self.__customer = simple_form.values[0].split("\t")[1]
        self.__order = simple_form.values[0].split("\t")[2]
        self.__mission = simple_form.values[0].split("\t")[3]
        self.__enter = simple_form.values[0].split("\t")[4]
        self.__exit = simple_form.values[0].split("\t")[5]
        self.__msg = simple_form.values[0].split("\t")[6]

    def report(self):
        super().attendance(self.__workday, self.__customer, self.__order, self.__mission, self.__enter,
                           self.__exit, self.__msg)


if __name__ == '__main__':
    print("main")

simple_form = SimpleForm()
try:
    bot = Reporter()
except WebDriverException:
    installer = InstallDriver()
    bot = Reporter()

try:
    bot.get("https://matrix.net.hilan.co.il/login")
except WebDriverException:
    print("ERROR:\tWeb page unreachable. Check your connection or try later.")

try:
    bot.feed_user('id', 'password_nm', 'j9k+uLGHa~2r_bL_')
    bot.feed_pass('id', 'user_nm', '12422')
    bot.click_button('class name', 'hbutton2-lg')
    try:
        msg = bot.wait_for_class('UserText').text
        print(msg)
    except TimeoutException:
        print("ERROR:\tLogin using given username and password failed. Password was changed.")
except NoSuchElementException:
    print("ERROR:\tInternal Error. Login button locator not found.")
# bot.payslip('ctl00_mp_lnkShowPaySlip')
# bot.back()
# bot.attendance(16)
# bot.attendance()
# bot.attendance(22, 55, 419813, 441841, "09:00", "19:00",
#                   "Presentation of GenomeIL + preparing system to load in TEST + testing Hadassah")
# bot.attendance(15)
# bot.attendance(18)

bot.report()
simple_form.popup()
bot.close()
