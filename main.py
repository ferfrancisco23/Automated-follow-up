import selenium.common.exceptions
from selenium import webdriver
import json
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
from time import sleep
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from urllib.parse import urlparse
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import random

class AutoInitialEmailFollowUp:

    def __init__(self, use_personal_browser):


        try:
            self.service = ChromeService(executable_path=ChromeDriverManager().install())
        except ValueError:
            self.service = ChromeService(executable_path=ChromeDriverManager(driver_version="114.0.5735.90").install())

        self.chrome_options = Options()
        self.use_personal_browser = use_personal_browser
        if use_personal_browser:
            profile_directory = r'C:\Users\Kev-HP-Pav-Gaming\AppData\Local\Google\Chrome\User Data\ '
            self.chrome_options.add_argument(f'--user-data-dir={profile_directory}')
            self.chrome_options.add_argument(r'--profile-directory=Default')

        self.chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

        self.auto_initial_email_followup_driver = webdriver.Chrome(service=self.service, options=self.chrome_options)


        self.auto_initial_email_followup_driver.get("http://trieste.io")

        self.wait = WebDriverWait(self.auto_initial_email_followup_driver, 10)


        with open("3rd_initial_template.txt", "r", encoding='utf-8') as template:
            self.third_initial_template = template.read()

        # Check if user already logged in to Trieste, else: log in to trieste
        try:
            trieste_username = self.auto_initial_email_followup_driver.find_element(By.ID, "user_email")
        except selenium.common.exceptions.NoSuchElementException:
            print("Already logged in to Trieste")
        else:
            print("Logging in to Trieste...")
            trieste_password = self.auto_initial_email_followup_driver.find_element(By.ID, "user_password")
            login_submit_button = self.auto_initial_email_followup_driver.find_element(By.NAME, "commit")

            trieste_username.send_keys(os.environ.get("TRIESTE_USERNAME"))
            trieste_password.send_keys(os.environ.get("TRIESTE_PASSWORD"))
            login_submit_button.click()


    def auto_third_initial(self, url):
        # Open lead, click first "Resend" link from the top of the page
        self.auto_initial_email_followup_driver.execute_script("window.open('');")
        self.auto_initial_email_followup_driver.switch_to.window(self.auto_initial_email_followup_driver.window_handles[-1])

        self.auto_initial_email_followup_driver.get(url)

        resend_button = self.wait.until(ec.element_to_be_clickable((By.LINK_TEXT, 'Resend')))
        resend_button.click()
        contact_form_textbox = self.auto_initial_email_followup_driver.find_element(By.ID, 'site_link_contact_url')

        email_string = self.scrape_email_body()
        email_text_box = self.auto_initial_email_followup_driver.find_element(By.ID, 'email_body')

        self.generate_final_email(email_string=email_string)
        email_text_box.send_keys(self.generate_final_email(email_string=email_string))

        # check if contact form is empty, click show and move on to next lead with new tab if not
        if len(contact_form_textbox.get_attribute('value')) > 0:
            print(f"{self.auto_initial_email_followup_driver.current_url} send button not clicked")
            show_contact_form = self.auto_initial_email_followup_driver.find_element(By.XPATH, '//*[@id="site_properties"]/ul/li[5]/div[2]/a')

            with open("contact_form_3rd_initial.txt", "a") as writer:
                writer.write(f"{self.auto_initial_email_followup_driver.current_url}\n")

            show_contact_form.click()
        else:
            email_address_selector = Select(self.auto_initial_email_followup_driver.find_element(By.ID, 'email_email_account_id'))
            current_email_address_index = int(email_address_selector.first_selected_option.get_attribute('value'))

            current_email_address_index += 2

            if current_email_address_index > 5:
                current_email_address_index -= 6

            email_address_selector.select_by_index(current_email_address_index)
            change_email_tickbox = self.auto_initial_email_followup_driver.find_element(By.ID, 'site_link_change_email_account')
            change_email_tickbox.click()

            # add click send button here
            send_button = self.auto_initial_email_followup_driver.find_element(By.ID, 'send_email_submit')
            if not self.use_personal_browser:
                send_button.click()
                print(f"{self.auto_initial_email_followup_driver.current_url} send button clicked")

    def scrape_email_body(self):
        # check to make sure resend template is correct before adding third initial template

        email_text = self.auto_initial_email_followup_driver.find_element(By.ID, 'email_body').get_attribute('value')

        # recurse if variable is empty, or if 'email signature' tag is found
        if "<<<email_signature>>>" in email_text or email_text == "":
            sleep(2)
            return self.scrape_email_body()
        else:
            return email_text


    def check_email_before_append_2nd_initial(self):

        email_text = self.auto_initial_email_followup_driver.find_element(By.ID, 'email_body').get_attribute('value')

        # recurse if variable is empty, or if 'email signature' tag is found
        if "<<<email_signature>>>" in email_text or email_text == "":
            sleep(2)
            return self.check_email_before_append_2nd_initial()
        elif email_text.startswith("Hi"):
            return email_text
        else:
            sleep(2)
            return self.scrape_email_body()

    def check_email_before_send_2nd_initial(self):

        email_text = self.auto_initial_email_followup_driver.find_element(By.ID, 'email_body').get_attribute('value')

        # recurse if variable is empty, or if 'email signature' tag is found
        if "Yours," in email_text or "Warm regards," in email_text:
            return True
        elif "Kind regards," in email_text or "Best," in email_text:
            return True
        else:
            sleep(2)
            return self.check_email_before_send_2nd_initial()

    def generate_final_email(self, email_string):

        first_comma_index = email_string.index("Hi") + email_string[email_string.index("Hi"):].index(",")
        signature_index = email_string.index("Faith Cormier")
        lead_url = self.auto_initial_email_followup_driver.find_element(By.ID, 'site_link_from_url').get_attribute('value')

        final_email = self.third_initial_template + email_string[signature_index:]
        final_email = final_email.replace("XXXXXXX", self.extract_domain(lead_url))
        return final_email

    def extract_domain(self, url):

        parsed_url = urlparse(url)
        domain = parsed_url.netloc
        return domain

    def auto_2nd_intial(self, lead_url):
        self.auto_initial_email_followup_driver.execute_script("window.open('');")
        self.auto_initial_email_followup_driver.switch_to.window(self.auto_initial_email_followup_driver.window_handles[-1])
        self.auto_initial_email_followup_driver.get(lead_url)


        # check for resend
        try:
            resend_button = self.wait.until(ec.visibility_of_all_elements_located((By.LINK_TEXT, 'Resend')))
        except selenium.common.exceptions.TimeoutException:
            with open("2nd_initial_contact_form.txt", 'a', encoding='UTF-8') as second_initial_writer:
                second_initial_writer.write(f"{self.auto_initial_email_followup_driver.current_url}\n")
        else:
            resend_button[0].click()

            contact_form_textbox = self.auto_initial_email_followup_driver.find_element(By.ID, 'site_link_contact_url')

            if self.check_email_before_append_2nd_initial():

                email_select_dropdown = Select(
                    self.auto_initial_email_followup_driver.find_element(By.ID, 'email_email_account_id'))
                selected_email = email_select_dropdown.first_selected_option.text

                # select proper textbox 2nd initial based on email address presently used
                if selected_email == "faithcormierlgw@gmail.com":
                    template_select = Select(
                        self.auto_initial_email_followup_driver.find_element(By.XPATH, '//*[@id="email_template_id_810"]'))
                    template_select.select_by_visible_text("[Wise OG] 2 Initial Email Follow Up")
                elif selected_email == "faith.cormier@wise-marketing.co.uk":
                    template_select = Select(
                        self.auto_initial_email_followup_driver.find_element(By.XPATH, '//*[@id="email_template_id_642"]'))
                    template_select.select_by_visible_text("[Wise WMG] 2 Initial Email Follow Up")
                elif selected_email == "f.cormier@wise-marketing.co.uk":
                    template_select = Select(
                        self.auto_initial_email_followup_driver.find_element(By.XPATH, '//*[@id="email_template_id_839"]'))
                    template_select.select_by_visible_text("[Wise WMG] 2 Initial Email Follow Up")
                elif selected_email == "faithcormierlgw@outlook.com" or selected_email == "faith.cormier@letsgetwise.co.uk":
                    email_select_dropdown.select_by_visible_text("faith.cormier@wise-outreach.co.uk")

                    template_select = Select(
                        self.auto_initial_email_followup_driver.find_element(By.XPATH,
                                                                             '//*[@id="email_template_id_591"]'))
                    template_select.select_by_visible_text("[Wise WO] 2 Initial Email Follow Up")

                    change_email_tickbox = self.auto_initial_email_followup_driver.find_element(By.ID,
                                                                                                'site_link_change_email_account')
                    change_email_tickbox.click()

                    # template_select = Select(
                    #     self.auto_initial_email_followup_driver.find_element(By.XPATH, '//*[@id="email_template_id_1046"]'))
                    # template_select.select_by_visible_text("[Wise OG] 2 Initial Email Follow Up")
                # elif selected_email == "faith.cormier@letsgetwise.co.uk":
                #     template_select = Select(
                #         self.auto_initial_email_followup_driver.find_element(By.XPATH, '//*[@id="email_template_id_417"]'))
                #     template_select.select_by_visible_text("[Wise LGW] 2 Initial Email Follow Up")
                elif selected_email == "faith.cormier@wise-outreach.co.uk":
                    template_select = Select(
                        self.auto_initial_email_followup_driver.find_element(By.XPATH, '//*[@id="email_template_id_591"]'))
                    template_select.select_by_visible_text("[Wise WO] 2 Initial Email Follow Up")


            if len(contact_form_textbox.get_attribute('value')) > 0 or len(resend_button) > 1:
                with open("2nd_initial_contact_form.txt", 'a', encoding='UTF-8') as second_initial_writer:
                    second_initial_writer.write(f"{self.auto_initial_email_followup_driver.current_url}\n")
                show_contact_form = self.auto_initial_email_followup_driver.find_element(By.XPATH, '//*[@id="site_properties"]/ul/li[5]/div[2]/a')
                show_contact_form.click()
                print(f"{lead_url} Send not clicked")

            else:

                if self.check_email_before_send_2nd_initial():
                    submit_button = self.auto_initial_email_followup_driver.find_element(By.ID, 'send_email_submit')
                    if not self.use_personal_browser:
                        sleep_timer = random.randint(60, 125)
                        sleep(sleep_timer)
                        submit_button.click()
                        print(f"{lead_url} Send clicked")




site_links = []
with open("config.json", "r") as config:
    config_data = json.load(config)

delay_per_lead = config_data["delay"]

user_task = input("What would you like to do?\n")

user_input = input("Paste site link urls here (1 keyword(s) per line): ")

# allow user to input keywords, 1 line per keyword
while user_input != '':
    site_links.append(user_input)
    user_input = input()

if user_task == "2nd_initial_contact_form":
    Auto_Follow_Up = AutoInitialEmailFollowUp(use_personal_browser=True)
    for site_link in site_links:
        Auto_Follow_Up.auto_2nd_intial(lead_url=site_link)
elif user_task == "3rd_initial_contact_form":
    Auto_Follow_Up = AutoInitialEmailFollowUp(use_personal_browser=True)
    for site_link in site_links:
        Auto_Follow_Up.auto_third_initial(url=site_link)
elif user_task == "auto_2nd_initial":
    Auto_Follow_Up = AutoInitialEmailFollowUp(use_personal_browser=False)
    for site_link in site_links:
        Auto_Follow_Up.auto_2nd_intial(lead_url=site_link)

elif user_task == "auto_3rd_initial":
    Auto_Follow_Up = AutoInitialEmailFollowUp(use_personal_browser=False)
    for site_link in site_links:
        Auto_Follow_Up.auto_third_initial(url=site_link)
        sleep(delay_per_lead)


input("continue?")


#  raise exception_class(message, screen, stacktrace, alert_text)  # type: ignore[call-arg]  # mypy is not smart enough here
# selenium.common.exceptions.UnexpectedAlertPresentException: Alert Text: Please correct all dynamic tags before sending this email
# Message: unexpected alert open: {Alert text : Please correct all dynamic tags before sending this email}