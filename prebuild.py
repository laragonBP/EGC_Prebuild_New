# -*- coding: utf-8 -*-

from selenium import webdriver
import time
import pandas as pd

class EnterpriseGreenPrebuildCrawler:

    login_url = 'https://enterprisecommunity.secure.force.com/greenportal/ENT_SiteLogin'
    indices_list = [
        [0, 1, 2, 3, 4, 5, 6],
        [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
        [0, 1, 2, 3, 4, 5, 6, 7],
        [0, 1, 2, 3, 4, 5],
        [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
        [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
        [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16],
        [0, 1, 2, 3, 4, 5]

    ]

    row = 0

    def __init__(self, email, password, project_url, excel_file_path):
        self.email = email
        self.password = password
        self.web_driver = webdriver.Chrome('chromedriver.exe')
        self.project_url = project_url
        self.excel_file_path = excel_file_path


    def fill_in_form(self):
        try:
            self._read_excel_file()
            self._login()  # Logs in to EGC portal and enters project URLs
            self._fill_in_will_not_pursue_credits()
            # self._save_data()
        except Exception as e:
            print(e)
        finally:
            self._quit_chrome_session()


    def _read_excel_file(self):
        self.all_form_data = pd.read_excel(excel_file_path,'For Python',header=0,converters={'how_implemented':str,'documentation':str,'plan_type':str,'champion':str,'additional_information':str})

    def _login(self):
        self.web_driver.get(self.login_url) # enters login_url in web driver
        self.web_driver.find_element_by_id('pg:SiteTemplate:loginForm:username').send_keys(self.email) #enters username in login page
        self.web_driver.find_element_by_id('pg:SiteTemplate:loginForm:passwd').send_keys(self.password) #enters password in login page
        self.web_driver.find_element_by_name('pg:SiteTemplate:loginForm:j_id25').click() #clicks on Login
        #needs a delay before going to project page
        time.sleep(1)
        self.web_driver.get(self.project_url) #Goes to project URL
        time.sleep(1)

    def open_all_accordians(self): # Opens accordians
        accordian_list = self.web_driver.find_elements_by_css_selector('.accordion.groupKey')
        for accordian in accordian_list:
            accordian.click()
            time.sleep(1)

    def _fill_in_will_not_pursue_credits(self):
        self.open_all_accordians()
        for section_index, form_index_list in enumerate(self.indices_list):
            for form_index in form_index_list:
                try:
                    if not self.read_only(section_index, form_index):
                        self.fill_in_single_credit_form(section_index, form_index, self.all_form_data, self.row)
                except Exception as e:
                    print('Could not set data for section {0}, form index {1}'.format(section_index + 1, form_index))
                    print(e)
                    time.sleep(1)
                self.row = self.row + 1

    def read_only(self, form_index, field_index):
        base_id = 'j_id0:SiteTemplate:j_id54:j_id60:{0}:j_id63:{1}:otherOpt'
        return not self.web_driver.find_element_by_id(base_id.format(form_index, field_index)).is_enabled()

    def fill_in_single_credit_form(self, section_index, form_index, form_data, row):
        self.fill_in_how_implemented(section_index, form_index, 'N/A' if form_data.loc[row, 'how_implemented']=="'N/A'" else form_data.loc[row, 'how_implemented'])
        self.fill_in_documentation(section_index, form_index, 'N/A' if form_data.loc[row, 'documentation']=="'N/A'" else form_data.loc[row, 'documentation'])
        self.fill_in_champion(section_index, form_index, 'N/A' if form_data.loc[row, 'champion']=="'N/A'" else form_data.loc[row, 'champion'])
        self.fill_in_plan_type(section_index, form_index, 'N/A' if form_data.loc[row, 'plan_type']=="'N/A'" else form_data.loc[row, 'plan_type'])
        self.fill_in_additional_information(section_index, form_index, 'N/A' if form_data.loc[row, 'additional_information']=="'N/A'" else form_data.loc[row, 'additional_information'])

    def fill_in_how_implemented(self, section_index, form_index, dropdown_text):
        base_id = 'j_id0:SiteTemplate:j_id54:j_id60:{0}:j_id63:{1}:how_implemented'
        self.select_from_dropdown(base_id.format(section_index, form_index), dropdown_text)

    def fill_in_documentation(self, section_index, form_index, dropdown_text):
        base_id = 'j_id0:SiteTemplate:j_id54:j_id60:{0}:j_id63:{1}:location_in_project_docs'
        self.select_from_dropdown(base_id.format(section_index, form_index), dropdown_text)

    def fill_in_champion(self, section_index, form_index, dropdown_text):
        base_id = 'j_id0:SiteTemplate:j_id54:j_id60:{0}:j_id63:{1}:champion'
        self.select_from_dropdown(base_id.format(section_index, form_index), dropdown_text)

    def fill_in_plan_type(self, section_index, form_index, field_text):
        base_id = 'j_id0:SiteTemplate:j_id54:j_id60:{0}:j_id63:{1}:spec_page_number_plan_type'
        self.web_driver.find_element_by_id(base_id.format(section_index, form_index)).clear()
        self.web_driver.find_element_by_id(base_id.format(section_index, form_index)).send_keys(field_text)

    def fill_in_additional_information(self, section_index, form_index, field_text):
        base_id = 'j_id0:SiteTemplate:j_id54:j_id60:{0}:j_id63:{1}:otherOpt'
        self.web_driver.find_element_by_id(base_id.format(section_index, form_index)).clear()
        self.web_driver.find_element_by_id(base_id.format(section_index, form_index)).send_keys(field_text)

    def select_from_dropdown(self, dropdown_id, option_text):
        dropdown = self.web_driver.find_element_by_id(dropdown_id)
        dropdown.click()
        dropdown.find_elements_by_xpath('.//*[contains(text(), "{0}")]'.format(option_text))[0].click()
        dropdown.click()

    def _save_data(self):
        self.web_driver.find_element_by_name('j_id0:SiteTemplate:j_id54:j_id101').click()
        print('Will not pursue credit fields saved.')

    def _quit_chrome_session(self):
        self.web_driver.quit()



if __name__ == '__main__':
    email = 'egc@brightpower.com'
    password = 'bp2128035868'
    project_url = 'https://enterprisecommunity.secure.force.com/greenportal/ENT_IntendedMethods?id=a4U1O0000000kvMUAQ&status=pre'
    excel_file_path = 'C:/Users/lgonzalez/Documents/Python Scripts/23_11_EGC_Checklist_Portal_Template_New_Construction.xlsx'

    crawler = EnterpriseGreenPrebuildCrawler(email, password, project_url, excel_file_path)
    try:
        crawler.fill_in_form()
    except Exception as e:
        print(e)

