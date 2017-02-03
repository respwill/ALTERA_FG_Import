#Job 8:7
#Though your beginning was small, yet your latter end would greatly increase.
import sys

sys.path.append("D:\Python")
from HI_tool import emes_login
from selenium import webdriver
import pandas as pd
import os

class fgParser():
    def __init__(self):
        self.target_fg_df = pd.read_excel("target_fg.xlsx")
        self.target_pin_df = pd.DataFrame(columns=["pin"])
        self.result_df = pd.DataFrame(columns=["FG", "PDL", "Target device", "Option ID", "Test time1", "Test time2", "Test time3", "Test time4", "MD", "Structure"])
        self.tt1 = ""
        self.tt2 = ""
        self.tt3 = ""
        self.tt4 = ""
        self.driver = webdriver.PhantomJS()
        login = emes_login.access(self.driver)
        login.connecting()

    def get_pin(self):
        self.driver.get("http://aak1ws01/eMES/testpdb/PINDisplay.jsp")
        self.fg_input = self.driver.find_element_by_name("fg")
        self.find_button = self.driver.find_element_by_name("find")
        self.check_radio_button = self.driver.find_element_by_name("check")
        for fg in self.target_fg_df["Target FG"]:
            print("getting pin from {}".format(str(fg)))
            self.fg_input.clear()
            self.fg_input.send_keys(str(fg))
            self.check_radio_button.click()
            self.find_button.click()
            self.elements = self.driver.find_elements_by_css_selector("p > a > span > font")
            self.target_pin_df.set_value(len(self.target_pin_df),"pin",self.elements[0].text)
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.fg_input = self.driver.find_element_by_name("fg")
            self.find_button = self.driver.find_element_by_name("find")
            self.check_radio_button = self.driver.find_element_by_name("check")

    def get_pin_info(self):
        for pin in self.target_pin_df["pin"]:
            print("getting information from {}".format(str(pin)))
            self.driver.get("http://aak1ws01/eMES/testpdb/pinView.do?PinNo={}".format(str(pin)))
            self.element2 = self.driver.find_elements_by_css_selector("p > span > font")
            for i, d in enumerate(self.element2):
                if d.text.replace(" ","") == "TEST1":
                    self.tt_start_location = self.element2[i+1].text.find("TEST TIME:")
                    self.tt_end_location = self.element2[i+1].text.find(",",self.tt_start_location)
                    self.tt1 = self.element2[i+1].text[self.tt_start_location+10 : self.tt_end_location]
                elif d.text.replace(" ","") == "TEST2":
                    self.tt_start_location = self.element2[i+1].text.find("TEST TIME:")
                    self.tt_end_location = self.element2[i+1].text.find(",",self.tt_start_location)
                    self.tt2 = self.element2[i+1].text[self.tt_start_location+10 : self.tt_end_location]
                elif d.text.replace(" ","") == "TEST3":
                    self.tt_start_location = self.element2[i+1].text.find("TEST TIME:")
                    self.tt_end_location = self.element2[i+1].text.find(",",self.tt_start_location)
                    self.tt3 = self.element2[i+1].text[self.tt_start_location+10 : self.tt_end_location]
                elif d.text.replace(" ","") == "TEST4":
                    self.tt_start_location = self.element2[i+1].text.find("TEST TIME:")
                    self.tt_end_location = self.element2[i+1].text.find(",",self.tt_start_location)
                    self.tt4 = self.element2[i+1].text[self.tt_start_location+10 : self.tt_end_location]
                else:
                    continue

            self.pdl = self.element2[8].text.strip()
            self.structure_code = self.element2[23].text.strip()

            self.element3 = self.driver.find_elements_by_css_selector("td > pre > font")
            self.target_device = self.element3[0].text.replace(" ","")
            self.md_end_point = self.element3[2].text.find("/")
            self.md = self.element3[2].text[:self.md_end_point].replace(" ","")

            self.element4 = self.driver.find_elements_by_css_selector("p > a > span > font")
            self.fg_end_location = self.element4[2].text.find("/")
            self.fg = self.element4[2].text[:self.fg_end_location].replace(" ","")

            self.element5 = self.driver.find_elements_by_css_selector("p > font > span")
            self.option_id = self.element5[0].text.strip()

            # "FG", "PDL", "Target device", "Option ID", "Test time1", "Test time2", "Test time3", "Test time4", "MD", "Structure"
            self.result_df.loc[len(self.result_df)] = self.fg, self.pdl, self.target_device, self.option_id, self.tt1, self.tt2, self.tt3, self.tt4, self.md, self.structure_code

        writer = pd.ExcelWriter("{}/result/{}".format(os.getcwd(),"result.xlsx"),engine="xlsxwriter")


        if "result" not in os.listdir(os.getcwd()):
            os.mkdir(os.path.join(os.getcwd(),"result"))
            self.result_df.to_excel(writer,"result")
        else:
            self.result_df.to_excel(writer,"result")
        work_sheet = writer.sheets["result"]
        work_sheet.set_column("B:B",13)
        work_sheet.set_column("C:C",13)
        work_sheet.set_column("D:D",22)
        work_sheet.set_column("E:E",13)
        work_sheet.set_column("F:I",13)
        work_sheet.set_column("J:K",24)
        writer.save()
        writer.close()

altera = fgParser()
altera.get_pin()
altera.get_pin_info()