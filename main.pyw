#!python3.6
# -*- coding: Windows-1251 -*-
import sys
sys.path.insert(0, 'pkgs')
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox
import design
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import os
import shutil
import csv
import sqlite3
from xlsxwriter.workbook import Workbook


district_list_all = ('001', '801', '802', '803', '805', '807', '812', '813', '815', '501', '502', '504', '505', '506',
                     '507', '508', '509', '510', '511', '512', '513', '701', '702', '703', '705', '707', '708', '709',
                     '710', '711', '712', '713', '714', '715', '716', '717', '201', '202', '203', '205', '206', '207',
                     '208', '209', '210', '211', '212', '213', '214', '215', '216', '20', '45', '32', '42', '132', '4',
                     '11', '27', '227', '327', '427', '527', '7', '36', '56', '156', '6', '23', '47', '50', '150', '5', 
					 '8', '9', '16', '17', '21', '30', '31', '37', '38', '39', '40', '43', '55', '116', '140')

district_list_m = ('001', '801', '802', '803', '805', '807', '812', '813', '815', '501', '502', '504', '505', '506',
                   '507', '508', '509', '510', '511', '512', '513', '701', '702', '703', '705', '707', '708', '709',
                   '710', '711', '712', '713', '714', '715', '716', '717', '201', '202', '203', '205', '206', '207',
                   '208', '209', '210', '211', '212', '213', '214', '215', '216')

district_list_mo = ('20', '45', '32', '42', '132', '4',  '11', '27', '227', '327', '427', '527', '7', '36', '56',
                    '156', '6', '23', '47', '50', '150', '5', '8', '9', '16', '17', '21', 
				   '30', '31', '37', '38', '39', '40', '43', '55', '116', '140')

server_list_all = ('178', '179', '183', '184', '80', '206', '209', '210')

server_list_m = ('178', '179', '183', '184')

server_list_mo = ('80', '206', '209', '210')

gu_list = ('1', '2', '3', '4', '5', 'УПФР')


class MainWindow(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        #szvm
        self.szvmRaList.addItems(district_list_all)
        self.szvmServerList.addItems(server_list_all)
        self.szvmManButton.clicked.connect(self.szvm_man_run)
        self.szvmWprButton.clicked.connect(self.szvm_wpr_run)
        self.szvmPopayButton.clicked.connect(self.szvm_popay_run)
        self.szvmPopenButton.clicked.connect(self.szvm_popen_run)
        self.szvmPeButton.clicked.connect(self.szvm_pe_run)
        #prbz
        self.prbzRaList.addItems(district_list_all)
        self.prbzServerList.addItems(server_list_all)
        self.prbzManButton.clicked.connect(self.prbz_man_run)
        self.prbzWprButton.clicked.connect(self.prbz_wpr_run)
        self.prbzPopayButton.clicked.connect(self.prbz_popay_run)
        self.prbzPopenButton.clicked.connect(self.prbz_popen_run)
        self.prbzPeButton.clicked.connect(self.prbz_pe_run)
        #edv
        self.edvRaList.addItems(district_list_all)
        self.edvServerList.addItems(server_list_all)
        self.edvManButton.clicked.connect(self.edv_man_run)
        self.edvPodmoButton.clicked.connect(self.edv_podmo_run)
        self.edvPopenButton.clicked.connect(self.edv_popen_run)
        self.edvWprButton.clicked.connect(self.edv_wpr_run)
        #deduction
        self.deductionGuList.addItems(gu_list)
        self.deductionMoscowButton.clicked.connect(self.deduction_run_m)
        self.deductionMoButton.clicked.connect(self.deduction_run_mo)

    def properties(self):
        global district_list
        global login_url
        global server_url
        global region
        global man_index
        global pe_index
        global po_index
        global podmo_index
        global popay_index
        global popen_index
        global district
        global status
        global gu
        global download_folder

        tab = self.verticalTabWidget.currentIndex()

        if tab == 0:
            district = self.szvmRaList.currentText()
            server = self.szvmServerList.currentText()
            download_folder = 'C:\VIB'
        elif tab == 1:
            district = self.prbzRaList.currentText()
            server = self.prbzServerList.currentText()
            download_folder = 'C:\VIB'
        elif tab == 2:
            district = self.edvRaList.currentText()
            server = self.edvServerList.currentText()
            download_folder = 'C:\VIB'
        elif tab == 3:
            gu = self.deductionGuList.currentText()
            download_folder = 'D:\ADV8D'

        if tab == 0 or tab == 1 or tab == 2:
            if district in district_list_m and server in server_list_mo:
                QMessageBox.information(self, 'Ошибка', 'Несоответствие района и сервера')
                status = 1
            elif district in district_list_mo and server in server_list_m:
                QMessageBox.information(self, 'Ошибка', 'Несоответствие района и сервера')
                status = 1
            else:
                status = 0

            if server == '178':
                district_list = district_list_m
                login_url = 'http://10.87.0.178/MainWAR/login.html'
                server_url = 'http://10.87.0.178/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
                region = '087'
                man_index = "form1:table1:25:rowSelect1__input_sel"
                pe_index = "form1:table1:32:rowSelect1__input_sel"
                po_index = "form1:table1:24:rowSelect1__input_sel"
                podmo_index = "form1:table1:28:rowSelect1__input_sel"
                popay_index = "form1:table1:31:rowSelect1__input_sel"
                popen_index = "form1:table1:32:rowSelect1__input_sel"
            elif server == '179':
                district_list = district_list_m
                login_url = 'http://10.87.0.179:9080/ViplataWEB/login.jsp'
                server_url = 'http://10.87.0.179:9080/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
                region = '087'
                man_index = "form1:table1:25:rowSelect1__input_sel"
                pe_index = "form1:table1:32:rowSelect1__input_sel"
                po_index = "form1:table1:24:rowSelect1__input_sel"
                podmo_index = "form1:table1:28:rowSelect1__input_sel"
                popay_index = "form1:table1:31:rowSelect1__input_sel"
                popen_index = "form1:table1:32:rowSelect1__input_sel"
            elif server == '183':
                district_list = district_list_m
                login_url = 'http://10.87.0.183:9080/ViplataWEB/login.jsp'
                server_url = 'http://10.87.0.183:9080/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
                region = '087'
                man_index = "form1:table1:25:rowSelect1__input_sel"
                pe_index = "form1:table1:32:rowSelect1__input_sel"
                po_index = "form1:table1:24:rowSelect1__input_sel"
                podmo_index = "form1:table1:28:rowSelect1__input_sel"
                popay_index = "form1:table1:31:rowSelect1__input_sel"
                popen_index = "form1:table1:32:rowSelect1__input_sel"
            elif server == '184':
                district_list = district_list_m
                login_url = 'http://10.87.0.184:9080/ViplataWEB/login.jsp'
                server_url = 'http://10.87.0.184:9080/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
                region = '087'
                man_index = "form1:table1:25:rowSelect1__input_sel"
                pe_index = "form1:table1:32:rowSelect1__input_sel"
                po_index = "form1:table1:24:rowSelect1__input_sel"
                podmo_index = "form1:table1:28:rowSelect1__input_sel"
                popay_index = "form1:table1:31:rowSelect1__input_sel"
                popen_index = "form1:table1:32:rowSelect1__input_sel"
            elif server == '80':
                district_list = district_list_mo
                login_url = 'http://10.87.0.80/MainWAR/login.html'
                server_url = 'http://10.87.0.80/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
                region = '060'
                man_index = "form1:table1:23:rowSelect1__input_sel"
                pe_index = "form1:table1:30:rowSelect1__input_sel"
                po_index = "form1:table1:23:rowSelect1__input_sel"
                podmo_index = "form1:table1:27:rowSelect1__input_sel"
                popay_index = "form1:table1:30:rowSelect1__input_sel"
                popen_index = "form1:table1:31:rowSelect1__input_sel"
            elif server == '206':
                district_list = district_list_mo
                login_url = 'http://10.87.0.206:9080/ViplataWEB/login.jsp'
                server_url = 'http://10.87.0.206:9080/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
                region = '060'
                man_index = "form1:table1:23:rowSelect1__input_sel"
                pe_index = "form1:table1:30:rowSelect1__input_sel"
                po_index = "form1:table1:23:rowSelect1__input_sel"
                podmo_index = "form1:table1:27:rowSelect1__input_sel"
                popay_index = "form1:table1:30:rowSelect1__input_sel"
                popen_index = "form1:table1:31:rowSelect1__input_sel"
            elif server == '209':
                district_list = district_list_mo
                login_url = 'http://10.87.0.209:9080/ViplataWEB/login.jsp'
                server_url = 'http://10.87.0.209:9080/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
                region = '060'
                man_index = "form1:table1:23:rowSelect1__input_sel"
                pe_index = "form1:table1:30:rowSelect1__input_sel"
                po_index = "form1:table1:23:rowSelect1__input_sel"
                podmo_index = "form1:table1:27:rowSelect1__input_sel"
                popay_index = "form1:table1:30:rowSelect1__input_sel"
                popen_index = "form1:table1:31:rowSelect1__input_sel"
            elif server == '210':
                district_list = district_list_mo
                login_url = 'http://10.87.0.210:9080/ViplataWEB/login.jsp'
                server_url = 'http://10.87.0.210:9080/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
                region = '060'
                man_index = "form1:table1:23:rowSelect1__input_sel"
                pe_index = "form1:table1:30:rowSelect1__input_sel"
                po_index = "form1:table1:23:rowSelect1__input_sel"
                podmo_index = "form1:table1:27:rowSelect1__input_sel"
                popay_index = "form1:table1:30:rowSelect1__input_sel"
                popen_index = "form1:table1:31:rowSelect1__input_sel"

    def browser(self, dowload_folder):
        global driver
        binary = r'Mozilla Firefox\firefox.exe'
        options = Options()
        options.headless = False
        options.binary = binary
        cap = DesiredCapabilities().FIREFOX
        cap["marionette"] = True
        profile = webdriver.FirefoxProfile()
        profile.set_preference("browser.download.folderList", 2)
        profile.set_preference("browser.download.manager.showWhenStarting", False)
        profile.set_preference("browser.download.dir", dowload_folder)
        profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv")
        driver = webdriver.Firefox(firefox_options=options, firefox_profile=profile, capabilities=cap, executable_path="geckodriver.exe")

    def login(self, login_url):
        driver.get(login_url)
        username = driver.find_element_by_name("j_username")
        username.send_keys("S88019195")
        password = driver.find_element_by_name("j_password")
        password.send_keys("TOPKEK17")
        log_in = driver.find_element_by_name("action")
        log_in.click()

    def szvm_man(self, server, man_index, district_list, region):
        driver.get(server)
        pf = driver.find_element_by_id("form1:table1:4:rowSelect1__input_sel")
        pf.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        man = driver.find_element_by_id(man_index)
        man.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        page2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        page2.click()
        fa = driver.find_element_by_id("form1:table1:38:rowSelect1__input_sel")
        fa.click()
        id = driver.find_element_by_id("form1:table1:58:rowSelect1__input_sel")
        id.click()
        page3 = driver.find_element_by_id("form1:table1:web1__pagerWeb__2")
        page3.click()
        im = driver.find_element_by_id("form1:table1:60:rowSelect1__input_sel")
        im.click()
        page4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        page4.click()
        npers = driver.find_element_by_id("form1:table1:93:rowSelect1__input_sel")
        npers.click()
        ot = driver.find_element_by_id("form1:table1:101:rowSelect1__input_sel")
        ot.click()
        page5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        page5.click()
        ra = driver.find_element_by_id("form1:table1:121:rowSelect1__input_sel")
        ra.click()
        rdat = driver.find_element_by_id("form1:table1:122:rowSelect1__input_sel")
        rdat.click()
        re = driver.find_element_by_id("form1:table1:123:rowSelect1__input_sel")
        re.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:5:menu1"))
        rayon_settings.select_by_visible_text("один из (через запятую)")
        rayons = driver.find_element_by_id("form1:table1:5:text7")
        rayons.send_keys(district_list)
        re_settings = Select(driver.find_element_by_id("form1:table1:7:menu1"))
        re_settings.select_by_visible_text("равен")
        re_keys = driver.find_element_by_id("form1:table1:7:text7")
        re_keys.send_keys(region)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def szvm_pe(self, server, man_index, pe_index, district, region):
        driver.get(server)
        pf = driver.find_element_by_id("form1:table1:4:rowSelect1__input_sel")
        pf.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        man = driver.find_element_by_id(man_index)
        man.click()
        pf_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        pf_page_2.click()
        pe = driver.find_element_by_id(pe_index)
        pe.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        man_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        man_page_2.click()
        man_id = driver.find_element_by_id("form1:table1:58:rowSelect1__input_sel")
        man_id.click()
        man_page_5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        man_page_5.click()
        man_ra = driver.find_element_by_id("form1:table1:121:rowSelect1__input_sel")
        man_ra.click()
        man_re = driver.find_element_by_id("form1:table1:123:rowSelect1__input_sel")
        man_re.click()
        man_page_6 = driver.find_element_by_id("form1:table1:web1__pagerWeb__5")
        man_page_6.click()
        pe_dat = driver.find_element_by_id("form1:table1:177:rowSelect1__input_sel")
        pe_dat.click()
        man_page_7 = driver.find_element_by_id("form1:table1:web1__pagerWeb__6")
        man_page_7.click()
        pe_datnp_tr = driver.find_element_by_id("form1:table1:186:rowSelect1__input_sel")
        pe_datnp_tr.click()
        pe_divp = driver.find_element_by_id("form1:table1:193:rowSelect1__input_sel")
        pe_divp.click()
        man_page_9 = driver.find_element_by_id("form1:table1:web1__pagerWeb__8")
        man_page_9.click()
        man_page_11 = driver.find_element_by_id("form1:table1:web1__pagerWeb__10")
        man_page_11.click()
        pe_otkaz = driver.find_element_by_id("form1:table1:309:rowSelect1__input_sel")
        pe_otkaz.click()
        pe_pgp = driver.find_element_by_id("form1:table1:315:rowSelect1__input_sel")
        pe_pgp.click()
        man_page_13 = driver.find_element_by_id("form1:table1:web1__pagerWeb__12")
        man_page_13.click()
        pe_ugo1 = driver.find_element_by_id("form1:table1:370:rowSelect1__input_sel")
        pe_ugo1.click()
        pe_ugo2 = driver.find_element_by_id("form1:table1:371:rowSelect1__input_sel")
        pe_ugo2.click()
        pe_utr = driver.find_element_by_id("form1:table1:372:rowSelect1__input_sel")
        pe_utr.click()
        pe_vgo1 = driver.find_element_by_id("form1:table1:382:rowSelect1__input_sel")
        pe_vgo1.click()
        pe_vgo2 = driver.find_element_by_id("form1:table1:383:rowSelect1__input_sel")
        pe_vgo2.click()
        pe_vtr = driver.find_element_by_id("form1:table1:384:rowSelect1__input_sel")
        pe_vtr.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:1:menu1"))
        rayon_settings.select_by_visible_text("равен")
        rayons = driver.find_element_by_id("form1:table1:1:text7")
        rayons.send_keys(district)
        re_settings = Select(driver.find_element_by_id("form1:table1:2:menu1"))
        re_settings.select_by_visible_text("равен")
        re_keys = driver.find_element_by_id("form1:table1:2:text7")
        re_keys.send_keys(region)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def szvm_popay(self, server, po_index, popay_index, district, region):
        driver.get(server)
        vpl = driver.find_element_by_id("form1:table1:9:rowSelect1__input_sel")
        vpl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        po = driver.find_element_by_id(po_index)
        po.click()
        vpl_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        vpl_page_2.click()
        popay = driver.find_element_by_id(popay_index)
        popay.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        page_2.click()
        po_id = driver.find_element_by_id("form1:table1:30:rowSelect1__input_sel")
        po_id.click()
        page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        page_4.click()
        po_ra = driver.find_element_by_id("form1:table1:95:rowSelect1__input_sel")
        po_ra.click()
        po_re = driver.find_element_by_id("form1:table1:97:rowSelect1__input_sel")
        po_re.click()
        popay_amount = driver.find_element_by_id("form1:table1:111:rowSelect1__input_sel")
        popay_amount.click()
        page_5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        page_5.click()
        popay_np = driver.find_element_by_id("form1:table1:129:rowSelect1__input_sel")
        popay_np.click()
        popay_razdel = driver.find_element_by_id("form1:table1:138:rowSelect1__input_sel")
        popay_razdel.click()
        popay_srokpo = driver.find_element_by_id("form1:table1:142:rowSelect1__input_sel")
        popay_srokpo.click()
        popay_sroks = driver.find_element_by_id("form1:table1:143:rowSelect1__input_sel")
        popay_sroks.click()
        popay_teqsrokpo = driver.find_element_by_id("form1:table1:145:rowSelect1__input_sel")
        popay_teqsrokpo.click()
        popay_teqsroks = driver.find_element_by_id("form1:table1:146:rowSelect1__input_sel")
        popay_teqsroks.click()
        page_6 = driver.find_element_by_id("form1:table1:web1__pagerWeb__5")
        page_6.click()
        popay_vidvpl = driver.find_element_by_id("form1:table1:151:rowSelect1__input_sel")
        popay_vidvpl.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:1:menu1"))
        rayon_settings.select_by_visible_text("равен")
        rayons = driver.find_element_by_id("form1:table1:1:text7")
        rayons.send_keys(district)
        re_settings = Select(driver.find_element_by_id("form1:table1:2:menu1"))
        re_settings.select_by_visible_text("равен")
        re_keys = driver.find_element_by_id("form1:table1:2:text7")
        re_keys.send_keys(region)
        vidvpl_settings = Select(driver.find_element_by_id("form1:table1:10:menu1"))
        vidvpl_settings.select_by_visible_text("равен")
        vidvpl_set = driver.find_element_by_id("form1:table1:10:text7")
        vidvpl_set.send_keys("10")
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def szvm_popen(self, server, po_index, popen_index, district, region):
        driver.get(server)
        vpl = driver.find_element_by_id("form1:table1:9:rowSelect1__input_sel")
        vpl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        po = driver.find_element_by_id(po_index)
        po.click()
        vpl_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        vpl_page_2.click()
        popen = driver.find_element_by_id(popen_index)
        popen.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        page_2.click()
        po_id = driver.find_element_by_id("form1:table1:30:rowSelect1__input_sel")
        po_id.click()
        page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        page_4.click()
        po_ra = driver.find_element_by_id("form1:table1:95:rowSelect1__input_sel")
        po_ra.click()
        po_re = driver.find_element_by_id("form1:table1:97:rowSelect1__input_sel")
        po_re.click()
        po_page_6 = driver.find_element_by_id("form1:table1:web1__pagerWeb__5")
        po_page_6.click()
        popen_np = driver.find_element_by_id("form1:table1:157:rowSelect1__input_sel")
        popen_np.click()
        popen_nvp = driver.find_element_by_id("form1:table1:160:rowSelect1__input_sel")
        popen_nvp.click()
        popen_pen_b = driver.find_element_by_id("form1:table1:168:rowSelect1__input_sel")
        popen_pen_b.click()
        popen_g1 = driver.find_element_by_id("form1:table1:170:rowSelect1__input_sel")
        popen_g1.click()
        popen_g2 = driver.find_element_by_id("form1:table1:171:rowSelect1__input_sel")
        popen_g2.click()
        popen_pen_s = driver.find_element_by_id("form1:table1:174:rowSelect1__input_sel")
        popen_pen_s.click()
        po_page_7 = driver.find_element_by_id("form1:table1:web1__pagerWeb__6")
        po_page_7.click()
        popen_rn = driver.find_element_by_id("form1:table1:192:rowSelect1__input_sel")
        popen_rn.click()
        popen_sposob = driver.find_element_by_id("form1:table1:195:rowSelect1__input_sel")
        popen_sposob.click()
        popen_srokpo = driver.find_element_by_id("form1:table1:196:rowSelect1__input_sel")
        popen_srokpo.click()
        popen_sroks = driver.find_element_by_id("form1:table1:197:rowSelect1__input_sel")
        popen_sroks.click()
        popen_statusr = driver.find_element_by_id("form1:table1:198:rowSelect1__input_sel")
        popen_statusr.click()
        popen_teqsrokpo = driver.find_element_by_id("form1:table1:203:rowSelect1")
        popen_teqsrokpo.click()
        popen_teqsroks = driver.find_element_by_id("form1:table1:204:rowSelect1__input_sel")
        popen_teqsroks.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:1:menu1"))
        rayon_settings.select_by_visible_text("равен")
        rayons = driver.find_element_by_id("form1:table1:1:text7")
        rayons.send_keys(district)
        re_settings = Select(driver.find_element_by_id("form1:table1:2:menu1"))
        re_settings.select_by_visible_text("равен")
        re_keys = driver.find_element_by_id("form1:table1:2:text7")
        re_keys.send_keys(region)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def szvm_wpr(self, server, district_list):
        driver.get(server)
        kl = driver.find_element_by_id("form1:table1:2:rowSelect1__input_sel")
        kl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        kl_page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        kl_page_4.click()
        wpr = driver.find_element_by_id("form1:table1:118:rowSelect1__input_sel")
        wpr.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        wpr_kod = driver.find_element_by_id("form1:table1:19:rowSelect1__input_sel")
        wpr_kod.click()
        wpr_name = driver.find_element_by_id("form1:table1:28:rowSelect1__input_sel")
        wpr_name.click()
        wpr_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        wpr_page_2.click()
        wpr_nus = driver.find_element_by_id("form1:table1:34:rowSelect1__input_sel")
        wpr_nus.click()
        wpr_ra = driver.find_element_by_id("form1:table1:38:rowSelect1__input_sel")
        wpr_ra.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:3:menu1"))
        rayon_settings.select_by_visible_text("один из (через запятую)")
        rayons = driver.find_element_by_id("form1:table1:3:text7")
        rayons.send_keys(district_list)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def szvm_man_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.szvm_man(server_url, man_index, ','.join(district_list), region)

    def szvm_pe_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.szvm_pe(server_url, man_index, pe_index, district, region)

    def szvm_popay_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.szvm_popay(server_url, po_index, popay_index, district, region)

    def szvm_popen_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.szvm_popen(server_url, po_index, popen_index, district, region)

    def szvm_wpr_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.szvm_wpr(server_url, ','.join(district_list))

    def prbz_man(self, server, man_index, district_list, region):
        driver.get(server)
        pf = driver.find_element_by_id("form1:table1:4:rowSelect1__input_sel")
        pf.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        man = driver.find_element_by_id(man_index)
        man.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        page2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        page2.click()
        fa = driver.find_element_by_id("form1:table1:38:rowSelect1__input_sel")
        fa.click()
        id = driver.find_element_by_id("form1:table1:57:rowSelect1__input_sel")
        id.click()
        im = driver.find_element_by_id("form1:table1:59:rowSelect1__input_sel")
        im.click()
        page4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        page4.click()
        npers = driver.find_element_by_id("form1:table1:92:rowSelect1__input_sel")
        npers.click()
        ot = driver.find_element_by_id("form1:table1:100:rowSelect1__input_sel")
        ot.click()
        page5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        page5.click()
        ra = driver.find_element_by_id("form1:table1:120:rowSelect1__input_sel")
        ra.click()
        rdat = driver.find_element_by_id("form1:table1:121:rowSelect1__input_sel")
        rdat.click()
        re = driver.find_element_by_id("form1:table1:122:rowSelect1__input_sel")
        re.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:5:menu1"))
        rayon_settings.select_by_visible_text("один из (через запятую)")
        rayons = driver.find_element_by_id("form1:table1:5:text7")
        rayons.send_keys(district_list)
        re_settings = Select(driver.find_element_by_id("form1:table1:7:menu1"))
        re_settings.select_by_visible_text("равен")
        re_keys = driver.find_element_by_id("form1:table1:7:text7")
        re_keys.send_keys(region)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def prbz_pe(self, server, man_index, pe_index, district, region):
        driver.get(server)
        pf = driver.find_element_by_id("form1:table1:4:rowSelect1__input_sel")
        pf.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        man = driver.find_element_by_id(man_index)
        man.click()
        pf_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        pf_page_2.click()
        pe = driver.find_element_by_id(pe_index)
        pe.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        man_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        man_page_2.click()
        man_id = driver.find_element_by_id("form1:table1:57:rowSelect1__input_sel")
        man_id.click()
        man_page_5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        man_page_5.click()
        man_ra = driver.find_element_by_id("form1:table1:120:rowSelect1__input_sel")
        man_ra.click()
        man_re = driver.find_element_by_id("form1:table1:122:rowSelect1__input_sel")
        man_re.click()
        man_page_6 = driver.find_element_by_id("form1:table1:web1__pagerWeb__5")
        man_page_6.click()
        pe_dat = driver.find_element_by_id("form1:table1:176:rowSelect1__input_sel")
        pe_dat.click()
        man_page_7 = driver.find_element_by_id("form1:table1:web1__pagerWeb__6")
        man_page_7.click()
        pe_datnp_tr = driver.find_element_by_id("form1:table1:185:rowSelect1__input_sel")
        pe_datnp_tr.click()
        pe_divp = driver.find_element_by_id("form1:table1:192:rowSelect1__input_sel")
        pe_divp.click()
        man_page_9 = driver.find_element_by_id("form1:table1:web1__pagerWeb__8")
        man_page_9.click()
        man_page_11 = driver.find_element_by_id("form1:table1:web1__pagerWeb__10")
        man_page_11.click()
        pe_otkaz = driver.find_element_by_id("form1:table1:308:rowSelect1__input_sel")
        pe_otkaz.click()
        pe_pgp = driver.find_element_by_id("form1:table1:314:rowSelect1__input_sel")
        pe_pgp.click()
        man_page_13 = driver.find_element_by_id("form1:table1:web1__pagerWeb__12")
        man_page_13.click()
        pe_ugo1 = driver.find_element_by_id("form1:table1:369:rowSelect1__input_sel")
        pe_ugo1.click()
        pe_ugo2 = driver.find_element_by_id("form1:table1:370:rowSelect1__input_sel")
        pe_ugo2.click()
        pe_utr = driver.find_element_by_id("form1:table1:371:rowSelect1__input_sel")
        pe_utr.click()
        pe_vgo1 = driver.find_element_by_id("form1:table1:381:rowSelect1__input_sel")
        pe_vgo1.click()
        pe_vgo2 = driver.find_element_by_id("form1:table1:382:rowSelect1__input_sel")
        pe_vgo2.click()
        pe_vtr = driver.find_element_by_id("form1:table1:383:rowSelect1__input_sel")
        pe_vtr.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:1:menu1"))
        rayon_settings.select_by_visible_text("равен")
        rayons = driver.find_element_by_id("form1:table1:1:text7")
        rayons.send_keys(district)
        re_settings = Select(driver.find_element_by_id("form1:table1:2:menu1"))
        re_settings.select_by_visible_text("равен")
        re_keys = driver.find_element_by_id("form1:table1:2:text7")
        re_keys.send_keys(region)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def prbz_popay(self, server, po_index, popay_index, district, region):
        driver.get(server)
        vpl = driver.find_element_by_id("form1:table1:9:rowSelect1__input_sel")
        vpl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        po = driver.find_element_by_id(po_index)
        po.click()
        vpl_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        vpl_page_2.click()
        popay = driver.find_element_by_id(popay_index)
        popay.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        po_id = driver.find_element_by_id("form1:table1:29:rowSelect1__input_sel")
        po_id.click()
        page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        page_4.click()
        po_ra = driver.find_element_by_id("form1:table1:94:rowSelect1__input_sel")
        po_ra.click()
        po_re = driver.find_element_by_id("form1:table1:96:rowSelect1__input_sel")
        po_re.click()
        popay_amount = driver.find_element_by_id("form1:table1:110:rowSelect1__input_sel")
        popay_amount.click()
        page_5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        page_5.click()
        popay_np = driver.find_element_by_id("form1:table1:128:rowSelect1__input_sel")
        popay_np.click()
        popay_nvp = driver.find_element_by_id("form1:table1:131:rowSelect1__input_sel")
        popay_nvp.click()
        popay_razdel = driver.find_element_by_id("form1:table1:137:rowSelect1__input_sel")
        popay_razdel.click()
        popay_rn = driver.find_element_by_id("form1:table1:138:rowSelect1__input_sel")
        popay_rn.click()
        popay_sposob = driver.find_element_by_id("form1:table1:140:rowSelect1__input_sel")
        popay_sposob.click()
        popay_srokpo = driver.find_element_by_id("form1:table1:141:rowSelect1__input_sel")
        popay_srokpo.click()
        popay_sroks = driver.find_element_by_id("form1:table1:142:rowSelect1__input_sel")
        popay_sroks.click()
        popay_teqsrokpo = driver.find_element_by_id("form1:table1:144:rowSelect1__input_sel")
        popay_teqsrokpo.click()
        popay_teqsroks = driver.find_element_by_id("form1:table1:145:rowSelect1__input_sel")
        popay_teqsroks.click()
        page_6 = driver.find_element_by_id("form1:table1:web1__pagerWeb__5")
        page_6.click()
        popay_vidvpl = driver.find_element_by_id("form1:table1:150:rowSelect1__input_sel")
        popay_vidvpl.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:1:menu1"))
        rayon_settings.select_by_visible_text("равен")
        rayons = driver.find_element_by_id("form1:table1:1:text7")
        rayons.send_keys(district)
        re_settings = Select(driver.find_element_by_id("form1:table1:2:menu1"))
        re_settings.select_by_visible_text("равен")
        re_keys = driver.find_element_by_id("form1:table1:2:text7")
        re_keys.send_keys(region)
        vidvpl_settings = Select(driver.find_element_by_id("form1:table1:13:menu1"))
        vidvpl_settings.select_by_visible_text("один из (через запятую)")
        vidvpl_set = driver.find_element_by_id("form1:table1:13:text7")
        vidvpl_set.send_keys("10,12")
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def prbz_popen(self, server, po_index, popen_index, district, region):
        driver.get(server)
        vpl = driver.find_element_by_id("form1:table1:9:rowSelect1__input_sel")
        vpl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        po = driver.find_element_by_id(po_index)
        po.click()
        vpl_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        vpl_page_2.click()
        popen = driver.find_element_by_id(popen_index)
        popen.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        po_id = driver.find_element_by_id("form1:table1:29:rowSelect1__input_sel")
        po_id.click()
        po_page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        po_page_4.click()
        po_ra = driver.find_element_by_id("form1:table1:94:rowSelect1__input_sel")
        po_ra.click()
        po_re = driver.find_element_by_id("form1:table1:96:rowSelect1__input_sel")
        po_re.click()
        po_page_6 = driver.find_element_by_id("form1:table1:web1__pagerWeb__5")
        po_page_6.click()
        popen_np = driver.find_element_by_id("form1:table1:156:rowSelect1__input_sel")
        popen_np.click()
        popen_nvp = driver.find_element_by_id("form1:table1:159:rowSelect1__input_sel")
        popen_nvp.click()
        popen_pen_b = driver.find_element_by_id("form1:table1:167:rowSelect1__input_sel")
        popen_pen_b.click()
        popen_g1 = driver.find_element_by_id("form1:table1:169:rowSelect1__input_sel")
        popen_g1.click()
        popen_g2 = driver.find_element_by_id("form1:table1:170:rowSelect1__input_sel")
        popen_g2.click()
        popen_pen_s = driver.find_element_by_id("form1:table1:173:rowSelect1__input_sel")
        popen_pen_s.click()
        po_page_7 = driver.find_element_by_id("form1:table1:web1__pagerWeb__6")
        po_page_7.click()
        popen_rn = driver.find_element_by_id("form1:table1:191:rowSelect1__input_sel")
        popen_rn.click()
        popen_sposob = driver.find_element_by_id("form1:table1:194:rowSelect1__input_sel")
        popen_sposob.click()
        popen_srokpo = driver.find_element_by_id("form1:table1:195:rowSelect1__input_sel")
        popen_srokpo.click()
        popen_sroks = driver.find_element_by_id("form1:table1:196:rowSelect1__input_sel")
        popen_sroks.click()
        popen_statusr = driver.find_element_by_id("form1:table1:197:rowSelect1__input_sel")
        popen_statusr.click()
        popen_teqsrokpo = driver.find_element_by_id("form1:table1:202:rowSelect1")
        popen_teqsrokpo.click()
        popen_teqsroks = driver.find_element_by_id("form1:table1:203:rowSelect1__input_sel")
        popen_teqsroks.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:1:menu1"))
        rayon_settings.select_by_visible_text("равен")
        rayons = driver.find_element_by_id("form1:table1:1:text7")
        rayons.send_keys(district)
        re_settings = Select(driver.find_element_by_id("form1:table1:2:menu1"))
        re_settings.select_by_visible_text("равен")
        re_keys = driver.find_element_by_id("form1:table1:2:text7")
        re_keys.send_keys(region)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def prbz_wpr(self, server, district_list):
        driver.get(server)
        kl = driver.find_element_by_id("form1:table1:2:rowSelect1__input_sel")
        kl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        kl_page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        kl_page_4.click()
        wpr = driver.find_element_by_id("form1:table1:118:rowSelect1__input_sel")
        wpr.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        wpr_kod = driver.find_element_by_id("form1:table1:19:rowSelect1__input_sel")
        wpr_kod.click()
        wpr_name = driver.find_element_by_id("form1:table1:28:rowSelect1__input_sel")
        wpr_name.click()
        wpr_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        wpr_page_2.click()
        wpr_nus = driver.find_element_by_id("form1:table1:34:rowSelect1__input_sel")
        wpr_nus.click()
        wpr_ra = driver.find_element_by_id("form1:table1:38:rowSelect1__input_sel")
        wpr_ra.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:3:menu1"))
        rayon_settings.select_by_visible_text("один из (через запятую)")
        rayons = driver.find_element_by_id("form1:table1:3:text7")
        rayons.send_keys(district_list)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def prbz_man_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.prbz_man(server_url, man_index, ','.join(district_list), region)

    def prbz_pe_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.prbz_pe(server_url, man_index, pe_index, district, region)

    def prbz_popay_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.prbz_popay(server_url, po_index, popay_index, district, region)

    def prbz_popen_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.prbz_popen(server_url, po_index, popen_index, district, region)

    def prbz_wpr_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.prbz_wpr(server_url, ','.join(district_list))

    def edv_man(self, server, man_index, district_list):
        driver.get(server)
        pf = driver.find_element_by_id("form1:table1:4:rowSelect1__input_sel")
        pf.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        man = driver.find_element_by_id(man_index)
        man.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        adr_index = driver.find_element_by_id("form1:table1:1:rowSelect1__input_sel")
        adr_index.click()
        page2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        page2.click()
        fa = driver.find_element_by_id("form1:table1:38:rowSelect1__input_sel")
        fa.click()
        id = driver.find_element_by_id("form1:table1:57:rowSelect1__input_sel")
        id.click()
        im = driver.find_element_by_id("form1:table1:59:rowSelect1__input_sel")
        im.click()
        page4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        page4.click()
        npers = driver.find_element_by_id("form1:table1:92:rowSelect1__input_sel")
        npers.click()
        ot = driver.find_element_by_id("form1:table1:100:rowSelect1__input_sel")
        ot.click()
        page5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        page5.click()
        ra = driver.find_element_by_id("form1:table1:120:rowSelect1__input_sel")
        ra.click()
        rdat = driver.find_element_by_id("form1:table1:121:rowSelect1__input_sel")
        rdat.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:6:menu1"))
        rayon_settings.select_by_visible_text("один из (через запятую)")
        rayons = driver.find_element_by_id("form1:table1:6:text7")
        rayons.send_keys(district_list)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def edv_podmo(self, server, po_index, podmo_index, district):
        driver.get(server)
        vpl = driver.find_element_by_id("form1:table1:9:rowSelect1__input_sel")
        vpl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        po = driver.find_element_by_id(po_index)
        po.click()
        podmo = driver.find_element_by_id(podmo_index)
        podmo.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        po_id = driver.find_element_by_id("form1:table1:29:rowSelect1__input_sel")
        po_id.click()
        po_page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        po_page_4.click()
        po_ra = driver.find_element_by_id("form1:table1:94:rowSelect1__input_sel")
        po_ra.click()
        podmo_dpw = driver.find_element_by_id("form1:table1:117:rowSelect1__input_sel")
        podmo_dpw.click()
        podmo_page_5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        podmo_page_5.click()
        podmo_kat_dmo = driver.find_element_by_id("form1:table1:122:rowSelect1__input_sel")
        podmo_kat_dmo.click()
        podmo_np = driver.find_element_by_id("form1:table1:124:rowSelect1__input_sel")
        podmo_np.click()
        podmo_nvp = driver.find_element_by_id("form1:table1:127:rowSelect1__input_sel")
        podmo_nvp.click()
        podmo_pw = driver.find_element_by_id("form1:table1:132:rowSelect1__input_sel")
        podmo_pw.click()
        podmo_s_dmo = driver.find_element_by_id("form1:table1:134:rowSelect1__input_sel")
        podmo_s_dmo.click()
        podmo_sposob = driver.find_element_by_id("form1:table1:136:rowSelect1__input_sel")
        podmo_sposob.click()
        podmo_spokpo = driver.find_element_by_id("form1:table1:137:rowSelect1__input_sel")
        podmo_spokpo.click()
        podmo_spoks = driver.find_element_by_id("form1:table1:138:rowSelect1__input_sel")
        podmo_spoks.click()
        podmo_teqsrokpo = driver.find_element_by_id("form1:table1:140:rowSelect1__input_sel")
        podmo_teqsrokpo.click()
        podmo_teqsroks = driver.find_element_by_id("form1:table1:141:rowSelect1__input_sel")
        podmo_teqsroks.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:1:menu1"))
        rayon_settings.select_by_visible_text("равен")
        rayons = driver.find_element_by_id("form1:table1:1:text7")
        rayons.send_keys(district)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def edv_popen(self, server, po_index, popen_index, district):
        driver.get(server)
        vpl = driver.find_element_by_id("form1:table1:9:rowSelect1__input_sel")
        vpl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        po = driver.find_element_by_id(po_index)
        po.click()
        vpl_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        vpl_page_2.click()
        popen = driver.find_element_by_id(popen_index)
        popen.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        po_id = driver.find_element_by_id("form1:table1:29:rowSelect1__input_sel")
        po_id.click()
        po_page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        po_page_4.click()
        po_ra = driver.find_element_by_id("form1:table1:94:rowSelect1__input_sel")
        po_ra.click()
        po_page_5 = driver.find_element_by_id("form1:table1:web1__pagerWeb__4")
        po_page_5.click()
        popen_dpw = driver.find_element_by_id("form1:table1:134:rowSelect1__input_sel")
        popen_dpw.click()
        po_page_6 = driver.find_element_by_id("form1:table1:web1__pagerWeb__5")
        po_page_6.click()
        popen_np = driver.find_element_by_id("form1:table1:156:rowSelect1__input_sel")
        popen_np.click()
        popen_nvp = driver.find_element_by_id("form1:table1:159:rowSelect1__input_sel")
        popen_nvp.click()
        popen_pen_b = driver.find_element_by_id("form1:table1:167:rowSelect1__input_sel")
        popen_pen_b.click()
        popen_g1 = driver.find_element_by_id("form1:table1:169:rowSelect1__input_sel")
        popen_g1.click()
        popen_g2 = driver.find_element_by_id("form1:table1:170:rowSelect1__input_sel")
        popen_g2.click()
        popen_pen_s = driver.find_element_by_id("form1:table1:173:rowSelect1__input_sel")
        popen_pen_s.click()
        po_page_7 = driver.find_element_by_id("form1:table1:web1__pagerWeb__6")
        po_page_7.click()
        popen_pw = driver.find_element_by_id("form1:table1:190:rowSelect1__input_sel")
        popen_pw.click()
        popen_sposob = driver.find_element_by_id("form1:table1:194:rowSelect1__input_sel")
        popen_sposob.click()
        popen_srokpo = driver.find_element_by_id("form1:table1:195:rowSelect1__input_sel")
        popen_srokpo.click()
        popen_sroks = driver.find_element_by_id("form1:table1:196:rowSelect1__input_sel")
        popen_sroks.click()
        popen_teqsrokpo = driver.find_element_by_id("form1:table1:202:rowSelect1")
        popen_teqsrokpo.click()
        popen_teqsroks = driver.find_element_by_id("form1:table1:203:rowSelect1__input_sel")
        popen_teqsroks.click()
        popen_u_go1 = driver.find_element_by_id("form1:table1:205:rowSelect1__input_sel")
        popen_u_go1.click()
        popen_u_go2 = driver.find_element_by_id("form1:table1:206:rowSelect1__input_sel")
        popen_u_go2.click()
        popen_u_tr = driver.find_element_by_id("form1:table1:207:rowSelect1__input_sel")
        popen_u_tr.click()
        popen_v_go1 = driver.find_element_by_id("form1:table1:209:rowSelect1__input_sel")
        popen_v_go1.click()
        po_page_8 = driver.find_element_by_id("form1:table1:web1__pagerWeb__7")
        po_page_8.click()
        popen_v_go2 = driver.find_element_by_id("form1:table1:210:rowSelect1__input_sel")
        popen_v_go2.click()
        popen_v_tr = driver.find_element_by_id("form1:table1:211:rowSelect1__input_sel")
        popen_v_tr.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:1:menu1"))
        rayon_settings.select_by_visible_text("равен")
        rayons = driver.find_element_by_id("form1:table1:1:text7")
        rayons.send_keys(district)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def edv_wpr(self, server, district_list):
        driver.get(server)
        kl = driver.find_element_by_id("form1:table1:2:rowSelect1__input_sel")
        kl.click()
        button_1 = driver.find_element_by_id("form1:button2")
        button_1.click()
        kl_page_4 = driver.find_element_by_id("form1:table1:web1__pagerWeb__3")
        kl_page_4.click()
        wpr = driver.find_element_by_id("form1:table1:118:rowSelect1__input_sel")
        wpr.click()
        button_2 = driver.find_element_by_id("form1:button2")
        button_2.click()
        wpr_kod = driver.find_element_by_id("form1:table1:19:rowSelect1__input_sel")
        wpr_kod.click()
        wpr_name = driver.find_element_by_id("form1:table1:28:rowSelect1__input_sel")
        wpr_name.click()
        wpr_page_2 = driver.find_element_by_id("form1:table1:web1__pagerWeb__1")
        wpr_page_2.click()
        wpr_nus = driver.find_element_by_id("form1:table1:34:rowSelect1__input_sel")
        wpr_nus.click()
        wpr_ra = driver.find_element_by_id("form1:table1:38:rowSelect1__input_sel")
        wpr_ra.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        rayon_settings = Select(driver.find_element_by_id("form1:table1:3:menu1"))
        rayon_settings.select_by_visible_text("один из (через запятую)")
        rayons = driver.find_element_by_id("form1:table1:3:text7")
        rayons.send_keys(district_list)
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()

    def edv_man_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.edv_man(server_url, man_index, ','.join(district_list))

    def edv_podmo_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.edv_podmo(server_url, po_index, podmo_index, district)

    def edv_popen_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.edv_popen(server_url, po_index, popen_index, district)

    def edv_wpr_run(self):
        self.properties()
        if status == 0:
            self.browser(download_folder)
            self.login(login_url)
            self.edv_wpr(server_url, ','.join(district_list))

    def deduction_main(self, server, man_index, is_index, popen_index, gu_ra_list):
        driver.get(server)
        pf = driver.find_element_by_id("form1:table1:4:rowSelect1__input_sel")
        pf.click()
        vpl = driver.find_element_by_id('form1:table1:9:rowSelect1__input_sel')
        vpl.click()
        button_1 = driver.find_element_by_id('form1:button2')
        button_1.click()
        table_man = driver.find_element_by_id(man_index)
        table_man.click()
        table_page_4 = driver.find_element_by_id('form1:table1:web1__pagerWeb__3')
        table_page_4.click()
        table_is = driver.find_element_by_id(is_index)
        table_is.click()
        table_page_5 = driver.find_element_by_id('form1:table1:web1__pagerWeb__4')
        table_page_5.click()
        table_popen = driver.find_element_by_id(popen_index)
        table_popen.click()
        button_2 = driver.find_element_by_id('form1:button2')
        button_2.click()
        column_page_2 = driver.find_element_by_id('form1:table1:web1__pagerWeb__1')
        column_page_2.click()
        man_fa = driver.find_element_by_id('form1:table1:38:rowSelect1__input_sel')
        man_fa.click()
        column_page_3 = driver.find_element_by_id('form1:table1:web1__pagerWeb__2')
        column_page_3.click()
        man_im = driver.find_element_by_id('form1:table1:60:rowSelect1__input_sel')
        man_im.click()
        column_page_4 = driver.find_element_by_id('form1:table1:web1__pagerWeb__3')
        column_page_4.click()
        man_npers = driver.find_element_by_id('form1:table1:93:rowSelect1__input_sel')
        man_npers.click()
        man_ot = driver.find_element_by_id('form1:table1:101:rowSelect1__input_sel')
        man_ot.click()
        column_page_5 = driver.find_element_by_id('form1:table1:web1__pagerWeb__4')
        column_page_5.click()
        man_pw = driver.find_element_by_id('form1:table1:120:rowSelect1__input_sel')
        man_pw.click()
        man_ra = driver.find_element_by_id('form1:table1:121:rowSelect1__input_sel')
        man_ra.click()
        column_page_7 = driver.find_element_by_id('form1:table1:web1__pagerWeb__6')
        column_page_7.click()
        is_doc = driver.find_element_by_id('form1:table1:180:rowSelect1__input_sel')
        is_doc.click()
        is_docdv = driver.find_element_by_id('form1:table1:181:rowSelect1__input_sel')
        is_docdv.click()
        is_docnd = driver.find_element_by_id('form1:table1:182:rowSelect1__input_sel')
        is_docnd.click()
        is_mpu = driver.find_element_by_id('form1:table1:197:rowSelect1__input_sel')
        is_mpu.click()
        is_msu = driver.find_element_by_id('form1:table1:198:rowSelect1__input_sel')
        is_msu.click()
        column_page_8 = driver.find_element_by_id('form1:table1:web1__pagerWeb__7')
        column_page_8.click()
        is_reason_closed = driver.find_element_by_id('form1:table1:210:rowSelect1__input_sel')
        is_reason_closed.click()
        is_srokpo = driver.find_element_by_id('form1:table1:217:rowSelect1__input_sel')
        is_srokpo.click()
        is_sroks = driver.find_element_by_id('form1:table1:218:rowSelect1__input_sel')
        is_sroks.click()
        is_vv = driver.find_element_by_id('form1:table1:228:rowSelect1__input_sel')
        is_vv.click()
        column_page_10 = driver.find_element_by_id('form1:table1:web1__pagerWeb__9')
        column_page_10.click()
        column_page_12 = driver.find_element_by_id('form1:table1:web1__pagerWeb__11')
        column_page_12.click()
        popen_v_go1 = driver.find_element_by_id("form1:table1:330:rowSelect1__input_sel")
        popen_v_go1.click()
        popen_v_tr = driver.find_element_by_id("form1:table1:332:rowSelect1__input_sel")
        popen_v_tr.click()
        button_3 = driver.find_element_by_id("form1:button3")
        button_3.click()
        pw_settings = Select(driver.find_element_by_id("form1:table1:4:menu1"))
        pw_settings.select_by_visible_text("ни один из (через запятую)")
        pw_settings_in = driver.find_element_by_id('form1:table1:4:text7')
        pw_settings_in.send_keys('1,2,50')
        ra_settings = Select(driver.find_element_by_id('form1:table1:5:menu1'))
        ra_settings.select_by_visible_text('один из (через запятую)')
        ra_settings_in = driver.find_element_by_id('form1:table1:5:text7')
        ra_settings_in.send_keys(gu_ra_list)
        reason_closed_settings = Select(driver.find_element_by_id('form1:table1:11:menu1'))
        reason_closed_settings.select_by_visible_text('равен')
        reason_closed_settings_in = driver.find_element_by_id('form1:table1:11:text7')
        reason_closed_settings_in.send_keys('0')
        button_4 = driver.find_element_by_id("form1:button1")
        button_4.click()
        save = WebDriverWait(driver, 500).until(
            EC.presence_of_element_located((By.XPATH, "//input[@value='Сохранить в файл csv']")))
        save.click()

    def deduction_download(self):
        if os.path.isfile('D:/ADV8D/results.csv.part') == False:
            sleep(10)
        while os.path.isfile('D:/ADV8D/results.csv.part'):
            sleep(5)
        os.replace("D:/ADV8D/results.csv", "D:/ADV8D/temp.csv")

    def deduction_add_header_first(self):
        header = ['fa', 'im', 'npers', 'ot', 'pw', 'ra', 'doc', 'docdv', 'docnd', 'mpu', 'msu', 'reason_closed', 'srokpo',
                  'sroks', 'vv', 'v_go1', 'v_tr']

        with open('D:/ADV8D/do.csv', 'w', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)

        with open('D:/ADV8D/temp.csv', 'r', newline='') as f1:
            original = f1.read()

        with open('D:/ADV8D/do.csv', 'a', newline='') as f2:
            f2.write(original)

    def deduction_sql(self):
        con = sqlite3.connect('D:/ADV8D/tmp.db')
        cur = con.cursor()
        cur.execute('CREATE TABLE t (fa, im, npers, ot, pw, ra, doc, docdv, docnd, mpu, msu, reason_closed, srokpo, sroks, vv, v_go1, v_tr);')
        with open('D:/ADV8D/do.csv', 'r') as f:
            dr = csv.DictReader(f, delimiter=';')
            to_db = [(i['fa'], i['im'], i['npers'], i['ot'], i['pw'], i['ra'], i['doc'], i['docdv'], i['docnd'], i['mpu'],
                      i['msu'], i['reason_closed'], i['srokpo'], i['sroks'], i['vv'], i['v_go1'], i['v_tr']) for i in dr]
        cur.executemany("INSERT INTO t (fa, im, npers, ot, pw, ra, doc, docdv, docnd, mpu, msu, reason_closed, srokpo, sroks, vv, v_go1, v_tr) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", to_db)

        with open('D:/ADV8D/po.csv', 'w', newline='') as f:
            for row in cur.execute("SELECT ra, npers, fa, im, ot, doc, docdv, docnd, vv, mpu, msu, sroks, srokpo, v_go1, v_tr FROM t ORDER BY ra"):
                w = csv.writer(f, delimiter=';')
                w.writerow(row)

        con.commit()
        con.close()

    def deduction_add_header_second(self, region_prefix, gu):
        header = ['Район', 'СНИСЛ', 'Фамилия', 'Имя', 'Отчество', 'Номер исполнительного документа', 'Дата выдачи исполнительного документа',
                  'Название исполнительного документа', 'Вид взыскания', 'Ежемесячный процент удержания',
                  'Ежемесячная сумма удержания', 'Дата начала выплат', 'Дата окончания удержания', 'Вид госпенсии-1', 'Вид трудовой пенсии']

        with open('D:/ADV8D/uderzhanie_' + region_prefix + '_' + gu + '.csv', 'w', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(header)

        with open('D:/ADV8D/po.csv', 'r', newline='') as f1:
            original = f1.read()

        with open('D:/ADV8D/uderzhanie_' + region_prefix + '_' + gu + '.csv', 'a', newline='') as f2:
            f2.write(original)

    def deduction_to_excel(self, region_prefix, gu):
            workbook = Workbook('D:/ADV8/uderzhanie_' + region_prefix + '_' + gu + '.xlsx')
            worksheet = workbook.add_worksheet()
            with open('D:/ADV8D/uderzhanie_' + region_prefix + '_' + gu + '.csv', 'r') as f:
                reader = csv.reader(f, delimiter=';')
                for r, row in enumerate(reader):
                    for c, col in enumerate(row):
                        worksheet.write(r, c, col)
            workbook.close()

    def deduction_cleaner(self):
        folder = 'D:/ADV8D'
        for file in os.listdir(folder):
            file_path = os.path.join(folder, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                elif os.path.isfile(file_path):
                    shutil.rmtree(file_path)
            except:
                pass

    def deduction_run_m(self):
        region_prefix = 'm'
        login_url = 'http://10.87.0.178/MainWAR/login.html'
        server_url = 'http://10.87.0.178/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
        man_index = "form1:table1:25:rowSelect1__input_sel"
        is_index = 'form1:table1:109:rowSelect1__input_sel'
        popen_index = "form1:table1:136:rowSelect1__input_sel"
        gu1_m = ('001')
        gu2_m = ('801', '802', '803', '805', '807', '812', '813', '815')
        gu3_m = ('501', '502', '504', '505', '506', '507', '508', '509', '510', '511', '512', '513')
        gu4_m = ('701', '702', '703', '705', '707', '708', '709', '710', '711', '712', '713', '714', '715', '716', '717')
        gu5_m = ('201', '202', '203', '205', '206', '207', '208', '209', '210', '211', '212', '213', '214', '215', '216')

        self.properties()
        if gu == '1':
            gu_number = 'gu1'
            gu_ra_list = gu1_m
        elif gu == '2':
            gu_number = 'gu2'
            gu_ra_list = gu2_m
        elif gu == '3':
            gu_number = 'gu3'
            gu_ra_list = gu3_m
        elif gu == '4':
            gu_number = 'gu4'
            gu_ra_list = gu4_m
        elif gu == '5':
            gu_number = 'gu5'
            gu_ra_list = gu5_m

        self.browser(download_folder)
        self.login(login_url)
        self.deduction_main(server_url, man_index, is_index, popen_index, ','.join(gu_ra_list))
        self.deduction_download()
        self.deduction_add_header_first()
        self.deduction_sql()
        self.deduction_add_header_second(region_prefix, gu_number)
        self.deduction_to_excel(region_prefix, gu_number)
        self.deduction_cleaner()
        driver.close()
        os.system('taskkill /f /im firefox.exe')

    def deduction_run_mo(self):
        region_prefix = 'mo'
        login_url = 'http://10.87.0.80/MainWAR/login.html'
        server_url = 'http://10.87.0.80/MainWAR/faces/menu/_rlvid.jsp?_rap=pc_MainMenu.doLink111QAction&_rvip=/menu/MainMenu.jsp'
        man_index = "form1:table1:23:rowSelect1__input_sel"
        is_index = "form1:table1:107:rowSelect1__input_sel"
        popen_index = "form1:table1:133:rowSelect1__input_sel"
        gu1_mo = ('20', '45')
        gu2_mo = ('32', '42', '132')
        gu3_mo = ('4',  '11', '27', '227', '327', '427', '527')
        gu4_mo = ('7', '36', '56', '156')
        gu5_mo = ('6', '23', '47', '50', '150')
        upfr = ('5', '8', '9', '16', '17', '21', '30', '31', '37', '38', '39', '40', '43', '55', '116', '140')

        self.properties()
        if gu == '1':
            gu_number = 'gu1'
            gu_ra_list = gu1_mo
        elif gu == '2':
            gu_number = 'gu2'
            gu_ra_list = gu2_mo
        elif gu == '3':
            gu_number = 'gu3'
            gu_ra_list = gu3_mo
        elif gu == '4':
            gu_number = 'gu4'
            gu_ra_list = gu4_mo
        elif gu == '5':
            gu_number = 'gu5'
            gu_ra_list = gu5_mo
        elif gu == 'УПФР':
            gu_number = 'upfr'
            gu_ra_list = upfr

        self.browser(download_folder)
        self.login(login_url)
        self.deduction_main(server_url, man_index, is_index, popen_index, ','.join(gu_ra_list))
        self.deduction_download()
        self.deduction_add_header_first()
        self.deduction_sql()
        self.deduction_add_header_second(region_prefix, gu_number)
        self.deduction_to_excel(region_prefix, gu_number)
        self.deduction_cleaner()
        driver.close()
        os.system('taskkill /f /im firefox.exe')


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()
