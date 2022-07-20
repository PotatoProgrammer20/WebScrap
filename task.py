"""Template robot with Python."""

# +
import os
import re
import time

from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF

from excel import WorkWithExcel


class ParseAgencies:

    agencyNum = 26

    def __init__(self):
        self.browserLibrary = Selenium()
        self.excel = WorkWithExcel()
        self.agencyInfo = {}
        self.agencyAll = []
        self.uii = []
        self.uiiURL = []
        self.investmentTitle = []
        self.browserLibrary.set_download_directory(os.path.join(os.getcwd(), 'output'))
        self.pdf = PDF()
        self.investmentName = []
        self.uniqueInvestmentIdentifier = []

    def openWebsite(self, url):
        self.browserLibrary.open_available_browser(url)

    def clickElement(self, elem):
        self.browserLibrary.wait_until_element_is_visible(elem)
        self.browserLibrary.click_element(elem)

    def getAgencies(self, xpath):
        self.browserLibrary.wait_until_element_is_visible(
            "//div[@id='agency-tiles-widget']//div[@class='col-sm-4 text-center noUnderline']", 15)
        self.agencyAll = self.browserLibrary.find_elements(xpath)
        #print(xpath)
        agencyList = [agency.text.split("\n") for agency in self.agencyAll]
        agencyName = [value[0] for value in agencyList]
        total = [value[2] for value in agencyList]
        self.agencyInfo = {'Agency': agencyName, 'Total spending': total}

    def getAgencyPage(self, agencyNum):
        agency = self.agencyAll[agencyNum - 5]                  #this gives the position of the agency we selected for indiviual invest
        agencyURL = self.browserLibrary.find_element(agency).find_element_by_tag_name('a').get_attribute('href')
        self.browserLibrary.go_to(agencyURL)

    def getIndInvest(self, path, sheet):
        data = []

        self.excel.createSheet(path, sheet)

        self.browserLibrary.wait_until_element_is_visible("//table[@id='investments-table-object']", 15)
        self.browserLibrary.wait_until_page_contains_element(
            '//*[@id="investments-table-object_length"]/label/select', 15
        )
        self.browserLibrary.find_element('//*[@id="investments-table-object_length"]/label/select').click()
        self.browserLibrary.find_element('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
        self.browserLibrary.wait_until_element_is_visible("//a[@class='paginate_button next disabled']", 20)

        table = self.browserLibrary.find_element("//table[@id='investments-table-object']")
        rows = table.find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')

        for row in rows:
            value = row.find_elements_by_tag_name('td')

            try:
                self.uiiURL.append(value[0].find_element_by_tag_name('a').get_attribute('href'))
                self.investmentTitle.append(value[2].text)
                self.uii.append(value[0].text)
            except Exception:
                pass

            values = [val.text for val in value]
            data.append(values)

        self.excel.appendRow(data, path, sheet)

    def downloadPDF(self):
        for url in self.uiiURL:
            self.browserLibrary.go_to(url)
            self.browserLibrary.wait_until_page_contains_element('//div[@id="business-case-pdf"]')
            self.browserLibrary.find_element('//div[@id="business-case-pdf"]').click()
            self.browserLibrary.wait_until_element_is_visible("//div[@id='business-case-pdf']//span", 15)
            self.browserLibrary.wait_until_element_is_not_visible("//div[@id='business-case-pdf']//span", 15)
        time.sleep(3)

    def getPDF(self):
        for pdf_name in self.uii:
            text = self.pdf.get_text_from_pdf(f'output/{pdf_name}.pdf')
            self.investmentName.append(re.search(
                r'Name of this Investment:(.*)2\.', text[1]).group(1).strip())
            self.uniqueInvestmentIdentifier .append(re.search(
                r'Unique Investment Identifier \(UII\):(.*)Section B', text[1]).group(1).strip())

    def comparePDF(self):
        for i in range(len(self.investmentTitle )):
            if self.investmentName[i] == self.investmentTitle [i]:
                print(f'PDF name: {self.investmentName[i]} matches with its Title name')
            else:
                print(f'PDF name: {self.investmentName[i]} is different from its Title name')
            if self.uniqueInvestmentIdentifier [i] == self.uii[i]:
                print(f'PDF UII:{self.uniqueInvestmentIdentifier [i]} matches with its Title name')
            else:
                print(f'PDF UII:{self.uniqueInvestmentIdentifier [i]} is different from its Title name')

if __name__ == "__main__":
    parse = ParseAgencies()

    try:
        parse.openWebsite("https://itdashboard.gov/")
        parse.clickElement("node-23")
        parse.getAgencies('//div[@id="agency-tiles-widget"]//div[@class="col-sm-4 text-center noUnderline"]')
        parse.excel.createFile('output/agencies.xlsx')
        parse.excel.renameSheet('output/agencies.xlsx', 'Sheet', 'Agencies')
        parse.excel.appendRow(parse.agencyInfo, 'output/agencies.xlsx', 'Agencies')
        parse.getAgencyPage(parse.agencyNum)
        parse.getIndInvest('output/agencies.xlsx', 'Individual Investments')
        parse.downloadPDF()
        parse.getPDF()
        parse.comparePDF()

    finally:
        parse.browserLibrary.close_all_browsers()
# -


