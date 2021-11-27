import time
import os
from datetime import timedelta
from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from RPA.Excel.Files import Files


def wait_for_downloads_file(filename: str):
    print(f'Waiting for {filename}', end='')
    while filename not in os.listdir('output'):
        time.sleep(1)
        print(".", end="")
    print('\nFile downloaded!')


def rpa_parse_pdf_data(filename: str) -> dict:
    pdf = PDF()
    start = 'Name of this Investment: '
    sep = '2. Unique Investment Identifier (UII): '
    end = 'Section B:'
    page_text = pdf.get_text_from_pdf(filename, 1)[1].replace("\n", "")
    investment_name_index = page_text.find(start) + len(start)
    uii_index_start = page_text.find(sep)
    uii_index_end = uii_index_start + len(sep)
    end_index = page_text.find(end)
    investment = page_text[investment_name_index: uii_index_start]
    uii = page_text[uii_index_end:end_index]
    return {'uii': uii, 'investment': investment}


def get_departments(webdriver: Selenium):
    webdriver.open_headless_chrome_browser('https://itdashboard.gov/')
    # webdriver.open_chrome_browser('https://itdashboard.gov/')
    webdriver.click_element("//a[@href='#home-dive-in']")
    webdriver.wait_until_element_is_visible('//div[@id="node-23"]//div[@id="agency-tiles-container"]')
    departments = webdriver.find_elements('//div[@id="agency-tiles-widget"]//span[@class="h4 w200"]')
    budgets = webdriver.find_elements('//div[@id="agency-tiles-widget"]//span[@class=" h1 w900"]')
    return {"departments": departments,
            "budgets": budgets}


def write_budgets(workbook: Files, departments, budgets):
    workbook.rename_worksheet('Sheet', 'Departments')
    for row_num, data in enumerate(departments):
        workbook.set_cell_value(row_num + 1, 1, data.text)
    for row_num, data in enumerate(budgets):
        workbook.set_cell_value(row_num + 1, 2, data.text)
    print('Departments budgets: Done!')


def get_agencies_table(webdriver: Selenium, dep_to_scrap: str) -> list:
    webdriver.wait_until_element_is_visible(
        f"//div[@id='agency-tiles-widget']//span[contains(text(),'{dep_to_scrap}')]/..", timedelta(seconds=30))
    webdriver.click_element(f"//div[@id='agency-tiles-widget']//span[contains(text(),'{dep_to_scrap}')]/..")
    webdriver.wait_until_element_is_visible("//div[@id='investments-table-widget']", timedelta(seconds=30))
    webdriver.select_from_list_by_value(
        "//div[@id='investments-table-widget']//select[@class='form-control c-select']", "-1")
    webdriver.wait_until_page_contains_element(
        "//a[@id='investments-table-object_last' and contains(@class, 'disabled')]",
        timedelta(seconds=30))
    print('Show all entries.')
    table_rows = webdriver.find_elements('//div[@id="investments-table-object_wrapper"]//tbody//tr')
    for element in table_rows:
        urls = element.find_elements_by_xpath('.//a')
        for url in urls:
            link = url.get_attribute('href')
            file_name = f"{link[link.rfind('/') + 1:]}.pdf"
            download_file_single(webdriver, link, file_name)
            values = rpa_parse_pdf_data(f'output/{file_name}')
            print(f"Check {values['uii']}: {values['investment']} in table.")
            if (values['uii'] in element.text) and (values['investment'] in element.text):
                print('Table contains both values.')
            elif (values['uii'] in element.text) and (values['investment'] not in element.text):
                print(f'Table contain only {values["uii"]}.')
            elif (values['uii'] not in element.text) and (values['investment'] in element.text):
                print(f'Table contain only {values["investment"]}.')
            else:
                print("Table doesn't contain values.")
    table = [i.text for i in webdriver.find_elements('//div[@id="investments-table-object_wrapper"]//tbody//tr//td')
             if i.text != '']
    return table


def write_agencies(workbook: Files, table: list):
    workbook.create_worksheet('Agencies')
    row_num = 1
    col_num = 1
    for data in table:
        if data != '':
            if col_num == 8:
                row_num += 1
                col_num = 1
            workbook.set_cell_value(row_num, col_num, data)
            col_num += 1
    print('Agencies: Done!')


def download_file_single(webdriver: Selenium, link: str, filename: str):
    webdriver.execute_javascript(f'window.open("{link}")')
    webdriver.switch_window("NEW")
    webdriver.wait_until_element_is_visible(
        '//div[@id="investment-quick-stats-container"]//div[@id="business-case-pdf"]', timedelta(seconds=30))
    webdriver.click_element('//div[@id="investment-quick-stats-container"]//div[@id="business-case-pdf"]')
    webdriver.wait_until_element_does_not_contain(
        '//div[@id="investment-quick-stats-container"]//div[@id="business-case-pdf"]', 'Generating PDF...',
        timedelta(seconds=30))
    wait_for_downloads_file(filename)
    webdriver.close_window()
    webdriver.switch_window("MAIN")


def main():
    print('Initialize.')
    with open('config.txt') as file:
        dep_to_scrap = file.readline().strip()
    driver = Selenium()
    driver.auto_close = False
    driver.set_download_directory(f'{os.getcwd()}/output')
    print(f'Start!\nDepartment to search:{dep_to_scrap}')
    workbook = Files()
    workbook.create_workbook('write_data', fmt='xlsx')
    values = get_departments(driver)
    write_budgets(workbook, values['departments'], values['budgets'])
    table = get_agencies_table(driver, dep_to_scrap)
    write_agencies(workbook, table)
    workbook.save_workbook('output/write_data.xlsx')
    workbook.close_workbook()
    driver.close_all_browsers()


if __name__ == '__main__':
    main()
