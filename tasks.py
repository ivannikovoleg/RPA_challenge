import time
import os
from datetime import timedelta
from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from RPA.Excel.Files import Files


def wait_for_downloads(path: str):
    print('Waiting for downloads', end='')
    while any([filename.endswith('.crdownload') for filename in
               os.listdir(path)]):
        time.sleep(2)
        print(".", end="")
    print('\nAll files downloaded!')


def rpa_parse_pdf_data(filename: str) -> list:
    pdf = PDF()
    investment_name = 'Name of this Investment: '
    uii = '2. Unique Investment Identifier (UII): '
    end = 'Section B:'
    text = pdf.get_text_from_pdf(filename)
    page_text = text[1]
    investment_name_index = page_text.find(investment_name) + len(investment_name)
    uii_index_start = page_text.find(uii)
    uii_index_end = uii_index_start + len(uii)
    end_index = page_text.find(end)
    return [page_text[investment_name_index: uii_index_start],
            page_text[uii_index_end:end_index]]


def check_entries_by_elements(path: str, table: list):
    file_list = [os.path.join(path, file) for file in os.listdir(path) if file.endswith('.pdf')]
    for file in file_list:
        v = rpa_parse_pdf_data(file)
        print(f'Checking entries {v[1]}: {v[0]}')
        for element in table:
            if v[0] == element or v[1] == element:
                print(f'{element} in table.')


def get_departments(webdriver: Selenium):
    webdriver.open_headless_chrome_browser('https://itdashboard.gov/')
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


def download_file(webdriver: Selenium):
    urls = webdriver.find_elements('//div[@id="investments-table-object_wrapper"]//tbody//tr//a')
    links = [url.get_attribute("href") for url in urls]
    for link in links:
        webdriver.go_to(link)
        webdriver.wait_until_element_is_visible(
            '//div[@id="investment-quick-stats-container"]//div[@id="business-case-pdf"]', timedelta(seconds=30))
        webdriver.click_element('//div[@id="investment-quick-stats-container"]//div[@id="business-case-pdf"]')
        webdriver.wait_until_element_does_not_contain(
            '//div[@id="investment-quick-stats-container"]//div[@id="business-case-pdf"]', 'Generating PDF...',
            timedelta(seconds=30))
        time.sleep(5)
        print('Downloading file.')


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
    download_file(driver)
    wait_for_downloads('output')
    driver.close_all_browsers()
    check_entries_by_elements('output', table)


if __name__ == '__main__':
    main()
