from selenium import webdriver
import xlsxwriter

print 'URL=',
url = raw_input()
print 'range='
pages = (int(raw_input()), int(raw_input()))
print 'output file=',
filename = raw_input()


driver = webdriver.Chrome()

totalProducts = {}

idx = 0

for page in range(pages[0], pages[1] + 1):
    driver.get(url + '&_pgn=' + str(page) + '&_ipg=200&rt=nc')
    products = driver.find_elements_by_class_name("sresult")
    for product in products:
        info = product.find_elements_by_css_selector(".lvdetails")
        for inf in info:
            provider = inf.text
            if "From" in provider:
                if "Vietnam" in provider:
                    productName = product.find_element_by_class_name("vip").get_attribute("title")[26:]
                    productUrl = product.find_element_by_class_name("vip").get_attribute("href")
                    totalProducts[str(idx)] = {
                        'name': productName,
                        'url': productUrl
                    }
                    idx = idx + 1

workbook = xlsxwriter.Workbook(filename + ".xlsx")
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, "Ten")
worksheet.write(0, 1, "Link")

row = 1

for product in totalProducts.values():
    print product
    worksheet.write(row, 0, product['name'])
    worksheet.write_string(row, 1, product['url'])
    row = row + 1

workbook.close()
