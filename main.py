from bs4 import BeautifulSoup
import requests
import openpyxl


#Эта функция создает новый xlsx файл/открывает существующий и записывает туда базовые колонки
def open_and_create_xlsx():
    filepath = "output.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    my_list = ['ID', 'Наименование Товара', 'Ссылка на товар', 'Регулярная цена', 'Промо-цена', 'Бренд']
    ws.append(my_list)
    wb.save(filepath)
    return wb

#Находит последнюю страницу, возвращает integer
def find_last_page_number(response_from_server):
    max_page_number = 0
    bs = BeautifulSoup(response.text, "lxml")
    numbers = bs.findAll('a', "v-pagination__item catalog-paginate__item nuxt-link-active")

    for i in numbers:
        if int(i.text) > max_page_number:
            max_page_number = int(i.text)
    return max_page_number


# Возвращает list ссылок на все страницы
def get_all_pages(last_page_number):
    urls = []
    for i in range(1, last_page_number + 1):
        urls.append("https://online.metro-cc.ru/category/ovoshchi-i-frukty/frukty?page=" + str(i))
    return urls

#Возвращает list Брендов
def get_all_brands(resp):
    brands = []
    bs = BeautifulSoup(resp.text, "lxml").find('div', {'data-filter-group': "Бренд"})
    br = bs.findAll('span', 'catalog-checkbox__text is-clickable')
    for i in br:
        brands.append(i.text.strip())
    return brands

#Для каждого товара находит необходимые атрибуты. Ничего не вовзращает
def parse_through_page(page,brands,wb):
    response = requests.get(page)
    bs = BeautifulSoup(response.text, "lxml")
    tovari = bs.findAll('div',
                        'catalog-2-level-product-card product-card subcategory-or-type__products-item catalog--common offline-prices-sorting--best-level with-prices-drop')
    for tovar in tovari:
        if tovar != None:
            id = tovar['data-sku']
            name = tovar.find('span', 'product-card-name__text').text
            brand = ''
            for i in brands:
                if i.lower() in name.lower():
                    brand = i
            link = 'https://online.metro-cc.ru/' + tovar.find('a', {'data-qa': 'product-card-photo-link'})['href']
            price = tovar.find('span',
                               'product-price nowrap product-card-prices__actual style--catalog-2-level-product-card-major-actual catalog--common offline-prices-sorting--best-level')
            if price == None:
                price = tovar.find('span',
                                       'product-price nowrap product-card-prices__old style--catalog-2-level-product-card-major-old catalog--common offline-prices-sorting--best-level')

            new_price = tovar.find('span',
                                   'product-price nowrap product-card-prices__actual style--catalog-2-level-product-card-major-actual color--red catalog--common offline-prices-sorting--best-level')
            if price !=None:
                price = ''.join(filter(str.isdigit,price.text.strip()))
            if new_price != None:
                new_price = ''.join(filter(str.isdigit,new_price.text.strip()))
            dump_into_xlsx(wb, id, name, link,price,new_price,brand)

#Записывает данные из функции parse_through_page в xlsx
def dump_into_xlsx(xlsx, id, name, link, reg_pr, promo_pr, brand):
    xl = xlsx.active
    xl.append([id,name,link,reg_pr,promo_pr,brand])
    xlsx.save('output.xlsx')

#Время работы программы можно сократить при использовании мультипроцессинга (в комментариях). К сожалению, мультипроцессинг не поддерживает работу с xlsx
if __name__ == '__main__':
    url = 'https://online.metro-cc.ru/category/ovoshchi-i-frukty/frukty?page=1'

    response = requests.get(url)
    last_page = find_last_page_number(response)
    brands = get_all_brands(response)
    wb = open_and_create_xlsx()

    for i in get_all_pages(last_page):
        parse_through_page(i, brands, wb)
    # with Pool(5) as p:
    #    (p.map(partial(parse_through_page, brands=brands, wb = wb), get_all_pages(last_page)))
