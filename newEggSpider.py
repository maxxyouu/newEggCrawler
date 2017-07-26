"""
TASK: store information of newegg website as csv/xlrx format
"""
from urllib.request import urlopen
from urllib.parse import urlparse
import csv

from bs4 import BeautifulSoup as soup
import xlsxwriter

UNKNOWN = 'N/A'


class IndividualItem:
    """
    individual item information
    """
    def __init__(self, thumbnail, title, brand, shipping, price):
        self.thumbnail = thumbnail
        self.title = title
        self.brand = brand
        self.shipping = shipping
        self.price = price

    def __repr__(self):
        """
        representation of a object
        :return: str
        """
        return '(thumbnail:{}, title:{}, brand:{}, shipping:{}, price:{})'.format(self.thumbnail,
                                                                                  self.title,
                                                                                  self.brand,
                                                                                  self.shipping,
                                                                                  self.price)
class CSVconverter:
    """
    csv writer
    """
    def __init__(self, csv_name=''):
        """
        @return None
        """
        self.csv_name = csv_name

    def csv_writer(self, data):
        """
        write self.data into the csv file format
        @param dict data: {page#: list of item obj}
        """
        for pageKey in data:  # write data according to column
            with open('Page{}.csv'.format(str(pageKey)), 'w') as data:
                columnNames = ['thumbnail', 'title', 'brand', 'shipping', 'price']
                writer = csv.DictWriter(data, fieldnames=columnNames)
                writer.writeheader()
                item = data[pageKey]
                col_data = {'thumbnail': item.thumbnail,
                            'title': item.title,
                            'shipping': item.shipping,
                            'price': item.price}
                writer.writerow(col_data)


class XlsxConverter:
    """
    data writer
    """
    def __init__(self, xlsxWorkbook_name=''):
        """
        @return None
        """
        pass

    def xlsx_writer(self, dict_):
        """
        write self.data into the xlsx file format
        @param dict data: {page#: list of item obj}
        """
        categoryBook = xlsxwriter.Workbook('data.xlsx')
        headerProperties = categoryBook.add_format({'bold':True,
                                                    'align':'center',
                                                    'valign':'vcenter'})
        dataCellProperties = categoryBook.add_format({'align':'center',
                                                    'valign': 'vcenter'})
        for pageKey in dict_:
            pageSheet = categoryBook.add_worksheet('Page{}'.format(str(pageKey)))
            pageSheet.write('A1', 'Thumbnail', headerProperties)  # write header
            pageSheet.write('B1', 'Title', headerProperties)  # write header
            pageSheet.write('C1', 'Shipping', headerProperties)  # write header
            pageSheet.write('D1', 'Branding', headerProperties)  # write header
            pageSheet.write('E1', 'Price', headerProperties)
            row, col = 1, 0
            for itemObj in dict_[pageKey]:
                pageSheet.write(row, col, itemObj.thumbnail, dataCellProperties)
                pageSheet.write(row, col + 1, itemObj.title, dataCellProperties)
                pageSheet.write(row, col + 2, itemObj.shipping, dataCellProperties)
                pageSheet.write(row, col + 3, itemObj.brand, dataCellProperties)
                pageSheet.write(row, col + 4, itemObj.price, dataCellProperties)
                row += 1
        categoryBook.close()

class NewEggSpider:
    """
    newEgg crawler
    """
    def __init__(self, name='New Egg Web Spider'):
        """
        initialize the spider by gather all the categories and subcategory links as dictionarys
        @param dict categories: {cat_name: list of subcategory tuples}
        @param dict sub_categories: {subcat_name: url link str}
        @param str name: name of the crawler
        @return None
        """
        def _categoryCollections():
            # @return dict{category: list of subcategories tuples(subcategoryName, link)}
            result = {}
            category_containers = self.html_soup.select('dd.main-nav-subItem')
            for category_container in category_containers:
                try:
                    catetory_name = category_container.select_one('div.popover-wrap > a.main-nav-third-title').get_text()
                    # get list of subcategories
                    subs = [(subCategory.get_text(), subCategory['href']) 
                            for subCategory in category_container.select('div.main-nav-third-body a.main-nav-third-title')]
                except:
                    titleLink = category_container.select_one('a')
                    catetory_name = titleLink.get_text()
                    subs = [titleLink['href']]
                if catetory_name not in result:
                    result[catetory_name] = subs
            return result
        def _subCategoriesCollections():
            # @return dict{subcategoryName: str of urllink}
            return {link.get_text():link['href']
                    for link in self.html_soup.select('li.main-nav-third-item > a.main-nav-third-title')}

        self.html_soup = soup(urlopen('https://www.newegg.ca/'), 'html.parser')
        self.categories = _categoryCollections()
        self.sub_categories = _subCategoriesCollections()
        self.name = name
        self.duplicates = set()


    def _threholdFunc(self, totalItems, threhold=None):
        """return number of items that needed for test
        @param int totalItems: number of items in the newEgg webpage
        @param int threhold: int|None limits
        @return int
        """
        if isinstance(threhold, int) and threhold <= totalItems:  #not none
            return threhold
        else:
            return totalItems


    def _getPages(self, categoryName, threhold=None):
        """ get all the pages' url under the productLink with threhold limit
        @param str categoryName: name of the category
        @param threhold int|None: maximum pages allowed to be returned, default is None
        @return list of [str of page urls]
        """
        def _constructUrl(pageNum, productLink):
            # @param pageNum int
            # @param productLink str
            pageFormat = 'Page-'
            defaultQuery = 'PageSize=36&order=BESTMATCH'
            parsedUrl = urlparse(productLink)
            # find the id
            path = parsedUrl.path
            id_index = path.rfind('ID')
            end_index = path.find('/', id_index)  # return -1 if '/' is not found
            if 'tid' in parsedUrl.query.lower():
                if end_index != -1:
                    url = (parsedUrl.scheme + '://' + parsedUrl.netloc + parsedUrl.path[:end_index] + '/' + 
                           pageFormat + str(pageNum) + '?' + parsedUrl.query + '&' + defaultQuery)
                else:
                    url = (parsedUrl.scheme + '://' + parsedUrl.netloc + parsedUrl.path + '/' + pageFormat + 
                           str(pageNum) + '?' + parsedUrl.query + '&' + defaultQuery)
            else:
                if end_index != -1:
                    url = (parsedUrl.scheme + '://' + parsedUrl.netloc + parsedUrl.path[:end_index] + '/' + 
                           pageFormat + str(pageNum) + '?' + defaultQuery)
                else:
                    url = (parsedUrl.scheme + '://' + parsedUrl.netloc + parsedUrl.path + '/' + pageFormat + 
                           str(pageNum) + '?' + defaultQuery)
            return url

        if categoryName in self.sub_categories or categoryName in self.categories:
            if categoryName in self.sub_categories:
                link = self.sub_categories[categoryName]
            else:
                link = self.categories[categoryName][0]  # get the subcategory link according to the name
            htmlSoup = soup(urlopen(link), 'html.parser')
            try: 
                pageNav = htmlSoup.findAll('div', {'id': 'page_NavigationBar'})[-1]
            except: 
                totalPages = 1  # only one page
            else: 
                totalPages = pageNav.findAll('div', {'class': 'btn-group-cell'})[-2].button.get_text()
            limit = self._threholdFunc(int(totalPages), threhold)
            return [_constructUrl(i, link) for i in range(1, limit + 1)]
        return []


    def getPageProducts(self, subCategoryName, pageNum=1):
        """get all items in a single page
        @param str subCategoryName: sub category name
        @param int pageNum: page number of the subCategory
        @return list of product items
        """
        def _getThumbnailPerItem(individualItemContainer):
            """get thumbnail picture for each item
            notes: individualItem has class 'item-container'
            @return the url of the thumbnail
            """
            try:
                return individualItemContainer.select_one('a.item-img > img')['src']
            except:
                return UNKNOWN
        def _getTitlePerItem(individualItemContainer):
            """get the title for each product
            @return str
            """
            try:
                return individualItemContainer.find('a', class_='item-title').get_text()
            except:
                return UNKNOWN
        def _getBrandPerItem(individualItemContainer):
            """get the brand name for each product
            @return url of the brand img
            """
            try:
                return individualItemContainer.select_one('a.item-brand > img')['src']
            except:
                return UNKNOWN
        def _getShippingPerItem(individualItemContainer):
            """get shipping price for each product
            @return str
            """
            try:
                string = individualItemContainer.find('li', class_='price-ship').get_text()
                result = ''
                for char in string:
                    if char in '.0123456789':
                        result += char
                return '0' if len(result) == 0 else result
            except:
                return UNKNOWN
        def _getPricePerItem(individualItemContainer):
            """get the price for each product
            @return price: str
            """
            def _check_prices(price_string):
                modified_result = ''
                for digit in price_string:
                    if digit in '0123456789':
                        modified_result += digit
                return modified_result
            try:
                price = individualItemContainer.find('li', class_='price-current')
                main_price = float(_check_prices(price.strong.get_text()))
                decimals = float(price.sup.get_text())
                return str(main_price + decimals)
            except:
                return UNKNOWN
        # get all individual item containers of a SINGLE PAGE
        if subCategoryName in self.sub_categories or subCategoryName in self.categories:
            allPages = self._getPages(subCategoryName)
            if pageNum <= len(allPages):  # make sure the page in range
                pageNeeded = allPages[pageNum - 1]  # the link at pageNum index position
                pageSoup = soup(urlopen(pageNeeded), 'html.parser')
                return [IndividualItem(_getThumbnailPerItem(container),
                                       _getTitlePerItem(container),
                                       _getBrandPerItem(container),
                                       _getShippingPerItem(container),
                                       _getPricePerItem(container))
                        for container in pageSoup.select('div.item-container')]
        return []


    def getSingleSubCategoryProducts(self, subCategoryName):
        """
        get all products from a single sub category
        notes: a single subcategory has multiple pages
        ex: https://www.newegg.ca/Desktop-NAS/SubCategory/ID-124/Page-2
        @param str subCategoryName: name of the subCategory
        @return dict of list {page#, list of product items}
        """
        if subCategoryName in self.sub_categories or subCategoryName in self.categories:
            return {pageNum+1: self.getPageProducts(subCategoryName, pageNum + 1)
                    for pageNum in range(0, len(self._getPages(subCategoryName)))}
        return {}

    def getCategoryProducts(self, categoryName):
        '''
        note: a main category has multiple subcategories, each subcategory has multiple pages
        @param categoryNanme: str of the a category name
        @return list of dict products
        '''
        #find the div that has the categoryName
        if categoryName in self.sub_categories or categoryName in self.categories:
            return [self.getSingleSubCategoryProducts(name)
                    for name, _ in self.categories[categoryName]]
        return []  # invalid categoryName

    def crawlAllData(self):
        """crawl all the data from newEgg website
        @return dict{}: {categoryName: [dict{subcategory, list of items}...]}
        """
        result = {}
        # get all category links
        for main_cat in self.categories:
            sub_cate_list = []
            for sub_cate_name, _ in self.categories[main_cat]:
                sub_cate_dict = {}
                for page in range(len(self._getPages(sub_cate_name))):
                    # list of items for each subcategory
                    sub_cate_dict[sub_cate_name] = (sub_cate_dict.get(sub_cate_name, []) +
                                                    self.getPageProducts(sub_cate_name, page + 1))
                sub_cate_list.append(sub_cate_dict)
            result[main_cat] = sub_cate_list
        return result


    def convertDataAsCSV(self, data):
        CSVconverter().csv_writer(data)


    def convertDataAsXlsx(self, data):
        print(data)
        XlsxConverter().xlsx_writer(data)

    def __repr__(self):
        return self.name


if __name__ == '__main__':
    name = 'Wireless Routers'
    name2 = 'Powerline Networking'
    spider = NewEggSpider()
    data = spider.getSingleSubCategoryProducts(name2)
    spider.convertDataAsXlsx(data)
    # spider.convertDataAsXlsx(data)
