__author__ = 'Acheron'

import os
import sys
import time

import requests
from lxml import html
import urllib.parse
import json
import xlsxwriter
import logging



class BrainParser:

    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.worksheets = {}

    def login(self):
        login_url = 'http://opt.brain.com.ua/dealer/login'

        payload = {
                    "email": self.username,
                    "password": self.password,
                  }

        session_requests = requests.session()

        result = session_requests.get(login_url)

        tree = html.fromstring(result.text)
        authenticity_token = tree.xpath("//head/meta[@name='csrftoken']")[0]
        self.csrftoken = authenticity_token.attrib['content']

        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.116 Safari/537.36 OPR/40.0.2308.81",
               "Accept":"application/json, text/javascript, */*; q=0.01",
               "Accept-Encoding":"gzip, deflate, lzma",
               "Connection":"keep-alive",
               "Connection-Type":"application/x-www-form-urlencoded; charset=UTF-8",
               "Host":"opt.brain.com.ua",
               "Origin":"http://opt.brain.com.ua",
               "Referer":login_url,
               "X-Requested-With":"XMLHttpRequest",
               "X-CSRF-Token": self.csrftoken,
               "Accept-language":"ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4"
               }

        result = session_requests.post(
        login_url,
        data = payload,
        headers = headers)

        print(result.text)

        result = session_requests.get('http://opt.brain.com.ua')

        self.session_requests = session_requests
        return True


    def _getData(self, category=None, search=None):
        if category and not search:
            ref_url = 'http://opt.brain.com.ua/category/{0}'.format(category)
            search = ''
        elif search and not category:
            encoded_search = urllib.parse.urlencode({'s': search})
            ref_url = 'http://opt.brain.com.ua/search/detail?{0}'.format(encoded_search)
            category = ''
        else:
            raise AttributeError('Category OR search keyword should be used, not both!')

        page_num = '100'
        post_payload = {'search_str': search,
                         'route_targetID': '147',
                         'max_price': '0',
                         'just_discounted': '0',
                         'filters': '',
                         'delivery_method': 'route_delivery',
                         'mode': 'detail',
                         'sort_order': '',
                         'home_targetID': '0',
                         'currency': 'UAH',
                         'is_action': '0',
                         'page_num': page_num,
                         'delivery_days': '30',
                         'category': category,
                         'avail_type': 'delivery',
                         'just_bonused': '0',
                         'is_new': '0',
                         'regionID': '2',
                         'page_count': '100',
                         'iprice': '0',
                         'min_price': '0',
                         'sort_field': '',
                         'targetID': '29',
                         'ddp': '0'
                         }

        headers = {
            'Host':'opt.brain.com.ua',
            'Origin':'http://opt.brain.com.ua',
            'Referer':ref_url,
            "X-CSRF-Token": self.csrftoken,
            'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.116 Safari/537.36 OPR/40.0.2308.81'
        }

        url = 'http://opt.brain.com.ua/product_list'
        result = self.session_requests.post(url, data = post_payload, headers = headers)
        page_data = json.loads(result.text)

        # total/items_per_page + 1 if gt than page_count
        total_pages = int(int(page_data['total_count']) / 100)
        if (int(page_data['total_count']) - total_pages) != 0:
            total_pages += 1

        vendors = {}

        for page_idx in range(1, total_pages+1):
            print('Parsing {0} of {1} pages'.format(page_idx, total_pages))
            post_payload['page_num'] = page_idx
            result = self.session_requests.post(url, data = post_payload, headers = headers)

            if result.status_code == 200:
                page_data = json.loads(result.text)
                for item in page_data['rows']:

                    items_dict = {
                        'Vendor': None,
                        'NameRu': None,
                        'OptPrice': None,
                        'Articul': None,
                        'ProductCode': None,
                        'ImgName': None,
                        'ImgUrl': None
                        }

                    # All fields that would be used in price
                    items_dict['Vendor'] = item['Vendor']
                    items_dict['NameRu'] = item['NameRu']
                    items_dict['OptPrice'] = item['OptPrice']
                    items_dict['Articul'] = item['Articul']
                    items_dict['ProductCode'] = item['ProductCode']

                    try:
                        image = item['Thumbnail']
                        items_dict['ImgName'] = image
                        image_code = image.split('_')[0][-2:]
                        image_url = 'http://brain.com.ua/static/images/prod_img/{0}/{1}/{2}'.format(image_code[0], image_code[1], image)
                        items_dict['ImgUrl'] = image_url
                    except:
                        logging.debug('No image available for product: {0}'.format(item['NameRu']))

                    if item['Vendor'] in vendors.keys():
                        vendors[item['Vendor']].append(items_dict)
                    else:
                        vendors[item['Vendor']] = list()
                        vendors[item['Vendor']].append(items_dict)
            else:
                print('Error occurred while trying to get data.\nError code: {0}'.format(result.status_code))
                return None
            print('{0} of {1} done!'.format(page_idx, total_pages))
            # break <- for testing purpose to get only first page
        return vendors

    # downloading images
    def _pooler(self, vendors_dic):
        for k in sorted(vendors_dic):
            lst = vendors_dic[k]
            for item in lst:
                if item['ImgUrl']:
                    file_name = item['ImgName']
                    r = self.session_requests.get(item['ImgUrl'], stream=True)
                    if r.status_code == 200:
                        with open('images/{0}'.format(file_name), 'wb') as f:
                            for chunk in r.iter_content():
                                f.write(chunk)
                yield k, item

    # writing grabbed data to xlsx file
    def wtireXLS(self, vendors_dic, file_name='test.xlsx'):
        workbook = xlsxwriter.Workbook(file_name)

        img_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        wide_text_format = workbook.add_format({
                                        'bold': 1,
                                        'border':1,
                                        'align': 'center',
                                        'valign': 'vcenter'})

        normal_title_format = workbook.add_format({
                                        'border':1,
                                        'align': 'center',
                                        'valign': 'vcenter'})

        print('Downloading images.')
        for vendor, item in self._pooler(vendors_dic):
            if vendor in self.worksheets.keys():
                index, worksheet = self.worksheets[vendor]
                index += 3
                self.worksheets[vendor] = (index, worksheet)
            else:
                index = 3
                worksheet = workbook.add_worksheet(vendor)
                self.worksheets[vendor] = (index, worksheet)

                worksheet.set_column('A:A', 11)
                worksheet.set_column('B:B', 65)
                worksheet.set_column('C:C', 15)


                worksheet.merge_range('A1:A2', 'Фото', wide_text_format)
                worksheet.merge_range('B1:B2', 'Наименование', wide_text_format)
                worksheet.merge_range('C1:C2', 'Цена', wide_text_format)

            worksheet.merge_range('A{0}:A{1}'.format(index, index+2), '', img_format)
            worksheet.insert_image('A{0}'.format(index), 'images/{0}'.format(item['ImgName']),
                                   {'x_scale': 0.6, 'y_scale': 0.5, 'x_offset': 10, 'y_offset': 5,
                                    'positioning': 1})

            worksheet.write('B{0}'.format(index), item['NameRu'])
            worksheet.write('B{0}'.format(index+1), 'Артикуль: {0}'.format(item['Articul']))
            worksheet.write('B{0}'.format(index+2), 'Код: {0}'.format(item['ProductCode']))

            worksheet.merge_range('C{0}:C{1}'.format(index, index+2), item['OptPrice'], normal_title_format)

        workbook.close()
        return True




if __name__ == '__main__':
    DIR = os.path.dirname(os.path.realpath(sys.argv[0]))
    DATE = time.strftime("%d.%m.%Y")

    # filename
    FILE = 'toner_price_{0}.xlsx'.format(DATE)

    parser = BrainParser('username', 'password')
    parser.login()

    # chose category or search-key word
    vendors = parser._getData(category='Tonery_barabany-c1558')

    if vendors:
        parser.wtireXLS(vendors,file_name=FILE)
        print('Saved to: {0}\\{1}'.format(DIR, FILE))
    else:
        print('Exiting on error')