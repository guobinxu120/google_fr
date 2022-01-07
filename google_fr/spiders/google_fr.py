# -*- coding: utf-8 -*-
from scrapy import Spider, Request

import sys
import re, os, requests, urllib
from scrapy.utils.response import open_in_browser
from collections import OrderedDict
import time
from xlrd import open_workbook
from shutil import copyfile
import json, re, csv
from scrapy.http import FormRequest
from scrapy.http import TextResponse

def readExcel(path):
    wb = open_workbook(path)
    result = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        herders = []
        for row in range(0, number_of_rows):
            values = OrderedDict()
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if row == 0:
                    herders.append(value)
                else:

                    values[herders[col]] = value
            if len(values.values()) > 0:
                result.append(values)
        break

    return result


class AngelSpider(Spider):
    name = "google_fr"
    start_urls = 'https://www.google.fr/'
    count = 0
    use_selenium = True
    ean_codes = readExcel("input.xlsx")
    models = []
    headers = ['EAN', 'NAME']
    for i in range(6):
        headers.append('Price Shop '+str(i+1))
        headers.append('Name Shop '+str(i+1))
        headers.append('Link Shop '+str(i+1))
    for code in ean_codes:
        item = OrderedDict()
        item['EAN'] = code['EAN']
        # item['NAME'] = code['NAME']
        # item['shot description'] = ''
        # item['Marque'] = ''
        # item['Numéro de modèle'.decode("utf-8", "ignore")] = ''
        # item['Couleur'] = ''
        # item['Poids de l\'article'.decode("utf-8", "ignore")] = ''
        # item['Dimensions du produit (L x l x h)'] = ''
        # item['Puissance'] = ''
        # item['Voltage'] = ''
        # item['Fonction arrêt automatique'.decode("utf-8", "ignore")] = ''
        # item['Classe d\'énergie'.decode("utf-8", "ignore")] = ''
        models.append(item)
    def start_requests(self):
        header = {
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'en-US,en;q=0.9',
            'cookie': 'CONSENT=WP.26abeb.26ac93.26b0cf; NID=126=ErJ4bfuVKuf3d09WuXthW4r1PqTP2Wiwsn30J8DrT8S-2VDt1rIgJc3wZL1g1X7bv_-_bV_rkZ4KmLMVPxvCvFjtRNzjfT0RHz9ALV7oOqtW0NVm9NeL9vluiKLfTgJ-J7yy_CRrNg; DV=Q2JacmnoE_0UYKEIkm8vVLYXcZWaIxY; 1P_JAR=2018-3-18-15',
            'referer': 'https://www.google.fr/',
            'upgrade-insecure-requests': '1',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.162 Safari/537.36',
            'x-client-data': 'CIy2yQEIpLbJAQjAtskBCKmdygEIqKPKAQ=='
        }
        yield Request(self.start_urls, callback=self.parse, headers= header)
    def parse(self, response):
        for i, val in enumerate(self.models):
            ern_code = val['EAN']
            url = ''
            if ern_code != '':
                try:
                    ern_code = str(int(val['EAN']))
                except:
                    ern_code = str(val['EAN'])
                url ='https://www.google.fr/search?q={}&oq={}'.format(ern_code, ern_code)
            # elif val['NAME'] != '':
            #     ern_code = val['NAME'].replace(' ', '%20')
            #     url ='https://www.google.fr/search?q={}&oq={}'.format(ern_code, ern_code)
            yield Request(url, callback=self.parse1, meta={'order_num':i})

    def parse1(self, respond):
        index =respond.meta['order_num']
        item = OrderedDict()
        item['EAN'] = self.models[index]['EAN']
        # item['NAME'] = self.models[index]['NAME']


        prices_text = respond.xpath('//*[contains(@jsaction, "mouseover:pla.au;mouseout:pla.ru;pla.go")]/span[1]/text()').extract()
        prices = []
        for price_row in prices_text:
            prices.append(price_row.replace(' ', ''))

        urls = respond.xpath('//*[contains(@class, "jackpot-merchant")]/a[2]/@href').extract()
        names = respond.xpath('//*[contains(@class, "jackpot-merchant")]/a[2]/span[1]/text()').extract()


        for i in range(6):
            item['Price Shop '+str(i+1)] = ''
            item['Name Shop '+str(i+1)] = ''
            item['Link Shop '+str(i+1)] = ''
            if i < len(prices):
                item['Price Shop '+str(i+1)] = prices[i].replace(',', '.').replace('€', '')
            if i < len(names):
                item['Name Shop '+str(i+1)] = names[i]
            if i < len(urls):
                item['Link Shop '+str(i+1)] = urls[i]
        self.models[index] = item
        yield item





           