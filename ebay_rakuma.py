<<<<<<< HEAD
from collections import OrderedDict
import os
import eel
import pandas as pd
import re
import requests
import time
from googletrans import Translator
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException, ElementNotInteractableException
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from numpy import nan
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import subprocess
from bs4 import BeautifulSoup
from selenium.webdriver.support.select import Select

# 一覧ページから取得する情報
items_condition = []
items_name_ja = []
items_name_en = []
items_condition_en = []
items_price = []
top_img_urls = []
items_Descriptions = []
postage_price_list = []


# 各商品apiから取得する情報
items_id = []
items_image_url = []


def set_driver(driver_path):
    # Chromeドライバーの読み込み
    options = ChromeOptions()

    # 起動オプションの設定
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('log-level=3')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--incognito')          # シークレットモードの設定を付与
    options.add_argument('--user-agent=Chrome/87.0.42.88')
    options.add_argument('--single-process')
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-web-security')
    options.add_argument('--disable-desktop-notifications')
    options.add_argument('--disable-application-cache')
    options.add_argument("--disable-extensions")
    options.add_argument('--lang=ja')
    # options.add_argument('--ignore-certificate-errors')
    # options.add_argument('--ignore-ssl-errors')
    # options.add_argument('--incognito')          # シークレットモードの設定を付与

    # ChromeのWebDriverオブジェクトを作成する。
    return Chrome(driver_path, options=options)


class MyScraping():

    # 状態の判定
    # 与えられた文字列＝商品の状態によって特定の文字列を返す
    def judge_condition(self, string):
        # 未使用：Condition:Mint.
        # 未使用に近い：Condition:Near Mint.
        # 目立った傷や汚れなし：Condition:Excellent.
        # やや傷や汚れあり：Condition:Very Good.
        # 傷や汚れあり：Condition:Good.
        condition = ""

        if string == "未使用":
            condition = "Condition:Mint."

        elif string == "未使用に近い":
            condition = "Condition:Near Mint."

        elif string == "目立った傷や汚れなし":
            condition = "Condition:Excellent."

        elif string == "やや傷や汚れあり":
            condition = "Condition:Very Good."

        elif string == "傷や汚れあり":
            condition = "Condition:Good."

        elif string == "中古":
            condition = "Condition : Very Good."

        else:
            condition = nan

        return condition

    # 翻訳関数
    # 与えられた日本語の文字列を英語にして返す
    def translate(self, string):
        translator = Translator()
        result = translator.translate(string, src='ja', dest='en')
        return result.text

    # APIの設定
    # 楽天API。不要？
    def get_api(self, params):
        url = 'https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706?'
        result = requests.get(url, params=params)
        return result.json()

    # APIの設定
    # 楽天API。不要？
    def set_api_parameter(self):
        app_id = 1064309786905776599
        params = {
            "format": "json",
            "applicationId": [app_id],
            "formatVersion": 2,
        }
        return params

    # 画像URLの設定
    # 複数の画像URLを一つの文字列にまとめて返す
    def image_url_add(self, urls):
        image_url = ''
        for url in urls:
            image_url += url + "|"
        return image_url[:-1]

    # 翻訳前の不要ワードの設定
    # 第一引数の文字列に対して、第二引数のリストに含まれる文言を削除する処理を行う
    def delete_word_before(self, word, delete_before_list):
        for delete_word in delete_before_list:
            word = word.replace(delete_word, "")

        return word

    # 翻訳後の不要ワードの設定
    def delete_word_after(self, word, delete_after_list, add_name_list):

        code_regex = re.compile(
            '[!"#$&*+-.:;<=>?@^_`{|}~「」〔〕“”〈〉『』＆＊・＄＃＠。、？！｀＋￥◆★◇☆⇒%％◎○■/:×,]')
        word = code_regex.sub(" ", word)

        word = re.sub(r'\[([\w\d ]+)\]', '', word)
        word = re.sub(r'\(([\w\d ]+)\)', '', word)
        word = re.sub(r'\【([\w\d ]+)\】', '', word)

        word = word.replace("[]", "")
        word = word.replace("【】", "")

        word_list = word.split(" ")

        for delete_word in delete_after_list:
            if delete_word.lower() in word_list:
                word_list.remove(delete_word.lower())
            if delete_word.capitalize() in word_list:
                word_list.remove(delete_word.capitalize())

        for add_word in add_name_list:
            if add_word.lower() in word_list:
                word_list.remove(add_word.lower())
            if add_word.capitalize() in word_list:
                word_list.remove(add_word.capitalize())

        return word_list

    # ワードの追加
    def add_word(self, word_num, word, add_word_list):

        if len(word) == word_num:
            return word

        elif len(word) > word_num:
            while len(word) + 1 >= word_num:
                word_list = word.split(' ')
                word = ' '.join(word_list[:-1])

            if len(word) == word_num:
                return word

            else:
                word = word + ' '
                for add_word in add_word_list:
                    if add_word in word:
                        continue
                    if len(word) + len(add_word) >= word_num:
                        return word
                    word = word + add_word + " "
                else:
                    return word

        elif len(word) < word_num:
            for add_word in add_word_list:
                if add_word in word:
                    continue
                if len(word) + len(add_word) >= word_num:
                    return word
                word = word + add_word + " "
            else:
                return word

    # 商品説明文字列作成
    def Description_string(self, Description):
        html_text = '<font rwr=\"1\" size=\"4\" style=\"font-family:Arial\"><div class=\"banner_topr\" style=\"height: 200px; width: 100%; background-image: url(\'https://www.saatscommerce.com/te_rsc/XOther_Japanese_c04/us/images/c_top.jpg\'); background-repeat: no-repeat; background-size: cover; \"></div><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>Description</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"></p><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><span style=\"font-size: large;\"><span style=\"font-family: &quot;Noto Sans JP&quot;, sans-serif; background-color: rgb(255, 255, 255);\">Please see the attached photographs for detailed item conditions.<br><br>●●<br><br><b>This product is 100% authentic.<br>We have several other items in stock, so please contact us.<br><br><br>[Customer\'s Feedback!]<br>\"OUTSTANDING SELLER! Fast shipping. Thanks so much for a wonderful transaction.\"<br>\"Fast delivery, perfect product. 100% satisfied. 5 star to seller. Thank you..\"<br>\"Totally worth it!!!thanks again awesome deck!!It arrived thanks!!\"<br>\"Amazing communication and ship speed. Items exactly as described. A+++ Thank u!!\"<br>\"Really beautiful transaction and super fast shipment from Japan, thank you!\"<br>\"Excellent service provided by the seller from start to end! Highly Recommended!\"<br>\"***Highly recommend to purchase from this seller*** Really friendly and reliable\"<br><br>...Thank You!</b></span><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><span style=\"font-size: large;\"><span style=\"font-family: &quot;Noto Sans JP&quot;, sans-serif; background-color: rgb(255, 255, 255);\"><br></span></span></p></h2><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>Payment</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\">We only accept PayPal payments.<br>Please pay within 5 days after auction is finished.</p><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><br></p></h2><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>Shipping</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"></p><br style=\"font-size: 14px; font-weight: 400;\"><div id=\"shipping_duration_table\" style=\"font-size: 14px; font-weight: 400;\"><b>Shipping Duration</b><table class=\"table_shipping\" style=\"width: 770px; border: 1px solid rgb(0, 0, 0); border-collapse: collapse;\"><thead><tr><th style=\"width: 145px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">Shipping<br>Method</th><th style=\"width: 145px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">Asia, Middle East</th><th style=\"width: 145px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">America,Canada<br>Australia,Mexico</th><th style=\"width: 145px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">Europe<br>Russia</th><th style=\"width: 144px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">Other</th></tr></thead><tbody id=\"shipping_duration_tbody\"><tr><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">DHL, FedEX</td><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">2 - 3 days</td><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">3 - 5 days</td><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">3 - 4 days</td><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">5 - 7 days</td></tr></tbody></table></div><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">w/Tracking.</span><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">Expedited Shipping.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">We will ship to world wide by EMS or FedEX or DHL.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">We will decide the shipping method.</span><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">Please contact us in advance if specified.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">It may be delayed depending on the delivery area, but we will ship as soon as possible.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">We will ship your item within 2-10 business days, after your payment.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">We can ship ONLY to your Paypal address.</span><p style=\"font-size: 14px; font-weight: 400;\"></p></h2><div><span style=\"font-size: 14px; font-weight: 400;\"><br></span></div><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>Term of Sale</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\">We can accept return within 30 days after you receive it.<br>Please make sure to contact us first if you want to return it.<br>It is buyer\'s responsibility for the return shipping cost when sending back the item to us.<br>The item has to be returned in the original condition from the time you received.<br>No returns on products that show any evidences of use.</p><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><br></p></h2><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>About Us</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\">A specialty shop offering a wide variety of carefully selected Japanese brand items.<br>Even if you think to yourself, \"But the item I want is only available in Japan.\", make sure to consult with us before giving up hope.<br>We will look for and suggest the best products on your behalf.</p><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><br></p></h2><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>International Buyers - Please Note:</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\">Import duties, taxes and charges are not included in the item price or shipping charges. These charges are the buyer\'\'s responsibility.<br>Please check with your country\'\'s customs office to determine what these additional costs will be prior to bidding/buying.<br>These charges are normally collected by the delivering freight (shipping) company or when you pick the item up - do not confuse them for additional shipping charges.<br><br>If your country is a non-shipping country, please contact us.<br>I will estimate the shipping cost.</p></h2></font>'
        return html_text.replace('●●', Description)

    # 一覧ページから情報を取り出す
    # 商品一覧から取り出す情報
    # ・商品名
    # ・商品の状態（中古とか可,良いとか）
    # ・値段、送料

    def get_top_detail(self, driver, wait, main_url, num_list):
        # ファイルの読み込み
        delete_before_file = pd.read_csv(
            './' + eel.delete_word_before()(), header=None, names=['削除するワード_before'])
        delete_before_list = delete_before_file['削除するワード_before'].tolist()
        delete_after_file = pd.read_csv(
            './' + eel.delete_word_after()(), header=None, names=['削除するワード_after'])
        delete_after_list = delete_after_file['削除するワード_after'].tolist()
        add_word_file = pd.read_csv(
            './' + eel.fill_in_word()(), header=None, names=['追加するワード'])
        add_word_list = add_word_file['追加するワード'].tolist()

        # 商品名の処理
        add_name = eel.add_word()()
        add_name_len = 80 - len(add_name)
        add_name_split_dot = add_name.split(',')
        add_name_split_brank = add_name.replace(",", " ").split(" ")

        # ページごと
        for i in range(int(num_list[0]), int(num_list[1]) + 1):
            eel.view_log_js("\n" + str(i) + "ページ目開始")
            page_url = main_url + "&b=" + str(i+(i*100-100))
            driver.get(page_url)
            wait.until(EC.visibility_of_element_located(
                (By.CLASS_NAME, "Products__items")))
            # items = driver.find_elements_by_class_name('Products__items')

            html = driver.page_source.encode('utf-8')

            # ＝＝＝＝＝＝＝＝＝＝追記分＝＝＝＝＝＝＝＝＝＝＝＝
            soup = BeautifulSoup(html, 'lxml')

            items = soup.find_all("li", class_="Product")
            print(len(items))

            for item in items:
                item_id = item.findAll("a")[0].get("data-auction-id")
                # print('============================================')
                # print(item_id)

            item_counter = 0
            # アイテムごと
            for j, item in enumerate(items):
                item_id = item.findAll("a")[0].get("data-auction-id")
                page_url = 'https://page.auctions.yahoo.co.jp/jp/auction/' + item_id
                driver.get(page_url)
                wait.until(EC.visibility_of_element_located(
                    (By.CLASS_NAME, 'ProductImage__inner')))
                html = driver.page_source.encode('utf-8')
                soup = BeautifulSoup(html, 'lxml')
                Price_buynow = soup.find_all("div", class_="Price--buynow")
                if (len(Price_buynow) == 0):
                    eel.view_log_js(str(j + 1) + "/" +
                                    str(len(items)) + "商品目 即決価格なし 除外")
                    continue
                print('============================================')
                print('item_id')
                print(item_id)

                # 商品のコンディションID M列
                # 商品の状態情報抽出
                item_condition_element = soup.find_all(
                    "tr", class_="ProductTable__row")
                # テキスト化
                item_condition_text = re.sub(
                    r'[\n ]', '', item_condition_element[1].findAll("a")[0].text)
                # 除外対象か判別。対象の場合飛ばして次の商品へ移動
                if item_condition_text == '全体的に状態が悪い':
                    eel.view_log_js(str(j + 1) + "/" +
                                    str(len(items)) + "商品目 状態が悪いため除外")
                    continue
                    # コンディションIDを設定
                item_condition_id = '3000'

                # 商品名 N列
                # 商品名抽出
                item_name_ja_element = soup.find_all(
                    "h1", class_="ProductTitle__text")
                item_name_ja = item_name_ja_element[0].text
                item_name_ja_delete = self.delete_word_before(
                    item_name_ja, delete_before_list)
                item_name_en = self.translate(item_name_ja_delete)
                item_name_en_delete_list = self.delete_word_after(
                    item_name_en, delete_after_list, add_name_split_brank)
                item_name_en = ' '.join(
                    OrderedDict.fromkeys(item_name_en_delete_list))
                item_name_en = self.add_word(
                    add_name_len, item_name_en, add_word_list)
                item_name_en = add_name_split_dot[0] + " " + \
                    item_name_en + " " + add_name_split_dot[1]
                item_name_en = re.sub(r'([\s]+)', ' ', item_name_en)
                item_name_en = item_name_en.title()

                # 商品識別ID O列
                # 取得済み

                # 商品の値段 Q列
                # 配送料チェック
                postage_price_free_element = soup.find_all(
                    "span", class_="Price__postageValue--free")
                # 配送料無料ではない商品をピックアップする。
                if (len(postage_price_free_element) == 0):
                    # 着払いではない商品をピックアップする。
                    postage_price_text_element = Price_buynow[0].findAll(
                        "span", class_="Price__postageValue")[0].text
                    # if(postage_price_text_element != '着払い'):
                    # 料金はpostage_priceで取得し計算する
                    postage_price_list.append(item_id)

                price_element = Price_buynow[0].findAll(
                    "dd", class_="Price__value")[0].text
                item_price = 0
                if '税込' in price_element:
                    price_element = re.split('[円税込]', price_element)[3]
                    item_price = re.sub(r'[, ]', '', price_element)
                else:
                    item_price = price_element.split('円')[0].replace(
                        ',', '').replace('\n', '')

                # 商品画像(10枚まで) W列
                item_images_element = soup.find_all(
                    "div", class_="ProductImage__inner")
                limit = 10
                urls = []
                for k, item_image in enumerate(item_images_element):
                    urls.append(item_image.findAll("img")[0].get("src"))
                    if k + 1 == limit:
                        break
                image_urls = self.image_url_add(urls)

                # 商品説明 X列
                item_Description_element = soup.find_all(
                    "div", class_="ProductExplanation__commentBody")
                item_Description = self.Description_string(
                    item_Description_element[0].text)

                # 商品の状態 Y列
                item_condition_en = self.judge_condition(
                    item_condition_text.replace('\n', ''))

                # 各データを格納
                items_condition.append(item_condition_id)
                items_name_ja.append(item_name_ja)
                items_name_en.append(item_name_en)
                items_id.append(item_id)
                items_price.append(item_price)
                top_img_urls.append(image_urls)
                items_Descriptions.append(item_Description)
                items_condition_en.append(item_condition_en)

                eel.view_log_js(str(j + 1) + "/" + str(len(items)) + "商品目")
                item_counter += 1

            else:
                eel.view_log_js(str(i) + "ページ目終了")
                eel.view_log_js(str(item_counter) + "件抽出")
                # print('items_condition')
                # print(len(items_condition))
                # # print(items_condition)

                # print('items_name_ja')
                # print(len(items_name_ja))
                # # print(items_name_ja)

                # print('items_name_en')
                # print(len(items_name_en))
                # # print(items_name_en)

                # print('items_id')
                # print(len(items_id))
                # # print(items_id)

                # print('items_price')
                # print(len(items_price))
                # # print(items_price)

                # print('top_img_urls')
                # print(len(top_img_urls))
                # # print(top_img_urls)

                # print('items_Descriptions')
                # print(len(items_Descriptions))
                # # print(items_Descriptions)

                # print('items_condition_en')
                # print(len(items_condition_en))
                # # print(items_condition_en)

        return items_name_ja, top_img_urls

    # 各商品の検索
    # 楽天apiから取るもの
    # ・商品の説明
    # ・画像URL
    # ・商品id
    def get_item_details(self, driver, wait, item_name, top_img_url):

        # APIの設定
        params = self.set_api_parameter()
        params["keyword"] = item_name

        try:
            item = self.get_api(params)["Items"]

            if item:
                item_image_url = item[0]["mediumImageUrls"]
                item_image_url = self.image_url_add(item_image_url)
                item_url = item[0]["itemUrl"]
                item_id = item_url.replace(
                    'https://item.rakuten.co.jp/', '')[:-1]

            else:
                url = 'https://www.rakuten.co.jp/'
                driver.get(url)
                wait.until(EC.visibility_of_element_located(
                    (By.CLASS_NAME, 'input-text--37z4R')))
                driver.find_element_by_class_name(
                    'input-text--37z4R').send_keys(item_name)
                driver.find_element_by_xpath(
                    '//*[@id="wrapper"]/div[5]/div/div/div/div[1]/div/div/div/a').click()

                try:
                    driver.find_element_by_xpath(
                        '/html/body/div[3]/div[2]/div[3]/div/div[1]/div[2]/h2/a').click()
                    wait.until(EC.visibility_of_element_located(
                        (By.CLASS_NAME, 'item_desc')))
                    item_id = driver.current_url.replace(
                        'https://item.rakuten.co.jp/', '')[:-1]
                    item_image_url = top_img_url

                except NoSuchElementException:
                    item_image_url = top_img_url
                    item_id = nan

                except TimeoutException:
                    item_image_url = top_img_url
                    item_id = nan

        except KeyError:
            url = 'https://www.rakuten.co.jp/'
            driver.get(url)
            wait.until(EC.visibility_of_element_located(
                (By.CLASS_NAME, 'input-text--37z4R')))
            driver.find_element_by_class_name(
                'input-text--37z4R').send_keys(item_name)
            driver.find_element_by_xpath(
                '//*[@id="wrapper"]/div[5]/div/div/div/div[1]/div/div/div/a').click()

            try:
                driver.find_element_by_xpath(
                    '/html/body/div[3]/div[2]/div[3]/div/div[1]/div[2]/h2/a').click()
                wait.until(EC.visibility_of_element_located(
                    (By.CLASS_NAME, 'item_desc')))
                item_id = driver.current_url.replace(
                    'https://item.rakuten.co.jp/', '')[:-1]
                item_image_url = top_img_url

            except NoSuchElementException:
                item_image_url = top_img_url
                item_id = nan

            except TimeoutException:
                item_image_url = top_img_url
                item_id = nan

        items_image_url.append(item_image_url)
        items_id.append(item_id)

    def postage_price(self, driver, wait, Todofuken_name):
        # ===========Seleniumで送料取得================
        # postage_price_listの対象をSeleniumでスクレイピング
        # item_idでitems_idのindexを検索し、
        # そのinndex番号に合致するitems_priceの項目を更新する
        for i, item in enumerate(postage_price_list):
            print('item')
            print(item)
            driver.set_window_size('1920', '1080')
            driver.get("https://page.auctions.yahoo.co.jp/jp/auction/" + item)
            wait.until(EC.visibility_of_element_located(
                (By.CLASS_NAME, "ProductInformation__item")))
            driver.execute_script(
                'document.querySelector("#js-prMdl-close").click()')
            postageDetailBuy_element = driver.find_element_by_id(
                "postageDetailBuy")
            try:
                postageDetailBuy_element.click()
            except ElementNotInteractableException:
                print('ElementNotInteractableException')
                continue

            select_element = driver.find_element_by_xpath(
                "//div[@class='SelectWrap']/select")
            # select_element = driver.find_elements_by_class_name("Select")
            select_object = Select(select_element)
            select_object.select_by_visible_text(Todofuken_name)
            # time.sleep(3)

            postagePrice_element = driver.find_element_by_xpath(
                "//dd[@class='BidModal__postageDetail']/div[@class='BidModal__postagePrice']")

            postagePrice = (postagePrice_element.text).split('円')[
                0].replace(',', '')
            print('postagePrice')
            print(postagePrice)
            if(postagePrice == '送料未定'):
                continue
            elif(postagePrice == '送料未定（着払い）'):
                postagePrice = 3000
            index = items_id.index(item)
            items_price[index] = str(
                int(items_price[index]) + int(postagePrice))
            eel.view_log_js(str(i + 1) + "/" +
                            str(len(postage_price_list)) + "商品目")
            # time.sleep(5)


def main():

    try:
        path = os.getcwd()
        driver_path = ChromeDriverManager(path=path).install()

        driver = set_driver(driver_path)
        driver.implicitly_wait(10)
        wait = WebDriverWait(driver, 5)

        myscraping = MyScraping()
        num_list = eel.page()().split(',')
        Todofuken_name = eel.Todofuken_name()()
        name_list, top_img_urls = myscraping.get_top_detail(
            driver, wait, eel.url()(), num_list)
        eel.view_log_js('\n商品一覧ページからの収集が完了しました')
        eel.view_log_js('\n商品の送料情報を収集します')
        myscraping.postage_price(driver, wait, Todofuken_name)
        driver.quit()

        # i = 1
        # for name, img in zip(name_list, top_img_urls):
        #     if i == 1:
        #         driver = set_driver(driver_path)
        #         driver.implicitly_wait(30)
        #         wait = WebDriverWait(driver, 10)
        #     if i % 100 == 0:
        #         driver.quit()
        #         if os.name == 'nt':  # Windows
        #             driver = set_driver(driver_path)
        #         elif os.name == 'posix':  # Mac
        #             driver = set_driver(driver_path)
        #         driver.implicitly_wait(30)
        #         wait = WebDriverWait(driver, 10)

        #     myscraping.get_item_details(driver, wait, name, img)
        #     eel.view_log_js(str(i) + "/" + str(len(name_list)) + "件目")
        #     i += 1

        eel.view_log_js('商品の送料情報収集が完了しました')
        eel.view_log_js('\nファイルの作成をします')

        print('===============================')

        result_array = {
            "M": items_condition,
            "N": items_name_en,
            "O": items_id,
            "Q": items_price,
            "W": top_img_urls,
            "X": items_Descriptions,
            "Y": items_condition_en
        }
        result_df = pd.DataFrame(result_array)
        print('result_df1')
        print(len(result_df))

        result_df.dropna(subset=['O', 'Y'], inplace=True)
        print('result_df2')
        print(len(result_df))

        wb = openpyxl.Workbook()
        ws = wb.active

        ws["M1"] = "*ConditionID"
        ws["N1"] = "Title"
        ws["O1"] = "CustomLabel"
        ws["Q1"] = "元値"
        ws["W1"] = "PicURL"
        ws["X1"] = "*Description"
        ws["Y1"] = "ConditionDescription"

        for column_name, item_list in result_df.iteritems():
            i = 2
            print('len(item_list)')
            print(len(item_list))
            for item in item_list:
                ws[column_name + str(i)] = item
                i += 1

        wb.save(eel.file_name()())

        eel.view_log_js('ファイルの作成が完了しました')

    # except WebDriverException:
    #     eel.view_log_js("一度、driversファイルとdrivers.jsonを削除してからやり直してください")

    except Exception:
        import traceback
        t = traceback.format_exc()
        eel.view_log_js(t)
        eel.view_log_js("予期せぬエラーが発生しました。")
=======
from collections import OrderedDict
import os
import eel
import pandas as pd
import re
import requests
import time
from googletrans import Translator
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException, ElementNotInteractableException
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from numpy import nan
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import subprocess
from bs4 import BeautifulSoup
from selenium.webdriver.support.select import Select

# 一覧ページから取得する情報
items_condition = []
items_name_ja = []
items_name_en = []
items_condition_en = []
items_price = []
top_img_urls = []
items_Descriptions = []
postage_price_list = []


# 各商品apiから取得する情報
items_id = []
items_image_url = []


def set_driver(driver_path):
    # Chromeドライバーの読み込み
    options = ChromeOptions()

    # 起動オプションの設定
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('log-level=3')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--incognito')          # シークレットモードの設定を付与
    options.add_argument('--user-agent=Chrome/87.0.42.88')
    options.add_argument('--single-process')
    options.add_argument('--start-maximized')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-web-security')
    options.add_argument('--disable-desktop-notifications')
    options.add_argument('--disable-application-cache')
    options.add_argument("--disable-extensions")
    options.add_argument('--lang=ja')
    # options.add_argument('--ignore-certificate-errors')
    # options.add_argument('--ignore-ssl-errors')
    # options.add_argument('--incognito')          # シークレットモードの設定を付与

    # ChromeのWebDriverオブジェクトを作成する。
    return Chrome(driver_path, options=options)


class MyScraping():

    # 状態の判定
    # 与えられた文字列＝商品の状態によって特定の文字列を返す
    def judge_condition(self, string):
        # 未使用：Condition:Mint.
        # 未使用に近い：Condition:Near Mint.
        # 目立った傷や汚れなし：Condition:Excellent.
        # やや傷や汚れあり：Condition:Very Good.
        # 傷や汚れあり：Condition:Good.
        condition = ""

        if string == "未使用":
            condition = "Condition:Mint."

        elif string == "未使用に近い":
            condition = "Condition:Near Mint."

        elif string == "目立った傷や汚れなし":
            condition = "Condition:Excellent."

        elif string == "やや傷や汚れあり":
            condition = "Condition:Very Good."

        elif string == "傷や汚れあり":
            condition = "Condition:Good."

        elif string == "中古":
            condition = "Condition : Very Good."

        else:
            condition = nan

        return condition

    # 翻訳関数
    # 与えられた日本語の文字列を英語にして返す
    def translate(self, string):
        translator = Translator()
        result = translator.translate(string, src='ja', dest='en')
        return result.text

    # APIの設定
    # 楽天API。不要？
    def get_api(self, params):
        url = 'https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706?'
        result = requests.get(url, params=params)
        return result.json()

    # APIの設定
    # 楽天API。不要？
    def set_api_parameter(self):
        app_id = 1064309786905776599
        params = {
            "format": "json",
            "applicationId": [app_id],
            "formatVersion": 2,
        }
        return params

    # 画像URLの設定
    # 複数の画像URLを一つの文字列にまとめて返す
    def image_url_add(self, urls):
        image_url = ''
        for url in urls:
            image_url += url + "|"
        return image_url[:-1]

    # 翻訳前の不要ワードの設定
    # 第一引数の文字列に対して、第二引数のリストに含まれる文言を削除する処理を行う
    def delete_word_before(self, word, delete_before_list):
        for delete_word in delete_before_list:
            word = word.replace(delete_word, "")

        return word

    # 翻訳後の不要ワードの設定
    def delete_word_after(self, word, delete_after_list, add_name_list):

        code_regex = re.compile(
            '[!"#$&*+-.:;<=>?@^_`{|}~「」〔〕“”〈〉『』＆＊・＄＃＠。、？！｀＋￥◆★◇☆⇒%％◎○■/:×,]')
        word = code_regex.sub(" ", word)

        word = re.sub(r'\[([\w\d ]+)\]', '', word)
        word = re.sub(r'\(([\w\d ]+)\)', '', word)
        word = re.sub(r'\【([\w\d ]+)\】', '', word)

        word = word.replace("[]", "")
        word = word.replace("【】", "")

        word_list = word.split(" ")

        for delete_word in delete_after_list:
            if delete_word.lower() in word_list:
                word_list.remove(delete_word.lower())
            if delete_word.capitalize() in word_list:
                word_list.remove(delete_word.capitalize())

        for add_word in add_name_list:
            if add_word.lower() in word_list:
                word_list.remove(add_word.lower())
            if add_word.capitalize() in word_list:
                word_list.remove(add_word.capitalize())

        return word_list

    # ワードの追加
    def add_word(self, word_num, word, add_word_list):

        if len(word) == word_num:
            return word

        elif len(word) > word_num:
            while len(word) + 1 >= word_num:
                word_list = word.split(' ')
                word = ' '.join(word_list[:-1])

            if len(word) == word_num:
                return word

            else:
                word = word + ' '
                for add_word in add_word_list:
                    if add_word in word:
                        continue
                    if len(word) + len(add_word) >= word_num:
                        return word
                    word = word + add_word + " "
                else:
                    return word

        elif len(word) < word_num:
            for add_word in add_word_list:
                if add_word in word:
                    continue
                if len(word) + len(add_word) >= word_num:
                    return word
                word = word + add_word + " "
            else:
                return word

    # 商品説明文字列作成
    def Description_string(self, Description):
        html_text = '<font rwr=\"1\" size=\"4\" style=\"font-family:Arial\"><div class=\"banner_topr\" style=\"height: 200px; width: 100%; background-image: url(\'https://www.saatscommerce.com/te_rsc/XOther_Japanese_c04/us/images/c_top.jpg\'); background-repeat: no-repeat; background-size: cover; \"></div><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>Description</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"></p><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><span style=\"font-size: large;\"><span style=\"font-family: &quot;Noto Sans JP&quot;, sans-serif; background-color: rgb(255, 255, 255);\">Please see the attached photographs for detailed item conditions.<br><br>●●<br><br><b>This product is 100% authentic.<br>We have several other items in stock, so please contact us.<br><br><br>[Customer\'s Feedback!]<br>\"OUTSTANDING SELLER! Fast shipping. Thanks so much for a wonderful transaction.\"<br>\"Fast delivery, perfect product. 100% satisfied. 5 star to seller. Thank you..\"<br>\"Totally worth it!!!thanks again awesome deck!!It arrived thanks!!\"<br>\"Amazing communication and ship speed. Items exactly as described. A+++ Thank u!!\"<br>\"Really beautiful transaction and super fast shipment from Japan, thank you!\"<br>\"Excellent service provided by the seller from start to end! Highly Recommended!\"<br>\"***Highly recommend to purchase from this seller*** Really friendly and reliable\"<br><br>...Thank You!</b></span><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><span style=\"font-size: large;\"><span style=\"font-family: &quot;Noto Sans JP&quot;, sans-serif; background-color: rgb(255, 255, 255);\"><br></span></span></p></h2><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>Payment</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\">We only accept PayPal payments.<br>Please pay within 5 days after auction is finished.</p><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><br></p></h2><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>Shipping</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"></p><br style=\"font-size: 14px; font-weight: 400;\"><div id=\"shipping_duration_table\" style=\"font-size: 14px; font-weight: 400;\"><b>Shipping Duration</b><table class=\"table_shipping\" style=\"width: 770px; border: 1px solid rgb(0, 0, 0); border-collapse: collapse;\"><thead><tr><th style=\"width: 145px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">Shipping<br>Method</th><th style=\"width: 145px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">Asia, Middle East</th><th style=\"width: 145px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">America,Canada<br>Australia,Mexico</th><th style=\"width: 145px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">Europe<br>Russia</th><th style=\"width: 144px; background-color: rgb(221, 221, 221); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">Other</th></tr></thead><tbody id=\"shipping_duration_tbody\"><tr><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">DHL, FedEX</td><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">2 - 3 days</td><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">3 - 5 days</td><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">3 - 4 days</td><td style=\"background-color: rgb(255, 255, 255); border: 1px solid rgb(0, 0, 0); padding: 4px; border-collapse: collapse;\">5 - 7 days</td></tr></tbody></table></div><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">w/Tracking.</span><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">Expedited Shipping.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">We will ship to world wide by EMS or FedEX or DHL.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">We will decide the shipping method.</span><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">Please contact us in advance if specified.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">It may be delayed depending on the delivery area, but we will ship as soon as possible.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">We will ship your item within 2-10 business days, after your payment.</span><br style=\"font-size: 14px; font-weight: 400;\"><br style=\"font-size: 14px; font-weight: 400;\"><span style=\"font-size: 14px; font-weight: 400;\">We can ship ONLY to your Paypal address.</span><p style=\"font-size: 14px; font-weight: 400;\"></p></h2><div><span style=\"font-size: 14px; font-weight: 400;\"><br></span></div><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>Term of Sale</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\">We can accept return within 30 days after you receive it.<br>Please make sure to contact us first if you want to return it.<br>It is buyer\'s responsibility for the return shipping cost when sending back the item to us.<br>The item has to be returned in the original condition from the time you received.<br>No returns on products that show any evidences of use.</p><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><br></p></h2><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>About Us</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\">A specialty shop offering a wide variety of carefully selected Japanese brand items.<br>Even if you think to yourself, \"But the item I want is only available in Japan.\", make sure to consult with us before giving up hope.<br>We will look for and suggest the best products on your behalf.</p><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\"><br></p></h2><h2 style=\"color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><p style=\"margin: 0px; font-size: 30px; color: rgb(139, 127, 62);\"><b>International Buyers - Please Note:</b></p></h2><h2 style=\"font-size: 14pt; color: rgb(51, 51, 51); margin: 0px; padding: 0px 0px 0px 10px;\"><hr size=\"5\" color=\"#8B7F3E\" style=\"font-size: 14px; font-weight: 400;\"><p style=\"font-weight: 400; margin: 0px; padding: 0px 30px 0px 20px; line-height: 28.8px; font-size: 16px;\">Import duties, taxes and charges are not included in the item price or shipping charges. These charges are the buyer\'\'s responsibility.<br>Please check with your country\'\'s customs office to determine what these additional costs will be prior to bidding/buying.<br>These charges are normally collected by the delivering freight (shipping) company or when you pick the item up - do not confuse them for additional shipping charges.<br><br>If your country is a non-shipping country, please contact us.<br>I will estimate the shipping cost.</p></h2></font>'
        return html_text.replace('●●', Description)

    # 一覧ページから情報を取り出す
    # 商品一覧から取り出す情報
    # ・商品名
    # ・商品の状態（中古とか可,良いとか）
    # ・値段、送料

    def get_top_detail(self, driver, wait, main_url, num_list):
        # ファイルの読み込み
        delete_before_file = pd.read_csv(
            './' + eel.delete_word_before()(), header=None, names=['削除するワード_before'])
        delete_before_list = delete_before_file['削除するワード_before'].tolist()
        delete_after_file = pd.read_csv(
            './' + eel.delete_word_after()(), header=None, names=['削除するワード_after'])
        delete_after_list = delete_after_file['削除するワード_after'].tolist()
        add_word_file = pd.read_csv(
            './' + eel.fill_in_word()(), header=None, names=['追加するワード'])
        add_word_list = add_word_file['追加するワード'].tolist()

        # 商品名の処理
        add_name = eel.add_word()()
        add_name_len = 80 - len(add_name)
        add_name_split_dot = add_name.split(',')
        add_name_split_brank = add_name.replace(",", " ").split(" ")

        # ページごと
        for i in range(int(num_list[0]), int(num_list[1]) + 1):
            eel.view_log_js("\n" + str(i) + "ページ目開始")
            page_url = main_url + "&b=" + str(i+(i*100-100))
            driver.get(page_url)
            wait.until(EC.visibility_of_element_located(
                (By.CLASS_NAME, "Products__items")))
            # items = driver.find_elements_by_class_name('Products__items')

            html = driver.page_source.encode('utf-8')

            # ＝＝＝＝＝＝＝＝＝＝追記分＝＝＝＝＝＝＝＝＝＝＝＝
            soup = BeautifulSoup(html, 'lxml')

            items = soup.find_all("li", class_="Product")
            print(len(items))

            for item in items:
                item_id = item.findAll("a")[0].get("data-auction-id")
                # print('============================================')
                # print(item_id)

            item_counter = 0
            # アイテムごと
            for j, item in enumerate(items):
                item_id = item.findAll("a")[0].get("data-auction-id")
                page_url = 'https://page.auctions.yahoo.co.jp/jp/auction/' + item_id
                driver.get(page_url)
                wait.until(EC.visibility_of_element_located(
                    (By.CLASS_NAME, 'ProductImage__inner')))
                html = driver.page_source.encode('utf-8')
                soup = BeautifulSoup(html, 'lxml')
                Price_buynow = soup.find_all("div", class_="Price--buynow")
                if (len(Price_buynow) == 0):
                    eel.view_log_js(str(j + 1) + "/" +
                                    str(len(items)) + "商品目 即決価格なし 除外")
                    continue
                print('============================================')
                print('item_id')
                print(item_id)

                # 商品のコンディションID M列
                # 商品の状態情報抽出
                item_condition_element = soup.find_all(
                    "tr", class_="ProductTable__row")
                # テキスト化
                item_condition_text = re.sub(
                    r'[\n ]', '', item_condition_element[1].findAll("a")[0].text)
                # 除外対象か判別。対象の場合飛ばして次の商品へ移動
                if item_condition_text == '全体的に状態が悪い':
                    eel.view_log_js(str(j + 1) + "/" +
                                    str(len(items)) + "商品目 状態が悪いため除外")
                    continue
                    # コンディションIDを設定
                item_condition_id = '3000'

                # 商品名 N列
                # 商品名抽出
                item_name_ja_element = soup.find_all(
                    "h1", class_="ProductTitle__text")
                item_name_ja = item_name_ja_element[0].text
                item_name_ja_delete = self.delete_word_before(
                    item_name_ja, delete_before_list)
                item_name_en = self.translate(item_name_ja_delete)
                item_name_en_delete_list = self.delete_word_after(
                    item_name_en, delete_after_list, add_name_split_brank)
                item_name_en = ' '.join(
                    OrderedDict.fromkeys(item_name_en_delete_list))
                item_name_en = self.add_word(
                    add_name_len, item_name_en, add_word_list)
                item_name_en = add_name_split_dot[0] + " " + \
                    item_name_en + " " + add_name_split_dot[1]
                item_name_en = re.sub(r'([\s]+)', ' ', item_name_en)
                item_name_en = item_name_en.title()

                # 商品識別ID O列
                # 取得済み

                # 商品の値段 Q列
                # 配送料チェック
                postage_price_free_element = soup.find_all(
                    "span", class_="Price__postageValue--free")
                # 配送料無料ではない商品をピックアップする。
                if (len(postage_price_free_element) == 0):
                    # 着払いではない商品をピックアップする。
                    postage_price_text_element = Price_buynow[0].findAll(
                        "span", class_="Price__postageValue")[0].text
                    # if(postage_price_text_element != '着払い'):
                    # 料金はpostage_priceで取得し計算する
                    postage_price_list.append(item_id)

                price_element = Price_buynow[0].findAll(
                    "dd", class_="Price__value")[0].text
                item_price = 0
                if '税込' in price_element:
                    price_element = re.split('[円税込]', price_element)[3]
                    item_price = re.sub(r'[, ]', '', price_element)
                else:
                    item_price = price_element.split('円')[0].replace(
                        ',', '').replace('\n', '')

                # 商品画像(10枚まで) W列
                item_images_element = soup.find_all(
                    "div", class_="ProductImage__inner")
                limit = 10
                urls = []
                for k, item_image in enumerate(item_images_element):
                    urls.append(item_image.findAll("img")[0].get("src"))
                    if k + 1 == limit:
                        break
                image_urls = self.image_url_add(urls)

                # 商品説明 X列
                item_Description_element = soup.find_all(
                    "div", class_="ProductExplanation__commentBody")
                item_Description = self.Description_string(
                    item_Description_element[0].text)

                # 商品の状態 Y列
                item_condition_en = self.judge_condition(
                    item_condition_text.replace('\n', ''))

                # 各データを格納
                items_condition.append(item_condition_id)
                items_name_ja.append(item_name_ja)
                items_name_en.append(item_name_en)
                items_id.append(item_id)
                items_price.append(item_price)
                top_img_urls.append(image_urls)
                items_Descriptions.append(item_Description)
                items_condition_en.append(item_condition_en)

                eel.view_log_js(str(j + 1) + "/" + str(len(items)) + "商品目")
                item_counter += 1

            else:
                eel.view_log_js(str(i) + "ページ目終了")
                eel.view_log_js(str(item_counter) + "件抽出")
                # print('items_condition')
                # print(len(items_condition))
                # # print(items_condition)

                # print('items_name_ja')
                # print(len(items_name_ja))
                # # print(items_name_ja)

                # print('items_name_en')
                # print(len(items_name_en))
                # # print(items_name_en)

                # print('items_id')
                # print(len(items_id))
                # # print(items_id)

                # print('items_price')
                # print(len(items_price))
                # # print(items_price)

                # print('top_img_urls')
                # print(len(top_img_urls))
                # # print(top_img_urls)

                # print('items_Descriptions')
                # print(len(items_Descriptions))
                # # print(items_Descriptions)

                # print('items_condition_en')
                # print(len(items_condition_en))
                # # print(items_condition_en)

        return items_name_ja, top_img_urls

    # 各商品の検索
    # 楽天apiから取るもの
    # ・商品の説明
    # ・画像URL
    # ・商品id
    def get_item_details(self, driver, wait, item_name, top_img_url):

        # APIの設定
        params = self.set_api_parameter()
        params["keyword"] = item_name

        try:
            item = self.get_api(params)["Items"]

            if item:
                item_image_url = item[0]["mediumImageUrls"]
                item_image_url = self.image_url_add(item_image_url)
                item_url = item[0]["itemUrl"]
                item_id = item_url.replace(
                    'https://item.rakuten.co.jp/', '')[:-1]

            else:
                url = 'https://www.rakuten.co.jp/'
                driver.get(url)
                wait.until(EC.visibility_of_element_located(
                    (By.CLASS_NAME, 'input-text--37z4R')))
                driver.find_element_by_class_name(
                    'input-text--37z4R').send_keys(item_name)
                driver.find_element_by_xpath(
                    '//*[@id="wrapper"]/div[5]/div/div/div/div[1]/div/div/div/a').click()

                try:
                    driver.find_element_by_xpath(
                        '/html/body/div[3]/div[2]/div[3]/div/div[1]/div[2]/h2/a').click()
                    wait.until(EC.visibility_of_element_located(
                        (By.CLASS_NAME, 'item_desc')))
                    item_id = driver.current_url.replace(
                        'https://item.rakuten.co.jp/', '')[:-1]
                    item_image_url = top_img_url

                except NoSuchElementException:
                    item_image_url = top_img_url
                    item_id = nan

                except TimeoutException:
                    item_image_url = top_img_url
                    item_id = nan

        except KeyError:
            url = 'https://www.rakuten.co.jp/'
            driver.get(url)
            wait.until(EC.visibility_of_element_located(
                (By.CLASS_NAME, 'input-text--37z4R')))
            driver.find_element_by_class_name(
                'input-text--37z4R').send_keys(item_name)
            driver.find_element_by_xpath(
                '//*[@id="wrapper"]/div[5]/div/div/div/div[1]/div/div/div/a').click()

            try:
                driver.find_element_by_xpath(
                    '/html/body/div[3]/div[2]/div[3]/div/div[1]/div[2]/h2/a').click()
                wait.until(EC.visibility_of_element_located(
                    (By.CLASS_NAME, 'item_desc')))
                item_id = driver.current_url.replace(
                    'https://item.rakuten.co.jp/', '')[:-1]
                item_image_url = top_img_url

            except NoSuchElementException:
                item_image_url = top_img_url
                item_id = nan

            except TimeoutException:
                item_image_url = top_img_url
                item_id = nan

        items_image_url.append(item_image_url)
        items_id.append(item_id)

    def postage_price(self, driver, wait, Todofuken_name):
        # ===========Seleniumで送料取得================
        # postage_price_listの対象をSeleniumでスクレイピング
        # item_idでitems_idのindexを検索し、
        # そのinndex番号に合致するitems_priceの項目を更新する
        for i, item in enumerate(postage_price_list):
            print('item')
            print(item)
            driver.set_window_size('1920', '1080')
            driver.get("https://page.auctions.yahoo.co.jp/jp/auction/" + item)
            wait.until(EC.visibility_of_element_located(
                (By.CLASS_NAME, "ProductInformation__item")))
            driver.execute_script(
                'document.querySelector("#js-prMdl-close").click()')
            postageDetailBuy_element = driver.find_element_by_id(
                "postageDetailBuy")
            try:
                postageDetailBuy_element.click()
            except ElementNotInteractableException:
                print('ElementNotInteractableException')
                continue

            select_element = driver.find_element_by_xpath(
                "//div[@class='SelectWrap']/select")
            # select_element = driver.find_elements_by_class_name("Select")
            select_object = Select(select_element)
            select_object.select_by_visible_text(Todofuken_name)
            # time.sleep(3)

            postagePrice_element = driver.find_element_by_xpath(
                "//dd[@class='BidModal__postageDetail']/div[@class='BidModal__postagePrice']")

            postagePrice = (postagePrice_element.text).split('円')[
                0].replace(',', '')
            if(postagePrice == '送料未定'):
                continue
            elif(postagePrice == '送料未定（着払い）'):
                postagePrice = 3000
            index = items_id.index(item)
            items_price[index] = str(
                int(items_price[index]) + int(postagePrice))
            eel.view_log_js(str(i + 1) + "/" +
                            str(len(postage_price_list)) + "商品目")
            # time.sleep(5)


def main():

    try:
        path = os.getcwd()
        driver_path = ChromeDriverManager(path=path).install()

        driver = set_driver(driver_path)
        driver.implicitly_wait(10)
        wait = WebDriverWait(driver, 5)

        myscraping = MyScraping()
        num_list = eel.page()().split(',')
        Todofuken_name = eel.Todofuken_name()()
        name_list, top_img_urls = myscraping.get_top_detail(
            driver, wait, eel.url()(), num_list)
        eel.view_log_js('\n商品一覧ページからの収集が完了しました')
        eel.view_log_js('\n商品の送料情報を収集します')
        myscraping.postage_price(driver, wait, Todofuken_name)
        driver.quit()

        # i = 1
        # for name, img in zip(name_list, top_img_urls):
        #     if i == 1:
        #         driver = set_driver(driver_path)
        #         driver.implicitly_wait(30)
        #         wait = WebDriverWait(driver, 10)
        #     if i % 100 == 0:
        #         driver.quit()
        #         if os.name == 'nt':  # Windows
        #             driver = set_driver(driver_path)
        #         elif os.name == 'posix':  # Mac
        #             driver = set_driver(driver_path)
        #         driver.implicitly_wait(30)
        #         wait = WebDriverWait(driver, 10)

        #     myscraping.get_item_details(driver, wait, name, img)
        #     eel.view_log_js(str(i) + "/" + str(len(name_list)) + "件目")
        #     i += 1

        eel.view_log_js('商品の送料情報収集が完了しました')
        eel.view_log_js('\nファイルの作成をします')

        print('===============================')

        result_array = {
            "M": items_condition,
            "N": items_name_en,
            "O": items_id,
            "Q": items_price,
            "W": top_img_urls,
            "X": items_Descriptions,
            "Y": items_condition_en
        }
        result_df = pd.DataFrame(result_array)
        print('result_df1')
        print(len(result_df))

        result_df.dropna(subset=['O', 'Y'], inplace=True)
        print('result_df2')
        print(len(result_df))

        wb = openpyxl.Workbook()
        ws = wb.active

        ws["M1"] = "*ConditionID"
        ws["N1"] = "Title"
        ws["O1"] = "CustomLabel"
        ws["Q1"] = "元値"
        ws["W1"] = "PicURL"
        ws["X1"] = "*Description"
        ws["Y1"] = "ConditionDescription"

        for column_name, item_list in result_df.iteritems():
            i = 2
            print('len(item_list)')
            print(len(item_list))
            for item in item_list:
                ws[column_name + str(i)] = item
                i += 1

        wb.save(eel.file_name()())

        eel.view_log_js('ファイルの作成が完了しました')

    # except WebDriverException:
    #     eel.view_log_js("一度、driversファイルとdrivers.jsonを削除してからやり直してください")

    except Exception:
        import traceback
        t = traceback.format_exc()
        eel.view_log_js(t)
        eel.view_log_js("予期せぬエラーが発生しました。")
>>>>>>> origin/master
