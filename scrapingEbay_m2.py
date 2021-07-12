# -*- coding: utf-8 -*-
from ebaysdk.trading import Connection as Trading
from ebaysdk.exception import ConnectionError
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import gspread
from oauth2client.service_account import ServiceAccountCredentials 
import re
import time
import datetime

# 秘密鍵（JSONファイル）のファイル名を入力
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('sales-list-313306-b2642fcee6f9.json', scope)
gc = gspread.authorize(credentials)

# 「キー」でワークブックを取得
SPREADSHEET_KEY = '1Fa69Y7cYOsIrufw3GGQsR-wmXT587jQv9FSkK51pAKA'
wb = gc.open_by_key(SPREADSHEET_KEY)
ws = wb.sheet1  # 一番左の「シート1」を取得

time.sleep(1)

#アカウント情報を変数に格納
YOUR_APPID = "Tsuyoshi-ando-PRD-77a9ffde4-0c07d9d6"
YOUR_DEVID = "687a9377-cbe4-4fc7-8c56-021e8708914f"
YOUR_CERTID = "PRD-7a9ffde4b312-7c9b-4655-becc-3b16"
YOUR_AUTH_TOKEN = "AgAAAA**AQAAAA**aAAAAA**NfWuYA**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6ANmYKpC5CHpQSdj6x9nY+seQ**/EUGAA**AAMAAA**ImMyKZgYbdpUdS86NDc0TyCT6iu91cVUMLAB9U/rxeUNh2mXlU+KN70NWGOj3ZFzoI8RNrvyQqGIVAzGxkHpGw4LeH3Dcn5+1CBuwXEY16/8XqEHGuNfsQVm4sSxXTUpO09Cdlz80wr4IWM8hvraIkZiuMCqb6fJIvOSdkWQZpvfrOGQmF61P7s5Y+wUTcOfz6tG6/hUfT4Gzl8tpDFmBv28Wx36zvhIax62qEO+m/tny+4uKcXr7nBLiHkoHk0HnskLJff/Tv6TLde8VOkpKz6BFMHCNC/N5BoyBxUWDHTQ9JadedI+xw/5hmtSKQ25nmukUwplGHG29RFMwIj72GZMfbBMF8kGKo8oSBI3p3llMl7uDW05e4Q7cWk6NhrKoR41EM4HPIMmm6NN9ZuAKkPniW4zusW9EhgzWn/msI/Y079CuOGgPW0gSLZEskrqk12oqZPrGe3VXA+j7+cwwLQdsXtATZLz438bTVy6fNPxXTepBIsTO9KiNxtssLAMMuAbYcVgkwlhOvEmLCk/cBvKHxBLRb5ZS8Xk2D2fwtaUZnmpLLIOYqw8fHzjrPTzrZIqqLF2D6U6HPTeKML4MBuZZ+pzW3j8myPOsAsq86dPKIUneHEPxNa2fRx+wo7nuwfc9kLa3QMvsWWB1ASOhd1F+WOtpwHcXNCwakKmq9b1T+ahyVyCsrjp0g0SblnTSbkgSDcYIfQOSDJNito5U3DuHxBVCambXVIiEW1Ou2nh3XQxIZnGdNf/Va6I6Ge8"

time.sleep(1)

driver_path = '/Users/yuya/Downloads/chromedriver'
# driver_path = '/app/.chromedriver/bin/chromedriver'
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--remote-debugging-port=9222')


time.sleep(1)

# スプレッドシートに出力するための変数iを定義


try:
    # ebayのapiにアクセス
    api = Trading(appid=YOUR_APPID, devid=YOUR_DEVID, certid=YOUR_CERTID, token=YOUR_AUTH_TOKEN)

    while 1:
        try:
            getTo = input("収集開始日時(例2021-06-10)：")
            getFrom = input("収集終了日時(例2021-05-30)：")
            # 注文データ取得
            api.execute('GetOrders', {
                'CreateTimeFrom' : str(getFrom) + 'T00:00:00',
                'CreateTimeTo' : str(getTo) + 'T23:59:00',
            })
            time.sleep(1)
            break
        except Exception:
            print("""
------------------------------------
ERROR!!正しい日時を入力してください

------------------------------------
                """)


    item_num = len(api.response.reply.OrderArray.Order)
    ws.add_rows(item_num)
    print(str(item_num) + "個のOrderデータを収集します")

    data_num = 1

    while 1:
        flag = ws.cell(data_num, 1).value
        if not flag :
            break
        data_num += 1
    i = data_num - 1
    # 取得した注文データから各種データを取得
    for order in api.response.reply.OrderArray.Order:

        time.sleep(3)

        for txn in order.TransactionArray.Transaction:
            time.sleep(10)

            # Transaction Type : https://dev.commissionfactory.com/V1/Affiliate/Types/Transaction/
            data = {
                "created_time" : order.CreatedTime, # 売上日
                "order_No": order.OrderID, # Order No.
                "buyer_userName": order.BuyerUserID, #顧客名(名)                
                "item_No" : txn.Item.ItemID, # Item No.
                "item_name": txn.Item.Title, # Item Name
                "item_customLabel": txn.Item.SKU, # Custom Label
                "item_shippingName": order.ShippingAddress.Name, # 受取人
                "order_amountPaid": order.AmountPaid, # 売上金額
                "order_recordNumber": order.ShippingDetails.SellingManagerSalesRecordNumber # Sales record number
                }

            time.sleep(1)

            created_Time = str(data["created_time"])

            try:
                item_condition = txn.Item.ConditionID #Condition
            except:
                item_condition = 1000 #Condition

            time.sleep(1)

            if int(item_condition) == 1000:
                item_condition = "New"
            elif int(item_condition) == 1500:
                item_condition = "New other"
            elif int(item_condition) == 1750:
                item_condition = "New with defects"
            elif int(item_condition) == 2000:
                item_condition = "Certified refurbished"
            elif int(item_condition) == 2500:
                item_condition = "Seller refurbished"
            elif int(item_condition) == 2750:
                item_condition = "Like New"
            elif int(item_condition) == 3000:
                item_condition = "Used"
            elif int(item_condition) == 4000:
                item_condition = "Very Good"
            elif int(item_condition) == 5000:
                item_condition = "Good"            
            elif int(item_condition) == 6000:
                item_condition = "Acceptable"            
            else:
                item_condition = "For parts or not working"            


            if order.ShippingServiceSelected.ShippingServiceCost != "":
                order_shippingCost = order.ShippingServiceSelected.ShippingServiceCost
            else:
                order_shippingCost = 0

            ship_address1 = str(order.ShippingAddress.Street1)
            ship_address2 = str(order.ShippingAddress.Street2)
            ship_city = str(order.ShippingAddress.CityName)
            ship_state = str(order.ShippingAddress.StateOrProvince)
            ship_postal_code = str(order.ShippingAddress.PostalCode)
            ship_country = str(order.ShippingAddress.Country)
            ship_phone = str(order.ShippingAddress.Phone)

            order_shippingAddress = ship_address1 + ship_address2 + " " + ship_city + " " + ship_state + " " + ship_postal_code + " " + ship_country
            order_url = "https://www.ebay.com/sh/ord/details?srn=2168&orderid=" + data["order_No"]
            listing_url = "https://www.ebay.com/itm/" + data["item_No"]

            amazon_url = "https://www.amazon.co.jp/dp/" + data["item_customLabel"]

            # クローリング先
            urlMnsearch = "https://mnsearch.com/item?kwd=" + data["item_customLabel"]
            time.sleep(1)

            #webページ起動
            driver = webdriver.Chrome(options=options, executable_path=driver_path)
            driver.get(urlMnsearch)
            time.sleep(30)

            if item_condition == "New":

                try:
                    judgmentNew = driver.find_element_by_xpath("//section[@class='shop_list_wrapper _new_col']/table")
                    rakuten_url = judgmentNew.find_element_by_link_text("楽天市場").get_attribute("href")
                except Exception as e:
                    rakuten_url = "None"
                    print(e)

                try:
                    judgmentNew = driver.find_element_by_xpath("//section[@class='shop_list_wrapper _new_col']/table")
                    yahoo_url = judgmentNew.find_element_by_link_text("Yahooショッピング").get_attribute("href")
                except Exception as e:
                    yahoo_url = "None"
                    print(e)

                time.sleep(3)

            else:
                try:
                    judgmentUsed = driver.find_element_by_xpath("//section[@class='shop_list_wrapper _used_col ']/table")
                    rakuten_url = judgmentUsed.find_element_by_link_text("楽天市場").get_attribute("href")
                except Exception as e:
                    rakuten_url = "None"
                    print(e)

                try:
                    judgmentUsed = driver.find_element_by_xpath("//section[@class='shop_list_wrapper _used_col ']/table")
                    yahoo_url = judgmentUsed.find_element_by_link_text("Yahooショッピング").get_attribute("href")
                except Exception as e:
                    yahoo_url = "None"
                    print(e)

                time.sleep(3)

            driver.close()
            time.sleep(3)
            
            #スプレッドシートへ順に出力
            try:
                cell = ws.find(data["order_No"])
                print('既に存在します')
            except gspread.exceptions.CellNotFound:
                # print('新しく追加します')
                ws.update_cell(i+1, 1, created_Time)
                time.sleep(1)
                ws.update_cell(i+1, 2, data["order_No"])
                time.sleep(1)
                ws.update_cell(i+1, 3, order_url)
                time.sleep(1)
                ws.update_cell(i+1, 4, data["buyer_userName"])
                time.sleep(1)
                ws.update_cell(i+1, 5, data["item_No"])
                time.sleep(1)
                ws.update_cell(i+1, 6, listing_url)
                time.sleep(1)
                ws.update_cell(i+1, 7, data["item_name"])
                time.sleep(1)
                ws.update_cell(i+1, 8, item_condition)
                time.sleep(1)
                ws.update_cell(i+1, 9, data["item_customLabel"])
                time.sleep(1)
                ws.update_cell(i+1, 10, data["item_shippingName"])
                time.sleep(1)
                ws.update_cell(i+1, 11, data["order_amountPaid"].value)
                time.sleep(1)
                ws.update_cell(i+1, 12, order_shippingCost.value)
                time.sleep(1)
                ws.update_cell(i+1, 13, ship_country)
                time.sleep(1)
                ws.update_cell(i+1, 14, order_shippingAddress)
                time.sleep(1)
                ws.update_cell(i+1, 15, ship_phone)
                time.sleep(1)
                ws.update_cell(i+1, 16, re.sub(r"\D", "", data["order_recordNumber"]))
                time.sleep(1)
                ws.update_cell(i+1, 17, amazon_url)
                time.sleep(1)
                ws.update_cell(i+1, 18, rakuten_url)
                time.sleep(1)
                ws.update_cell(i+1, 19, yahoo_url)
                time.sleep(1)

                i = i+1
    i -= 1
    print(str(i) + "個のOrderデータを収集しました")


except ConnectionError as e:
    print(e)

