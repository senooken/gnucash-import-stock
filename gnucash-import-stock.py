#!/usr/bin/env python3
# coding: utf-8
## \file      gnucash-import-stock.py
## \author    SENOO, Ken
## \copyright CC0
## \date      Created: 2020-08-03T17:29+09:00

'''
# 概要
岡三オンライン証券での信用取引で発生する大量の取引を効率化するため，GnuCashへのインポート用データに変換する。

[株式約定履歴.csv] と [信用決済履歴.csv] を使う。

# 方針
1. [株式約定履歴.csv] を読み込む。
2. データ取り込み時に決済代金を計算して列に持つ。
3. 1.決済代金，2.取引区分，3.銘柄名，4.約定日，の順番にソート。
4. [信用決済履歴.csv] を読み込む。
5. 3と同じようにソート。
'''

'''
## [株式約定履歴.csv]
# 約定日	銘柄コード	銘柄名	市場	取引区分	預り	課税	約定数量	約定単価	手数料/諸経費等	税額	受渡日	受渡金額	決済損益
# 2020/07/01	1458	楽天ＥＴＦ－日経レバレッジ指数連動型	東証	現物買	特定		50	10830	632	63	2020/07/03	542195	542195

7行スキップ
使用データ
- 約定日
- 銘柄コード
- 銘柄名
- 取引区分
- 約定数量
- 約定単価
- 手数料/諸経費等
- 税額
- 受渡金額
'''

'''
## [信用決済履歴.csv]
# 取引区分	銘柄コード	銘柄名	市場区分	建区分	信用取引区分	預り	課税	新規建日	新規建単価	新規建代金	決済日	決済数量	決済単価	決済代金	約定差額	諸経費計	受渡日	受渡金額	決済損益	新規手数料	新規手数料(消費税)	決済手数料	決済手数料(消費税)	管理費	貸株料	金利	日数	逆日歩	書換料
# 信返売	8963	インヴィンシブル投資法人　投資証券	東証	買建	制度(6ヶ月)	特定	申告	2020/07/21	24260	727800	2020/07/21	30	24340	730200	2400	567	2020/07/27		1833	236	22	236	22	0	0	51	1	0	0

16行スキップ

使用データ
- 銘柄コード
- 銘柄名
- 決済日
- 決済数
- 決済単価
- 決済損益
- 決済手数料
- 決済手数料(消費税)
- 貸株料
- 金利
'''

import csv
import datetime
import random
import pathlib

## 決められたファイル名のcsvがない場合に*.csvを1個ずつ中身を確認して自動判定する処理を用意したい。
## いちいち，岡三オンライン証券からダウンロード後にファイル名の変更が手間だから。

trade_csv = '株式約定履歴.csv'
pay_csv = '信用決済履歴.csv'

## *.csvファイルから株式約定履歴と信用決済履歴を識別
for file in pathlib.Path('.').glob('*.csv'):
    with file.open(encoding='cp932') as f:
        header = f.readline().rstrip()
        if header == '株式約定履歴': trade_csv = file.name
        elif header == '信用決済履歴': pay_csv = file.name

with open(trade_csv, newline='', encoding='cp932') as trade_file, \
     open(pay_csv, newline='', encoding='cp932') as pay_file, \
     open('list.csv', 'w', newline='', encoding='cp932') as list_file, \
     open('import.csv', 'w', newline='', encoding='cp932') as import_file:
    ## 株式約定履歴.csvの取り込み
    SKIP_HEADER_LINES = 7
    for i in range(SKIP_HEADER_LINES):
        trade_file.readline()

    reader = csv.DictReader(trade_file)

    trade = []
    for row in reader:
        dic = row
        dic.update({'決済代金' : int(dic['約定数量'])*float(dic['約定単価']) \
                    , '貸株料' : 0, '金利' : 0, '決済損益': 0})
        trade.append(dic)

    ## 決済代金の計算
    # 受渡金額にすると，手数料の考慮が面倒くさいので，決済代金にする。
    ## 信用決済履歴.csvの取り込み
    SKIP_HEADER_LINES = 16
    for i in range(SKIP_HEADER_LINES):
        pay_file.readline()

    reader = csv.DictReader(pay_file)

    ## 株式約定履歴に信用決済履歴の金利・貸株料と決済損益を取り込む。
    for row_trade in trade:
        row_trade['売買代金'] = int(row_trade['決済代金']) \
            + int(row_trade['手数料/諸経費等']) + int(row_trade['税額'])
        if row_trade['取引区分'] == '現物売':
            row_trade['売買代金'] = int(row_trade['受渡金額'])
        if not row_trade['取引区分'].startswith('信返'): continue

        row_pay = next(reader)
        row_trade['貸株料'] = row_pay['貸株料']
        row_trade['金利'] = row_pay['金利']
        row_trade['決済損益'] = row_pay['決済損益']
        if row_trade['取引区分'] == '信返売':
            row_trade['売買代金'] -= \
                (int(row_trade['手数料/諸経費等']) + int(row_trade['税額']))*2 \
                + int(row_pay['金利'])
        else:
            row_trade['売買代金'] += \
                (int(row_trade['手数料/諸経費等']) + int(row_trade['税額']))*2 \
                + int(row_pay['貸株料'])

    trade = sorted(trade, key=lambda dic: (dic['約定日'], dic['取引区分'], dic['銘柄コード'] ,dic['決済代金']))

    writer = csv.DictWriter(list_file, fieldnames=trade[0].keys());
    writer.writeheader()

    writer.writerows(trade)

    ## GnuCashへの取り込み用に整形

    # Date,Transaction ID,Number,Description,Notes,Commodity/Currency,Void Reason,Action,Memo,Full Account Name,Account Name,Amount With Sym,Amount Num.,Reconcile,Reconcile Date,Rate/Price
    # 2020-07-27,3dda6469c3e48ecb278255e842cde2bf,,買付,,CURRENCY::JPY,,Buy,,個人.資産:流動資産:有価証券:信用:岡三オンライン証券:株式:9973 小僧寿し,9973 小僧寿し,"20,000 9973",20000,c,,84
    # ,,,,,,,,手数料,個人.資産:流動資産:有価証券:信用:岡三オンライン証券:株式:9973 小僧寿し,9973 小僧寿し,0 9973,0,c,,0
    # ,,,,,,,,消費税,個人.資産:流動資産:有価証券:信用:岡三オンライン証券:株式:9973 小僧寿し,9973 小僧寿し,0 9973,0,c,,0
    # ,,,,,,,,,個人.負債:流動負債:未払金:有価証券:信用:岡三オンライン証券,岡三オンライン証券,"JP¥-1,680,543",-1680543,c,,1
    # 2020-07-27,0682292e8e103fd08cf4348208c5a154,,売付,,CURRENCY::JPY,,Sell,,個人.資産:流動資産:有価証券:信用:岡三オンライン証券:株式:9973 小僧寿し,9973 小僧寿し,"10,000 9973-",-10000,c,,84
    # ,,,,,,,,,個人.費用:営業外費用:その他:支払手数料:証券会社:岡三オンライン証券,岡三オンライン証券,JP¥256,256,c,,1
    # ,,,,,,,,消費税,個人.費用:営業外費用:その他:支払手数料:証券会社:岡三オンライン証券,岡三オンライン証券,JP¥25,25,c,,1
    # ,,,,,,,,,個人.費用:営業外費用:利子割引料:岡三オンライン証券,岡三オンライン証券,JP¥60,60,c,,1
    # ,,,,,,,,,個人.資産:流動資産:未収入金:有価証券:岡三オンライン証券,岡三オンライン証券,"JP¥839,659",839659,c,,1
    # ,,,,,,,,,個人.費用:営業外費用:有価証券売却損:岡三オンライン証券,岡三オンライン証券,"JP¥10,625",10625,c,,1
    # ,,,,,,,,売買損益,個人.資産:流動資産:有価証券:信用:岡三オンライン証券:株式:9973 小僧寿し,9973 小僧寿し,0 9973,0,c,,0
    # 2020-07-27,feae2c63c7815a9c68dca581bed0603e,,売付,,CURRENCY::JPY,,Sell,,個人.資産:流動資産:有価証券:信用:岡三オンライン証券:株式:9973 小僧寿し,9973 小僧寿し,"10,000 9973-",-10000,c,,86
    # ,,,,,,,,,個人.費用:営業外費用:その他:支払手数料:証券会社:岡三オンライン証券,岡三オンライン証券,JP¥254,254,c,,1
    # ,,,,,,,,消費税,個人.費用:営業外費用:その他:支払手数料:証券会社:岡三オンライン証券,岡三オンライン証券,JP¥21,21,c,,1
    # ,,,,,,,,,個人.費用:営業外費用:利子割引料:岡三オンライン証券,岡三オンライン証券,JP¥59,59,c,,1
    # ,,,,,,,,,個人.資産:流動資産:未収入金:有価証券:岡三オンライン証券,岡三オンライン証券,"JP¥859,666",859666,c,,1
    # ,,,,,,,,売買損益,個人.資産:流動資産:有価証券:信用:岡三オンライン証券:株式:9973 小僧寿し,9973 小僧寿し,0 9973,0,c,,0
    # ,,,,,,,,,個人.収益:営業外収益:分離課税:有価証券売却益:岡三オンライン証券,岡三オンライン証券,"JP¥-19,395",-19395,c,,1

    header = 'Date,Transaction ID,Number,Description,Notes,Commodity/Currency,Void Reason,Action,Memo,Full Account Name,Account Name,Amount With Sym,Amount Num.,Reconcile,Reconcile Date,Rate/Price'.split(',')

    # BASE_ACCOUNT = '個人.資産:流動資産:有価証券:信用:岡三オンライン証券:株式:'
    # BASE_LIABILITY = '個人.負債:流動負債:未払金:有価証券:信用:岡三オンライン証券'
    BASE_ACCOUNT = '個人.資産:流動資産:有価証券:'
    BASE_LIABILITY = '個人.負債:流動負債:未払金:有価証券:'

    BASE_LOSS = '個人.費用:営業外費用:有価証券売却損:岡三オンライン証券'
    BASE_INCOME = '個人.収益:営業外収益:分離課税:有価証券売却益:岡三オンライン証券'
    BASE_FEE = '個人.費用:営業外費用:その他:支払手数料:証券会社:岡三オンライン証券'
    BASE_RATE = '個人.費用:営業外費用:利子割引料:岡三オンライン証券'
    BASE_ASSET = '個人.資産:流動資産:未収入金:有価証券'
    BASE_NATIONAL_TAX = '個人.費用:営業費用:販売費及び一般管理費:一般管理費:税金:国税:所得税:株式'
    BASE_LOCAL_TAX = '個人.費用:営業費用:販売費及び一般管理費:一般管理費:税金:地方税:道府県税:普通税:道府県民税:株式等譲渡所得割'
    BASE_BANK = '個人.資産:流動資産:その他:差入保証金:岡三オンライン証券'

    out = []

    total_real_liability = 0
    total_real_asset = 0
    total_credit_buy_liability = 0
    total_credit_buy_asset = 0
    total_credit_sell_liability = 0
    total_credit_sell_asset = 0

    for row in trade:
        if row['取引区分'] == '現物買': total_real_liability += row['売買代金']
        elif row['取引区分'] == '現物売': total_real_asset += row['売買代金']
        elif row['取引区分'] == '信新買': total_credit_buy_liability += row['売買代金']
        elif row['取引区分'] == '信返売': total_credit_buy_asset += row['売買代金']
        elif row['取引区分'] == '信新売': total_credit_sell_asset += row['売買代金']
        elif row['取引区分'] == '信返買': total_credit_sell_liability += row['売買代金']

        STOCK_NAME = row['銘柄コード'][0] + '000:' + row['銘柄コード'] + ' ' + row['銘柄名']
        dic = {}
        dic['Date'] = row['約定日']
        dic['Reconcile'] = 'c'
        dic['Commodity/Currency'] = 'CURRENCY::JPY'

        now = datetime.datetime.now()
        tid = now.strftime('%Y%m%d%H%M%S%f')  # 20桁
        tid += '{:012}'.format(int(str(random.random())[2:14])) # +12桁
        dic['Transaction ID'] = tid # 32桁

        BASE_REAL_ACCOUNT = BASE_ACCOUNT + '現物:岡三オンライン証券:株式:'
        BASE_REAL_LIABILITY = BASE_LIABILITY + '現物:岡三オンライン証券'
        BASE_REAL_ASSET = BASE_ASSET + ':現物:岡三オンライン証券';
        BASE_CREDIT_ACCOUNT = BASE_ACCOUNT + '信用:岡三オンライン証券:株式:'
        BASE_CREDIT_BUY_ASSET = BASE_ASSET + ':信返売:岡三オンライン証券'
        BASE_CREDIT_BUY_LIABILITY = BASE_LIABILITY + '信新買:岡三オンライン証券'
        BASE_CREDIT_SELL_ASSET = BASE_ASSET + ':信新売:岡三オンライン証券'
        BASE_CREDIT_SELL_LIABILITY = BASE_LIABILITY + '信返買:岡三オンライン証券'

        if row['取引区分'] == '現物売':
            # 1行目: 約定
            dic1 = dic.copy()
            dic1['Description'] = '売付'
            dic1['Full Account Name'] = BASE_REAL_ACCOUNT + STOCK_NAME
            dic1['Amount Num.'] = '-' + row['約定数量']
            dic1['Rate/Price'] = row['約定単価']
            out.append(dic1)

            # 2行目: 手数料
            dic2 = dic1.copy()
            del dic2['Date'], dic2['Description'], dic2['Commodity/Currency'], \
                dic2['Transaction ID']
            dic2['Full Account Name'] = BASE_FEE
            dic2['Amount Num.'] = row['手数料/諸経費等']
            dic2['Rate/Price'] = 1
            out.append(dic2)

            # 3行目: 消費税
            dic3 = dic2.copy()
            dic3['Memo'] = '消費税'
            dic3['Amount Num.'] = row['税額']
            out.append(dic3)

            # 4行目: 売買代金
            dic4 = dic3.copy()
            del dic4['Memo']
            dic4['Full Account Name'] = BASE_REAL_ASSET
            dic4['Amount Num.'] = row['売買代金']
            out.append(dic4)

            # 5行目: 売買損益
            dic5 = dic4.copy()
            dic5['Memo'] = '売買損益'
            dic5['Full Account Name'] = BASE_REAL_ACCOUNT + STOCK_NAME
            dic5['Amount Num.'] = row['決済損益']
            out.append(dic5)

            ## TODO: 現物の決済損益は譲渡益税.csvから読み取りが必要
            # 6行目: 損益
            dic6 = dic5.copy()
            del dic6['Memo']
            dic6['Amount Num.'] = -1 * int(row['決済損益'])

            dic6['Full Account Name'] = BASE_INCOME
            # if int(row['決済損益']) > 0:
            #     dic6['Full Account Name'] = BASE_INCOME
            # else:
            #     dic6['Full Account Name'] = BASE_LOSS
            out.append(dic6)

        elif (row['取引区分'] == '現物買') \
            or (row['取引区分'] == '信新買') or (row['取引区分'] == '信新売'):
            if row['取引区分'].startswith('現物'):
                ACCOUNT = BASE_REAL_ACCOUNT
                LIABILITY = BASE_REAL_LIABILITY
            else:
                ACCOUNT = BASE_CREDIT_ACCOUNT
                LIABILITY = BASE_CREDIT_BUY_LIABILITY

            # 1行目
            dic1 = dic.copy()
            dic1['Description'] = row['取引区分']
            dic1['Full Account Name'] = ACCOUNT + STOCK_NAME
            if row['取引区分'] == '信新売':
                dic1['Amount Num.'] = '-' + row['約定数量']
            else:
                dic1['Amount Num.'] = row['約定数量']
            dic1['Rate/Price'] = row['約定単価']
            out.append(dic1)

            # 2行目
            dic2 = dic1.copy()
            del dic2['Date'], dic2['Description'], dic2['Commodity/Currency'], \
                dic2['Transaction ID']
            # , dic['Action']
            dic2['Rate/Price'] = 1
            dic2['Memo'] = '手数料'
            dic2['Amount Num.'] = row['手数料/諸経費等']
            out.append(dic2)

            # 3行目
            dic3 = dic2.copy()
            dic3['Memo'] = '消費税'
            dic3['Amount Num.'] = row['税額']
            out.append(dic3)

            # 4行目
            dic4 = dic3.copy()
            del dic4['Memo']
            if row['取引区分'] == '信新売':
                dic4['Full Account Name'] = BASE_CREDIT_SELL_ASSET
                dic4['Amount Num.'] = str(row['売買代金'])
            else:
                dic4['Full Account Name'] = LIABILITY
                dic4['Amount Num.'] = '-' + str(row['売買代金'])
            out.append(dic4)

        elif (row['取引区分'] == '信返売') or (row['取引区分'] == '信返買'):
            # 1行目
            dic1 = dic.copy()
            dic1['Full Account Name'] = BASE_CREDIT_ACCOUNT + STOCK_NAME
            dic1['Description'] = row['取引区分']
            dic1['Rate/Price'] = row['約定単価']
            if row['取引区分'] == '信返売':
                dic1['Amount Num.'] = '-' + row['約定数量']
            else:
                dic1['Amount Num.'] = row['約定数量']
            out.append(dic1)

            # 2行目
            dic2 = dic1.copy()
            del dic2['Date'], dic2['Description'], dic2['Commodity/Currency'], \
                dic2['Transaction ID']
            dic2['Full Account Name'] = BASE_FEE
            dic2['Amount Num.'] = row['手数料/諸経費等']
            dic2['Rate/Price'] = 1
            out.append(dic2)

            # 3行目
            dic3 = dic2.copy()
            dic3['Memo'] = '消費税'
            dic3['Amount Num.'] = row['税額']
            out.append(dic3)

            # 4行目
            dic4 = dic3.copy()
            del dic4['Memo']
            if row['取引区分'] == '信返売':
                dic4['Full Account Name'] = BASE_RATE + ":金利"
                dic4['Amount Num.'] = row['金利']
            else:
                dic4['Full Account Name'] = BASE_RATE + ":貸株料"
                dic4['Amount Num.'] = row['貸株料']
            out.append(dic4)

            # 5行目
            dic5 = dic4.copy()
            if row['取引区分'] == '信返売':
                dic5['Full Account Name'] = BASE_CREDIT_BUY_ASSET
                dic5['Amount Num.'] = row['売買代金']
            else:
                dic5['Full Account Name'] = BASE_CREDIT_SELL_LIABILITY
                dic5['Amount Num.'] = -row['売買代金']
            out.append(dic5)

            # 6行目
            dic6 = dic5.copy()
            dic6['Memo'] = '売買損益'
            dic6['Full Account Name'] = BASE_CREDIT_ACCOUNT + STOCK_NAME
            dic6['Amount Num.'] = row['決済損益']
            out.append(dic6)

            # 7行目
            dic7 = dic6.copy()
            del dic7['Memo']
            dic7['Amount Num.'] = -1 * int(row['決済損益'])
            if int(row['決済損益']) > 0:
                dic7['Full Account Name'] = BASE_INCOME
            else:
                dic7['Full Account Name'] = BASE_LOSS
            out.append(dic7)

    ## 精算
    pay_date = datetime.datetime.strptime(row['約定日'], '%Y/%m/%d')
    pay_date += datetime.timedelta(days=2)
    if (pay_date.isoweekday() > 5):
        pay_date += datetime.timedelta(days=2)

    dic = {}
    dic['Date'] = pay_date.strftime('%Y/%m/%d')
    dic['Reconcile'] = 'c'
    dic['Commodity/Currency'] = 'CURRENCY::JPY'

    now = datetime.datetime.now()
    tid = now.strftime('%Y%m%d%H%M%S%f')  # 20桁
    tid += '{:012}'.format(int(str(random.random())[2:14])) # +12桁
    dic['Transaction ID'] = tid # 32桁

    # 1行目
    dic1 = dic.copy()
    dic1['Description'] = '精算'
    dic1['Full Account Name'] = BASE_BANK
    dic1['Rate/Price'] = 1
    dic1['Amount Num.'] = total_real_asset - total_real_liability \
        + total_credit_buy_asset - total_credit_buy_liability \
        + total_credit_sell_asset - total_credit_sell_liability
    out.append(dic1)

    # 2行目: 譲渡益税徴収額 (所得税)
    dic2 = dic1.copy()
    del dic2['Date'], dic2['Description'], dic2['Commodity/Currency'], \
        dic2['Amount Num.'], dic2['Transaction ID']
    dic2['Memo'] = '譲渡益税徴収額'
    dic2['Full Account Name'] = BASE_NATIONAL_TAX
    out.append(dic2)

    # 3行目: 譲渡益税徴収額 (地方税)
    dic3 = dic2.copy()
    dic3['Memo'] = '譲渡益税徴収額'
    dic3['Full Account Name'] = BASE_LOCAL_TAX
    out.append(dic3)

    # 4行目: 現物買
    dic4 = dic3.copy()
    dic4['Memo'] = '現物買'
    dic4['Full Account Name'] = BASE_REAL_LIABILITY
    dic4['Amount Num.'] = total_real_liability
    out.append(dic4)

    # 5行目: 現物売
    dic5 = dic4.copy()
    dic5['Memo'] = '現物売'
    dic5['Full Account Name'] = BASE_REAL_ASSET
    dic5['Amount Num.'] = -total_real_asset
    out.append(dic5)

    # 6行目: 信新買
    dic6 = dic5.copy()
    dic6['Memo'] = '信新買'
    dic6['Full Account Name'] = BASE_CREDIT_BUY_LIABILITY
    dic6['Amount Num.'] = total_credit_buy_liability
    out.append(dic6)

    # 7行目: 信用売
    dic7 = dic6.copy()
    dic7['Memo'] = '信返売'
    dic7['Full Account Name'] = BASE_CREDIT_BUY_ASSET
    dic7['Amount Num.'] = -total_credit_buy_asset
    out.append(dic7)

    # 8行目: 信新売
    dic8 = dic7.copy()
    dic8['Memo'] = '信新売'
    dic8['Full Account Name'] = BASE_CREDIT_SELL_ASSET
    dic8['Amount Num.'] = -total_credit_sell_asset
    out.append(dic8)

    # 9行目: 信返買
    dic9 = dic6.copy()
    dic9['Memo'] = '信返買'
    dic9['Full Account Name'] = BASE_CREDIT_SELL_LIABILITY
    dic9['Amount Num.'] = total_credit_sell_liability
    out.append(dic9)
    '''
    必須
    - Commodity/Currency: これがないと通貨単位が変わる。
    - Transaction ID: これがないと，複数の取引が同一取引とみなされまとめられる。
    feae2c63c7815a9c68dca581bed0603e


    有価証券の仕訳の場合，Ammount Num.は必須のようだ。代わりに，Rate/Priceは0でも問題なかった。

    '''

    writer = csv.DictWriter(import_file, fieldnames=header, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(out)
