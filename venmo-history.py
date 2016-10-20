# grabs html from venmo transaction history
# turns it into json
# writes that to csv


import os
import json
from bs4 import BeautifulSoup
import codecs
import xlwt

def set_Verbose():
    print("Verbose mode? y/n")
    verbose_pref = input()
    if verbose_pref == "y": 
        return True
    else:
        return False

verbose = set_Verbose()
ledger = {}
files_to_process = sorted(os.listdir())


for each_file in files_to_process:
    if ".html" not in each_file:
        continue
    if verbose: print("\n",each_file)

    html = codecs.open(each_file,'r','utf-8')
    soup = BeautifulSoup(html.read(),"html.parser")
    items = soup.find_all("div", class_="statement-item")

    for each_item in items:

        # nonpending-transactions
        #     statement-item
        #         item-details-left
        #             item-id
        #             item-date
        #         item-details-middle
        #             item-note
        #                 GRAB TEXT
        #             item-exchange
        #                 p
        #                     span .item-faded
        #                         GRAB TEXT
        #                 p
        #                     span .item-faded
        #                         GRAB TEXT
        #         item-details-right
        #             item-delta-pymt
        #                 GRAB TEXT
        #             item-source
        #                 span
        #                 funding-source-name
        #                     GRAB TEXT


        item_id = each_item.find(class_="item-id").text
        
        if verbose: print(item_id)

        item_date = each_item.find(class_="item-date").text
        item_note = each_item.find(class_="item-note").text
        
        to_and_from = each_item.find(class_="item-exchange")
        if to_and_from:
            to_and_from = to_and_from.contents
            item_from = to_and_from[0].find(class_="item-faded").text
            item_to = to_and_from[1].find(class_="item-faded").text
            if verbose: print(item_from,"to",item_to)
        else:
            item_from = ""
            item_to = ""

        item_amount = each_item.find(class_="item-delta").text
        if item_amount[0] == "+": item_amount = item_amount[1:]
        print(item_amount)

        item_source = each_item.find(class_="item-source")
        
        if item_source:
            item_source = item_source.find(class_="funding-source-name").text
        else:
            item_source = ""


        ledger[item_id] = { "Date" : item_date,
                            "Note" : item_note,
                            "From" : item_from,
                            "To" : item_to,
                            "Amount" : item_amount,
                            "Source" : item_source }

json.dump(ledger,open("ledger.json", "w"), indent=4)


# write to excel

book = xlwt.Workbook()
sheet1 = book.add_sheet("Sheet1")

cols = ["ID", "Date", "Amount", "From", "To", "Note", "Source"]
row = 0
for i in range(len(cols)):
    sheet1.write(row,i,cols[i])
row = 1

for each_item in ledger:
    # id in first column
    sheet1.write(row,0,each_item)
    
    for each_col in range(1,7):
        term = cols[each_col]
        entry = ledger[each_item][term]
        if verbose: print(entry, end=',')
        sheet1.write(row,each_col,entry)
    row +=1

book.save("venmo-history.xls")







