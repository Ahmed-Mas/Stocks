'''
Created on Oct 11, 2017

@author: ahmed
'''
import requests, openpyxl
from tickers import tickers

def roundRating(num):
    return round(num * 2) / 2

wb = openpyxl.load_workbook("C:\\Users\\ahmed\\Desktop\\Book1.xlsx")
ws = wb.active

x = 3
for sym in tickers:
    print("Downloading info for " + sym + "...")
    r = requests.get("https://query2.finance.yahoo.com/v7/finance/options/" + sym + "?")
    r.raise_for_status()
    jSon = r.json()
    jSon = jSon["optionChain"]["result"][0]
    print("Oki...")


    print("Populating excel sheet...")
    ws["A" + str(x)] = jSon["underlyingSymbol"]
    ws["B" + str(x)] = jSon["quote"]["regularMarketPrice"]

    ntm = roundRating(jSon["quote"]["regularMarketPrice"])
    tempx = x
    for call in jSon["options"][0]["calls"]:
        if ntm - 2 <= call["strike"] <= ntm + 2:
            ws["C" + str(tempx)] = call["bid"]
            ws["D" + str(tempx)] = call["ask"]
            ws["E" + str(tempx)] = call["strike"]
            tempx += 1
    tempx = x
    for put in jSon["options"][0]["puts"]:
        if ntm - 2 <= put["strike"] <= ntm + 2:
            ws["F" + str(tempx)] = put["bid"]
            ws["G" + str(tempx)] = put["ask"]
            tempx += 1


    x += 10
wb.save("C:\\Users\\ahmed\\Desktop\\Book1.xlsx")
print("Done...")
