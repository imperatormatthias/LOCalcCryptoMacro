# Crypto Macro for LibreOffice Calc
import json
import urllib.request

oDoc = XSCRIPTCONTEXT.getDocument()

with urllib.request.urlopen('https://api.coinmarketcap.com/v1/ticker/') as req:
    req = req.read()

response = json.loads(req.decode())

def getPrices():
    oSheet = oDoc.CurrentController.ActiveSheet
    for i in range(1,20):
        oCell1 = oSheet.getCellByPosition(0,i)

        for j in range(len(response)):
            if oCell1.String == response[j]['symbol']:
                oCell2 = oSheet.getCellByPosition(1,i)
                oCell2.Value= response[j]['price_usd']
