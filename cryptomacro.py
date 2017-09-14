# Crypto Macro for LibreOffice Calc
import json
import urllib.request

oDoc = XSCRIPTCONTEXT.getDocument()

coins = ['bitcoin', 'ethereum', 'ripple']

with urllib.request.urlopen('https://api.coinmarketcap.com/v1/ticker/') as req:
    req = req.read()

response = json.loads(req.decode())

def getPrices():
    for i in range(len(response)):
        for j in range(len(coins)):
            if response[i]['id'] == coins[j]:
                oSheet = oDoc.CurrentController.ActiveSheet
                oCell1 = oSheet.getCellByPosition(0,j+1)
                oCell1.String = response[i]['symbol']
                oCell2 = oSheet.getCellByPosition(1,j+1)
                oCell2.Value = response[i]['price_usd']
