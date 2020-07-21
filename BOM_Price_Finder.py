import json
import urllib
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
import xlrd
import pprint
def callback():
    #To print the File location. 
    name = askopenfilename() 
    print (name)
    
    #Opening the file  in  rb mode
    pp = pprint.PrettyPrinter(indent=4)
    workbook = xlrd.open_workbook(name,"rb")
    #getting the no of sheets
    sheets = workbook.sheet_names()
    required_data = []
    for sheet_name in sheets:
        sh = workbook.sheet_by_name(sheet_name)
        #iterate through each row to the end of the rows
        for rownum in range(sh.nrows):
            row_values = sh.row_values(rownum)
            required_data.append((row_values[2], row_values[5] ))
        #pp.pprint (required_data)
    slicedData = required_data[3:]
    #pp.pprint(slicedData)
    
    #To seperate the quantity and the mpn's
    mpn1 = []
    for item in slicedData:
        itemName = item[1]
        itemQuantity = item[0]
        mpn1.append(str(itemName))
    print("\n")
    url =""
    #creating the main loop.
    for eachMpn in mpn1:
        #creating the url with the api key 
        url = 'http://octopart.com/api/v3/parts/match?'
        url += '&queries=[{"mpn":"'+eachMpn+'"}]'
        url += '&apikey=c79a3aa9'
        #pp.pprint(str(url))
        
        #passing a web request using the url and collecting the response
        response = {}
        data = urllib.request.urlopen(url).read()
        
        response = json.loads(data)
        print("Response generated")
            #pp.pprint(response)
        for result in response['results']:
            #pp.pprint(result)
            for item in result['items']:
                print("The MPN is -----> {}".format(item["mpn"]))
                for offers in item['offers']:
                    #pp.pprint(offers['homepage_url'])
                    #pp.pprint(offers['name'])
                    pp.pprint(offers['prices'])
                       # for seller in offers[seller]:
                            #pp.pprint("Seller is {}".format(seller['homepage_url']))
                print("-----------------------------------------------------------------------------------------")
        
#if any error during the Tkinter callback to return
errmsg = 'Error!'
Button(text='File Open', command=callback).pack(fill=X)
mainloop()