from bs4 import BeautifulSoup
import requests
import pandas as pd
from xlwt import Workbook 

df = pd.read_excel(".\products.xlsx",sheet_name="Sheet4",header=None)

urls = []
for row in (df.iterrows()):
    urls.append("https://www.spx.com.tr" + row[1][0])


# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet1') 
  
sheet1.write(0, 0, 'Product Name') 
sheet1.write(0, 1, 'Product Brand') 
sheet1.write(0, 2, 'Product Price') 
sheet1.write(0, 3, 'Product Code') 
sheet1.write(0, 4, 'Product Availability Percentage (%)') 



counter = 1
for url in urls:

    response = requests.get(url)
    htmlDoc = response.content
    soup = BeautifulSoup(htmlDoc, "html.parser")

    productName = soup.find("h1",attrs={"class":"product__title"})
    productBrand = soup.find("a",attrs={"class":"product__brand"})
    productPrice = soup.find("div",attrs={"class":"product__price"})
    productCode = soup.find("div",attrs={"class":"product__code d-none d-lg-block"})


    productAvailability = soup.find_all("a",attrs={"class":"d-flex align-items-center justify-content-center text-reset product__variant product__size-variant mb-3 js-variant"})
    
    productAvailabilityDisabled = soup.find_all("a",attrs={"class":"d-flex align-items-center justify-content-center text-reset product__variant product__size-variant mb-3 js-variant disabled"})


    
    if len(productAvailability) == 0:
        productAvailabilityPercentage = 0
    
    else:
        productAvailabilityPercentage = (len(productAvailability) * 100) / (len(productAvailability) + len(productAvailabilityDisabled))
    
    
    specialCondition1 = soup.find_all("a",attrs={"class":"d-flex align-items-center justify-content-center text-reset product__variant product__size-variant mb-3 js-variant selected font-weight-bold"})
    specialCondition2 = soup.find_all("a",attrs={"class":"d-flex flex-column align-items-center text-reset product__variant product__color-variant js-variant p-1 border-black selected"})
    
    if productAvailabilityPercentage == 0 and (specialCondition1 or specialCondition2):
        productAvailabilityPercentage = 100

    sheet1.write(counter, 0, productName.text.strip()) 
    sheet1.write(counter, 1, productBrand.text.strip()) 
    sheet1.write(counter, 2, productPrice.text.strip()) 
    sheet1.write(counter, 3, productCode.text.strip()) 
    sheet1.write(counter, 4, int(productAvailabilityPercentage)) 
    counter += 1

    

wb.save('productDetails.xls') 