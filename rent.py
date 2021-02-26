import sys, os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import user
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pytesseract
import cv2, xlrd 
from PIL import Image
from array import array
import urllib.request as getImage

username = user.username
passkey = user.password


# setting tesseract app location
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'



# open an excel file
raw = xlrd.open_workbook('addresses.xlsx')

# pulling the first sheet of the file
sheet = raw.sheet_by_index(0)

# defining addresses array
addresses = sheet.col_values(0)
# calculating the number of addresses
count = len(addresses)





# selenium


PATH = "C:\Program Files (x86)\chromedriver.exe"

#driver = webdriver.Chrome(PATH)
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--force-device-scale-factor={1.5}")
driver = webdriver.Chrome(PATH, chrome_options=options)
#driver.execute_script("document.body.style.zoom='250%'")

driver.get("https://www.gosection8.com/logreg.aspx?user=landlord")



email = driver.find_element_by_id("txtUsername")
email.send_keys(username)
password = driver.find_element_by_id("txtPassword")
password.send_keys(passkey)
password.send_keys(Keys.RETURN)


driver.get("https://www.gosection8.com/ll/ll_CompareRent.aspx")

agree = driver.find_element_by_id("MainContentPlaceHolder_MainContent2_disclosureButton")

agree.send_keys(Keys.RETURN)


address = driver.find_element_by_id("MainContentPlaceHolder_MainContent2_AutocompleteAddress")
beds = driver.find_element_by_id("MainContentPlaceHolder_MainContent2_BedroomCount")


address.send_keys("8100 West Florissant Ave, Saint Louis, MO 63136")

beds.send_keys("4")
beds.send_keys(Keys.RETURN)

image = driver.find_element_by_id("MainContentPlaceHolder_MainContent2_compareRentChart")

image.screenshot("myimage.png")


oimage = cv2.imread('myimage.png')

# cimage = cv2.resize(oimage, (480,408), interpolation = cv2.INTER_AREA)

county = oimage[50:290, 120:180]
gray=cv2.cvtColor(county, cv2.COLOR_BGR2GRAY)


filename = "county.png".format(os.getpid())
cv2.imwrite(filename, gray) 
text = pytesseract.image_to_string(Image.open(filename))
print(text+"dag")



#print(len(text.partition('\n')[0]))
