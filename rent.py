import sys, os
import processImg
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import user
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
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

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--force-device-scale-factor={1.5}")
driver = webdriver.Chrome(PATH, options=options)
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


address.send_keys("5574 Waterman Blvd, Saint Louis, MO 63112")

beds.send_keys("4")
beds.send_keys(Keys.RETURN)

image = driver.find_element_by_id("MainContentPlaceHolder_MainContent2_compareRentChart")


image.screenshot("myimage.png")





# county = cv2.imread('myimage.png')
# gray= processImg.get_grayscale(county)
# thresh = processImg.thresholding(gray)
# filename = "processed.jpg".format(os.getpid())
# cv2.imwrite(filename, thresh) 
# targetImg = Image.open(filename)
# targetImg.save('withdpi.jpg', dpi=(300,300))



# finalImage = cv2.imread('withdpi.jpg')


# text = pytesseract.image_to_string(finalImage, config='--psm 6')
# print(text+"dag")


