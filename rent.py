import sys, os, time
import processImg
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import user
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import pytesseract
import cv2, xlrd 
from PIL import Image, ImageOps
from array import array
import urllib.request as getImage

username = user.username
passkey = user.password


# setting tesseract app location
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'



# open an excel file
raw = xlrd.open_workbook('addresses1.xlsx')

# pulling the first sheet of the file
sheet = raw.sheet_by_index(0)

# calculating the number of addresses
count = len(sheet.col_values(0))-1

addresses = []

for i in range(count):
    addresses.append( str(int(sheet.cell(i+1, 0).value)) + " " + str(sheet.cell(i+1, 1).value) + " " + str(sheet.cell(i+1, 2).value) + "," + " " + str(sheet.cell(i+1, 3).value) + "," + " " + str(sheet.cell(i+1, 4).value) + " " + str(int(sheet.cell(i+1, 5).value)) )

# selenium


PATH = "C:\Program Files (x86)\chromedriver.exe"

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--force-device-scale-factor={1.1}")


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

size = (450, 360)
size1 = (85,55,185,300)
size2 = (185,55,255,300)
size3 = (255,55,330,300)
size4 = (330,55,430,300)
imageList = []

for strAddress in addresses:
    for bed in range(4):
        address = driver.find_element_by_id("MainContentPlaceHolder_MainContent2_AutocompleteAddress")
        beds = driver.find_element_by_id("MainContentPlaceHolder_MainContent2_BedroomCount")
        address.clear()
        beds.clear()
        address.send_keys(strAddress)
        beds.send_keys(bed+1)
        beds.send_keys(Keys.RETURN)  
        image = driver.find_element_by_id("MainContentPlaceHolder_MainContent2_compareRentChart")
        image.screenshot("./charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+".png")
        openImg = Image.open("./charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+".png")
        resizeImg = ImageOps.fit(openImg, size, Image.ANTIALIAS)
        cropImg1 = resizeImg.crop(size1)
        cropImg2 = resizeImg.crop(size2)
        cropImg3 = resizeImg.crop(size3)
        cropImg4 = resizeImg.crop(size4)

        cropImg1.save("./charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+"-1"+".png")
        imageList.append("/charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+"-1"+".png")

        cropImg2.save("./charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+"-2"+".png")
        imageList.append("/charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+"-2"+".png")

        cropImg3.save("./charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+"-3"+".png")
        imageList.append("/charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+"-3"+".png")

        cropImg4.save("./charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+"-4"+".png")
        imageList.append("/charts/"+strAddress.replace(" ", "").replace(",", "")+"-"+str(bed+1)+"-4"+".png")




# driver.get("https://brandfolder.com/workbench/extract-text-from-image")
driverOCR = []

for i in range(16):
    driverOCR.append(webdriver.Chrome(PATH, options=options))

for driverSingle in driverOCR:
    driverSingle.get("https://brandfolder.com/workbench/extract-text-from-image")


print(driverOCR)

rentEst = []

for j in range(count):
    time.sleep(5)
    upload = []
    for k in range(16):
        upload.append(driverOCR[k].find_element_by_class_name("fsp-drop-pane__input"))
        upload[k].clear()
        upload[k].send_keys(os.getcwd() + imageList[16*j + k])

    time.sleep(5)

    for l in range(16):

        text = driverOCR[l].find_element_by_id("extracted_text").text

        while(len(text) == 0):
            uploadS = driverOCR[l].find_element_by_class_name("fsp-drop-pane__input")
            uploadS.clear()
            uploadS.send_keys(os.getcwd() + imageList[16*j + l])
            time.sleep(5)

        rentEst.append(driverOCR[l].find_element_by_id("extracted_text").text)

        driverOCR[l].refresh()


print(rentEst)




# upload = driver.find_element_by_class_name("fsp-drop-pane__input")
# upload.send_keys(os.getcwd() + "/a.png")
# time.sleep(5)
# textEl = driver.find_element_by_id("extracted_text")
# text = textEl.text
# print(text)




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


