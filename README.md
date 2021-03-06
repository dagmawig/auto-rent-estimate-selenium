# auto rent estimate with selenium

## Using python I automated rent estimate search on GoSection8.com website using street address
 ### The search is made for 4 differen bedroom sizes for each street address.

## I used selenium module to capture rent estimate chart and cropped image into smaller images to be fed to Optical Character Recognition (OCR) website
 ### Each rent estimate search gives a chart that contains rent estimate for 4 conditions (County, City, Zip and Radius)
 ### I cropped each chart into 4 different images where each image contains rent estimate data for one of the above conditions.

## I fed the cropped images to the OCR website and retrieved the rent estimate data and fed the data to an excel file.


### Captured chart images are found in charts folder
### street address data and the rent estimate data are found in addresses.xlsx file.