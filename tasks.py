from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
from RPA.Archive import Archive
import time,os
from robocorp.workitems import BusinessException


chromeInstance = Selenium()
chromeInstance.set_selenium_speed("0.5 Seconds")



in_strRootFolderName = os.getcwd()+"/"
in_strProcessFolderName = "RPA_Task_Level2"
in_strInputFolderPath = os.path.join(in_strRootFolderName,in_strProcessFolderName)+"/"
in_strInputPdfFolderPath = os.path.join(in_strRootFolderName,in_strProcessFolderName,"All_Orders")+"/"
in_strCsvFileName = "input.csv"
in_strInputCsvUrl="https://robotsparebinindustries.com/orders.csv"
in_strRobotSpareBinUrl = "https://robotsparebinindustries.com/#/robot-order"
int_delayShort = 1
int_delayMedium = 3
int_delayLong = 5

@task
def order_robots_from_RobotSpareBin():
    FolderValidation(in_strRootFolderName,in_strProcessFolderName)
    dt_InputCsvData = input_CSV_Downloader(in_strInputCsvUrl,in_strInputFolderPath+in_strCsvFileName,30)
    RobotSpareBinOpener(in_strRobotSpareBinUrl)
    for eachRow in dt_InputCsvData:
        strHead = eachRow["Head"]
        strBody = eachRow["Body"]
        strLegs = eachRow["Legs"]
        strOrderNumber = eachRow["Order number"]
        strAddress = eachRow["Address"]
        Form_Filler(strHead,strBody,strLegs,strAddress,strOrderNumber)
        Zip_Maker(in_strInputPdfFolderPath.rstrip("/"),in_strInputFolderPath+"Output.zip")
        


#Downloads input csv file
def input_CSV_Downloader(strInputUrl,strInputFolderPath,intDurationForDownload):
    HttpInstance = HTTP()
    HttpInstance.download(url=strInputUrl,target_file=strInputFolderPath,overwrite=True)
    if Wait_For_Download(strInputFolderPath,intDurationForDownload) :
        print("Input Csv file downloaded successfully")
        dt_Csv = Tables().read_table_from_csv(path=strInputFolderPath,header=True)
        return dt_Csv
    else:
        raise BusinessException(message="Failed to download input csv file with in the maximum duration")
    
def RobotSpareBinOpener(strSpareBinUrl):
    chromeInstance.open_available_browser(strSpareBinUrl,browser_selection="chrome",maximized=True)
    chromeInstance.maximize_browser_window()
    blnPageFound = ElementExists("//button[text()='OK']",30)
    if blnPageFound:
        chromeInstance.click_element("//button[text()='OK']")
        print("Home Page found and cliked on OK")
    else:
        raise BusinessException("Home page not for given URL")

def Form_Filler(strHead,strBody,strLegs,strShippingAddress,strOrderId):
    print("Processing order no - "+strOrderId)
    chromeInstance.select_from_list_by_index("//select[@id='head']",strHead)
    chromeInstance.click_element("//input[@id='id-body-"+strBody+"']")
    chromeInstance.input_text("//input[@placeholder='Enter the part number for the legs']",strLegs)
    chromeInstance.input_text("//input[@placeholder='Shipping address']",strShippingAddress)
    chromeInstance.click_element_when_clickable("//button[@id='preview']")
    ElementExists("//div[@id='robot-preview-image']",10)
    time.sleep(2)
    chromeInstance.wait_until_element_is_visible("//img[@alt='Head']",30)
    chromeInstance.wait_until_element_is_visible("//img[@alt='Body']",30)
    chromeInstance.wait_until_element_is_visible("//img[@alt='Legs']",30)
    blnSubmiteSuccess = False
    for i in range(10):
        chromeInstance.click_element_when_clickable("//button[@id='order']")
        blnSubmiteSuccess = ElementExists("//button[@Id='order-another']",2)
        if blnSubmiteSuccess:
            break
    if blnSubmiteSuccess:
        TakingSnap(in_strInputFolderPath+"Order ID "+strOrderId+".png")
        AddingIntoPdf(in_strInputFolderPath+"Order ID "+strOrderId+".png",in_strInputPdfFolderPath+"Order ID "+strOrderId+".pdf")
        RemoveFile(in_strInputFolderPath+"Order ID "+strOrderId+".png")
        chromeInstance.click_element_when_clickable("//button[@Id='order-another']")
        ElementExists("//button[text()='OK']")
        chromeInstance.click_element("//button[text()='OK']")
        print("Submit Success for order ID - "+strOrderId)
    else:
        raise BusinessException(message="Failed to place order after trying for 10 times")

def TakingSnap(strFilePath):
    chromeInstance.capture_page_screenshot(strFilePath)

def AddingIntoPdf(strFilePath,strOutputPath):
    pdfInstance = PDF()
    pdfInstance.add_files_to_pdf(files=[strFilePath],target_document=strOutputPath)

def RemoveFile(strFilePath):
    fs = FileSystem()
    fs.remove_file(strFilePath)
    print("Screenshot file deleted")

def Zip_Maker(strInputFolderPath,strOutputFolderPath):
    archive_Instance = Archive()
    archive_Instance.archive_folder_with_zip(strInputFolderPath,strOutputFolderPath)
    print("Turned folder into zip file successfully")

def FolderValidation(strRootFolder,strProcessFolderName):
    fs = FileSystem()
    if fs.does_directory_exist(strRootFolder+strProcessFolderName):
        fs.remove_directory(in_strRootFolderName+strProcessFolderName,True)
    fs.create_directory(in_strRootFolderName+in_strProcessFolderName)
    fs.create_directory(in_strInputPdfFolderPath)


#This function is used to wait until a file is downloaded
def Wait_For_Download(strFilePath,intTimeout):
    blnDownloadSuccess = False
    for i in range(intTimeout):
        time.sleep(int_delayShort)
        if os.path.exists(strFilePath):
            blnDownloadSuccess= True
            print("Download success")
            break
    return blnDownloadSuccess

def ElementExists(strXPath,intTimeout=20):
    try:
        chromeInstance.wait_until_element_is_visible(strXPath,intTimeout)
        return True
    except Exception as e:
        return False
    
 