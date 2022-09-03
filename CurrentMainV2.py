# Main Code - Image Processing + Excel Data Output
###############################################
# Imports
from numpy import negative
import pyautogui
import time
import os
import pytesseract
from openpyxl import load_workbook
import mouse
import logging
import cv2
import pygetwindow as gw
import ConfigHandler as Config
import keyboard
##################################
# Declarations
SheetDict = Config.ConfigRead('MAIN','SheetDict','dict')
ScreenRegionDict = Config.ConfigRead('MAIN','ScreenRegionDict','dict')
DEBUG=1 # For my personal use , maybe cleaned up after ....
OnlyRootDebug = Config.ConfigRead('MAIN','OnlyRootDebug','bool') #Used to disable all libs from using logging :)
pytesseract.pytesseract.tesseract_cmd = Config.ConfigRead('MAIN','tess_cmd') #Tesseract .exe bcs i cant use PATH on admin locked work device :/ need to maybe auto location
tessdefault_config = Config.ConfigRead('MAIN','tessdefault_config')
tessnumber_config = Config.ConfigRead('MAIN','tessnumber_config')
logging.getLogger().setLevel(logging.DEBUG) #Need to add to GUI as check + need to clean/add logging
URLList = Config.ConfigRead('MAIN','URLList','list')
#URLDict #Need to add some sort of autologin at start of the shift
DeleteImagesAfterUsage = Config.ConfigRead('MAIN','DeleteImagesAfterUsage','bool') #Auto-cleaning of all images , maybe i should add all unnecessary files (need to be true for sure when release)
LoadCheckDict = Config.ConfigRead('MAIN','LoadCheck','dict')
###############################################
# Classes and Definitions
def SheetDictData(x,ShiftCheck): #Assign proper cell number depending on hour of data collection , something is broken RN ! 
        match x:        #Maybe move to time table in time loop ?
            case 7 | 19 if ShiftCheck == 12: return 2;
            case 8 | 20 if ShiftCheck == 12: return 3;
            case 9 | 21 if ShiftCheck == 12: return 4;
            case 10| 22 if ShiftCheck == 12: return 5;
            case 11| 23 if ShiftCheck == 12: return 6;
            case 12| 24 if ShiftCheck == 12: return 7;
            case 13| 1  if ShiftCheck == 12: return 8;
            case 14| 2  if ShiftCheck == 12: return 9;
            case 15| 3  if ShiftCheck == 12: return 10;
            case 16| 4  if ShiftCheck == 12: return 11;
            case 17| 5  if ShiftCheck == 12: return 12;
            case 18| 6  if ShiftCheck == 12: return 13;
            case 7 | 15 | 23 if ShiftCheck == 8: return 2;
            case 8 | 16 | 24 if ShiftCheck == 8: return 3;
            case 9 | 17 | 1  if ShiftCheck == 8: return 4;
            case 10| 18 | 2  if ShiftCheck == 8: return 5;
            case 11| 19 | 3  if ShiftCheck == 8: return 6;
            case 12| 20 | 4  if ShiftCheck == 8: return 7;
            case 13| 21 | 5  if ShiftCheck == 8: return 8;
            case 14| 22 | 6  if ShiftCheck == 8: return 9;
            case _ : logging.critical("Failed to allocate data to sheet disc");return 15,    

def ScreenshotRegion(name,Left,Top,Width,Height): #Making screenshot , not really needed but code looks cleaner
    logging.debug("Creating screenshot : "+name+".png")
    pyautogui.screenshot(str(name+".png"),region=(Left,Top,Width,Height))

def LoadCheck(): #Checking if webpage is loaded fully...
    LoadGood=False
    ScreenshotRegion("LoadCheck",LoadCheckDict["LoadCheck"][0],LoadCheckDict["LoadCheck"][1],LoadCheckDict["LoadCheck"][2],LoadCheckDict["LoadCheck"][3])
    ScreenshotRegion("ScrapLoadCheck",LoadCheckDict["ScrapLoadCheck"][0],LoadCheckDict["ScrapLoadCheck"][1],LoadCheckDict["ScrapLoadCheck"][2],LoadCheckDict["ScrapLoadCheck"][3]) #Using 2 separate ways of comparing images is not best but LoadCheck wasnt working with cv2.norm somehow
    b1, g1, r1 = cv2.split(cv2.subtract(cv2.imread("LoadCheck.png"),cv2.imread("LoadGood.png"))) # This is good only if both things are vibrant colors
    ImgDiffGreen = 1 - cv2.norm(cv2.imread("ScrapLoadCheck.png"),cv2.imread("ScrapLoadGoodGreen.png"), cv2.NORM_L2 ) / ( 70 * 70 )
    ImgDiffRed = 1 - cv2.norm(cv2.imread("ScrapLoadCheck.png"),cv2.imread("ScrapLoadGoodRed.png"), cv2.NORM_L2 ) / ( 70 * 70 )
    ImgDiffWhite = 1 - cv2.norm(cv2.imread("ScrapLoadCheck.png"),cv2.imread("ScrapLoadGoodWhite.png"), cv2.NORM_L2 ) / ( 70 * 70 )
    logging.debug('ImgDiffGreen similarity = '+str(ImgDiffGreen)) 
    logging.debug('ImgDiffRed similarity = '+str(ImgDiffRed))
    logging.debug('ImgDiffWhite similarity = '+str(ImgDiffWhite)) #WHITE IS NOT USED ANYWHERE AT THE MOMENT
    if cv2.countNonZero(b1) == 0  and cv2.countNonZero(g1) == 0 and cv2.countNonZero(r1) == 0:
        LoadGood=True
    else:
        return 0
    if LoadGood==True and ImgDiffGreen>0.40 or LoadGood==True and ImgDiffRed>0.40: #0.40 is for moused grayness , 0.9 is normal
        DeleteFile("LoadCheck.png")
        DeleteFile("ScrapLoadCheck.png")
        logging.info("Loading of page completed...")
        return 1 #Basically not needed because when def ends here it will return None which is not 0 so loop continues but its nice way of programming
    else:
        logging.info('Loading not finished , waiting...')
        return 0; #replacing semicolon breaks whole cycling :)

def ImageProcess(IMGName,range2=-1,range=0,tescfg=tessnumber_config): #Image Data Extraction
    try:
        img = cv2.imread(str(IMGName)+'.png')
        img = cv2.resize(img, None, fx=2, fy=2,interpolation=cv2.INTER_CUBIC)
        value=pytesseract.image_to_string(img,config=tescfg)[range:range2]    
    except:
        logging.debug(IMGName+" error!")
        if DEBUG==1 :
            logging.exception('')
        return "error"
    else:
        if DEBUG==1:logging.debug(IMGName+" processed...")
        return value

#print ([k for k in logging.Logger.manager.loggerDict]) #### totally not from stackoverflow :)

def WindowFullScreen():
    time.sleep(1) #Maybe not needed , i just want to be sure to get proper window selection
    window=gw.getActiveWindow()   #Forcing window maximizing
    if window.isMaximized == True: 
        logging.debug("Window is maximized")
    else :
        window.maximize()
        logging.debug("Maximizing window")

def DeleteFile(value):
    if DeleteImagesAfterUsage == True: #Extremely lazy cleaning ....
        try:
            os.remove(value)
        except:
            pass
        else:
            pass

def ImageGrab(MachineName,HourValue,MachineURL):
    logging.getLogger('PIL.Image').disabled = OnlyRootDebug
    logging.getLogger('PIL').disabled = OnlyRootDebug
    logging.getLogger('PIL.TiffImagePlugin').disabled = OnlyRootDebug
    logging.getLogger('PIL.PngImagePlugin').disabled = OnlyRootDebug ### Some loggers get loaded during using of sublibs ....

    os.system("start chrome --start-maximized --new-window "+URLList[MachineURL]) #Start Browser (need to add auto login - most likely just mouse clicks)
    
    WindowFullScreen()
    
    LoadTimer=0
    LoadRefresh=1
    while LoadCheck()==0: #not great way to check i think but works
        LoadTimer=LoadTimer+1
        if LoadTimer>=15 and LoadRefresh==1:
            keyboard.press_and_release('F5')
            logging.info('Refreshing page')
            LoadRefresh=0
        if LoadTimer>=35:
            logging.info('Skipping loading...')
        time.sleep(1)
        logging.debug('LoadTimer= '+str(LoadTimer))
        logging.info("Waiting for page load")
    
    for each in ScreenRegionDict:
        logging.debug("Image Creation: "+each)
        x=ScreenRegionDict[each] #Redundant line but looks better ? + less listing through dict ? performance gains ?
        ScreenshotRegion(each+MachineName+str(HourValue),x[0],x[1],x[2],x[3])

    MouseCur=mouse.get_position() #Close Window (not pretty simualing mouse clicks but i dont want to mess with os.kill and os.pid)
    mouse.move(10,10)
    mouse.click("middle")
    mouse.move(MouseCur[0],MouseCur[1])

def ExcelOutput(MachineName,HourValue,ShiftCheck,PathToFile):
    logging.info("Applying data to excel")  #Excel Start
    wb = load_workbook(filename="AutoData_gen.xlsx") #NEED TO ADD TO CONFIG LATER
    sheet=wb["AutoData-"+MachineName]
    
    # !!MESSY!!
    SheetDictDataAnswer=SheetDictData(int(HourValue),ShiftCheck) #Getting info from main thread
    for each in SheetDict: #Editing cell location 
        x = SheetDict[each][0][0:1]
        x = x+str(SheetDictDataAnswer)
        SheetDict[each][0] = str(x)
        match each: #Editing cell value
            case "TimeFlag" : SheetDict[each][1] = str(SheetDictDataAnswer-1) #-1 = Offset used in excel for header
            case "TimeSend" : SheetDict[each][1] = time.strftime("%H:%M:%S",time.localtime())
            case "OEE" | "Scrap" : SheetDict[each][1] = ImageProcess(each+MachineName+str(HourValue),-1) #NOT TRUE , NEED TO MONITOR DATA -> Extra removing % , it safer to detect and then remove than to block with blacklist
            case "Product" : SheetDict[each][1] = ImageProcess(each+MachineName+str(HourValue),-5,tescfg='--psm 7 --oem 3') #Should be good , need more data
            case "OK" : #Hardcoded cleaning of value + Maybe still broken + Not pretty at all
                x = ImageProcess(each+MachineName+str(HourValue),4)
                if x=="error" or x=="": #Not pretty handling
                    SheetDict[each][1] = "error"
                    break  
                while x.isdigit() == False : #Still too many infinite loops
                    logging.debug("OK before:"+x)
                    y=negative(len(x)-1) 
                    #x = x[y:]
                    x = x[:-y]
                    logging.debug("OK after :"+x)
                    time.sleep(0.5)   #Maybe still broken , need to log data
                SheetDict[each][1] = x
            case 'NOK': # NEED TO FINISH , BAD REPORTING OF 0,1 
                SheetDict[each][1] = ImageProcess(each+MachineName+str(HourValue),-1,tescfg="--psm 7 --oem 3") #Need to log data
            case _ : SheetDict[each][1] = ImageProcess(each+MachineName+str(HourValue),-1)

        logging.debug(str(each)+" : "+str(SheetDict[each][0]+' : '+str(SheetDict[each][1])))
        sheet[str(SheetDict[each][0])]=str(SheetDict[each][1])   #Writing into excel
        DeleteFile(str(each+MachineName+str(HourValue))+".png")
    # !!MESSY!!

    try: #Saving Excel Document #Very barebones 
        wb.save(filename="AutoData_gen.xlsx")  #NEED TO ADD TO CONFIG LATER
    except:
        logging.warning("Excel edit failed") 
    else:
        logging.info('Excel edit success')
        os.startfile('AutoData_gen.xlsx') #NOT PRETTY WAY TO UPDATE DATA
        time.sleep(3)
        WindowFullScreen()
        time.sleep(1)
        MouseCur=mouse.get_position() #Close Window (os.kill dont work without admin privileges)
        mouse.move(1900,10)
        mouse.click("left")
        mouse.move(MouseCur[0],MouseCur[1])

