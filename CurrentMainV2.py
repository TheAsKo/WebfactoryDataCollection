# Main Code - Image Processing + Excel Data Output
##############################################################################################
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
##############################################################################################
# Declarations
SheetDict = Config.ConfigRead('MAIN','SheetDict','dict')
ScreenRegionDict = Config.ConfigRead('MAIN','ScreenRegionDict','dict')
TimeDEBUG = Config.ConfigRead('MAIN','TimeDebug','bool')  # For my personal use , maybe cleaned up after ....
OnlyRootDebug = Config.ConfigRead('MAIN','OnlyRootDebug','bool') #Used to disable all libs from using logging :)
pytesseract.pytesseract.tesseract_cmd = Config.ConfigRead('MAIN','tess_cmd') #Tesseract .exe bcs i cant use PATH on admin locked work device :/ need to maybe auto location
tessdefault_config = Config.ConfigRead('MAIN','tessdefault_config')
tessnumber_config = Config.ConfigRead('MAIN','tessnumber_config')
tessadvnumber_config = Config.ConfigRead('MAIN','tessadvnumber_config')
logging.getLogger().setLevel(logging.DEBUG) #Need to add to GUI as check + need to clean/add logging
URLList = Config.ConfigRead('MAIN','URLList','list')
#URLDict #Need to add some sort of autologin at start of the shift
DeleteImagesAfterUsage = Config.ConfigRead('MAIN','DeleteImagesAfterUsage','bool') #Auto-cleaning of all images , maybe i should add all unnecessary files (need to be true for sure when release)
LoadCheckDict = Config.ConfigRead('MAIN','LoadCheck','dict')
###
logging.getLogger('PIL.Image').disabled = OnlyRootDebug
logging.getLogger('PIL').disabled = OnlyRootDebug
logging.getLogger('PIL.TiffImagePlugin').disabled = OnlyRootDebug
logging.getLogger('PIL.PngImagePlugin').disabled = OnlyRootDebug 
### Some loggers get loaded during using of sublibs ....
##############################################################################################
# Classes and Definitions
def SheetDictData(x,ShiftCheck): #Assign proper cell number depending on hour of data collection , something is broken RN ! 
        match x:        #Maybe move to time table in time loop ?
            case 7 | 19 if ShiftCheck == 12: return 2
            case 8 | 20 if ShiftCheck == 12: return 3
            case 9 | 21 if ShiftCheck == 12: return 4
            case 10| 22 if ShiftCheck == 12: return 5
            case 11| 23 if ShiftCheck == 12: return 6
            case 12| 24 if ShiftCheck == 12: return 7
            case 13| 1  if ShiftCheck == 12: return 8
            case 14| 2  if ShiftCheck == 12: return 9
            case 15| 3  if ShiftCheck == 12: return 10
            case 16| 4  if ShiftCheck == 12: return 11
            case 17| 5  if ShiftCheck == 12: return 12
            case 18| 6  if ShiftCheck == 12: return 13
            case 7 | 15 | 23 if ShiftCheck == 8: return 2
            case 8 | 16 | 24 if ShiftCheck == 8: return 3
            case 9 | 17 | 1  if ShiftCheck == 8: return 4
            case 10| 18 | 2  if ShiftCheck == 8: return 5
            case 11| 19 | 3  if ShiftCheck == 8: return 6
            case 12| 20 | 4  if ShiftCheck == 8: return 7
            case 13| 21 | 5  if ShiftCheck == 8: return 8
            case 14| 22 | 6  if ShiftCheck == 8: return 9
            case _ : logging.critical("Failed to allocate data to sheet disc");return 15    
##############################################################################################
def ScreenshotRegion(name,Left,Top,Width,Height): #Making screenshot , not really needed but code looks cleaner
    log = logging.getLogger('ScreenshotRegion')
    log.debug("Creating screenshot : "+name+".png")
    pyautogui.screenshot(str(name+".png"),region=(Left,Top,Width,Height))
##############################################################################################
def LoadCheck(): #Checking if webpage is loaded fully...
    log = logging.getLogger('LoadCheck')
    LoadGood=False
    ScreenshotRegion("LoadCheck",LoadCheckDict["LoadCheck"][0],LoadCheckDict["LoadCheck"][1],LoadCheckDict["LoadCheck"][2],LoadCheckDict["LoadCheck"][3])
    ScreenshotRegion("ScrapLoadCheck",LoadCheckDict["ScrapLoadCheck"][0],LoadCheckDict["ScrapLoadCheck"][1],LoadCheckDict["ScrapLoadCheck"][2],LoadCheckDict["ScrapLoadCheck"][3]) #Using 2 separate ways of comparing images is not best but LoadCheck wasnt working with cv2.norm somehow
    b1, g1, r1 = cv2.split(cv2.subtract(cv2.imread("LoadCheck.png"),cv2.imread("LoadGood.png"))) # This is good only if both things are vibrant colors
    ImgDiffGreen = 1 - cv2.norm(cv2.imread("ScrapLoadCheck.png"),cv2.imread("ScrapLoadGoodGreen.png"), cv2.NORM_L2 ) / ( 70 * 70 )
    ImgDiffRed = 1 - cv2.norm(cv2.imread("ScrapLoadCheck.png"),cv2.imread("ScrapLoadGoodRed.png"), cv2.NORM_L2 ) / ( 70 * 70 )
    ImgDiffWhite = 1 - cv2.norm(cv2.imread("ScrapLoadCheck.png"),cv2.imread("ScrapLoadGoodWhite.png"), cv2.NORM_L2 ) / ( 70 * 70 )
    log.debug('ImgDiffGreen similarity = '+str(ImgDiffGreen)) 
    log.debug('ImgDiffRed similarity = '+str(ImgDiffRed))
    log.debug('ImgDiffWhite similarity = '+str(ImgDiffWhite)) #WHITE IS NOT USED ANYWHERE AT THE MOMENT
    if cv2.countNonZero(b1) == 0  and cv2.countNonZero(g1) == 0 and cv2.countNonZero(r1) == 0:
        LoadGood=True
    else:
        log.debug("First load not correct...")
        log.debug(str(cv2.countNonZero(b1)))
        log.debug(str(cv2.countNonZero(g1)))
        log.debug(str(cv2.countNonZero(r1)))
        return 0
    if LoadGood==True and ImgDiffGreen>0.40 or LoadGood==True and ImgDiffRed>0.40: #0.40 is for moused grayness , 0.9 is normal
        DeleteFile("LoadCheck.png")
        DeleteFile("ScrapLoadCheck.png")
        log.info("Loading of page completed...")
        return 1 #Basically not needed because when def ends here it will return None which is not 0 so loop continues but its nice way of programming
    else:
        log.info('Loading not finished , waiting...')
        return 0
##############################################################################################
def ImageProcess(IMGName,range2=-1,range=0,tescfg=tessdefault_config): #Image Data Extraction
    log = logging.getLogger('ImageProcess')
    try:
        img = cv2.imread(str(IMGName)+'.png') #Read Image
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) #Convert to Grayscale
        img = cv2.bitwise_not(cv2.threshold(img, 200, 200, cv2.THRESH_BINARY)[1]) #Thresh image for denoising , then invert
        img = cv2.resize(img, None, fx=2, fy=2,interpolation=cv2.INTER_CUBIC) #Resize for better image quality + cubic interpolation for smoothing
        value=pytesseract.image_to_string(img,config=tescfg)[range:range2] #Process image    
    except:
        log.debug(IMGName+" error!")
        log.exception('')
        return "error"
    else:
        return value
##############################################################################################
#print ([k for k in logging.Logger.manager.loggerDict]) #### totally not from stackoverflow :)
##############################################################################################
def WindowFullscreen(presleep=1,postsleep=0.1):
    log = logging.getLogger('WindowFullscreen')
    time.sleep(presleep) #Used for securing correct window 
    window=gw.getActiveWindow()   #Forcing window maximizing
    if window.isMaximized == True: 
        log.debug("Window is maximized")
    else :
        window.maximize()
        log.debug("Maximizing window")
    time.sleep(postsleep) #Used to wait for maximizing to happen
##############################################################################################
def DeleteFile(value):
    log = logging.getLogger('DeleteFile')
    if DeleteImagesAfterUsage == True: #Extremely lazy cleaning ....
        try:
            os.remove(value)
        except:
            log.debug("Delefing file : "+value+" failed.")
        else:
            log.debug("Delefing file : "+value+" successfull.")
##############################################################################################
def ImageGrab(MachineName,HourValue,MachineURL): #AFTER TIME CHECK - OK 
    log = logging.getLogger('ImageGrab')
    
    os.system("start chrome --start-maximized --new-window "+URLList[MachineURL]) #Start Browser (need to add auto login - most likely just mouse clicks)   
    WindowFullscreen()
    
    LoadTimer=0
    LoadRefresh=1
    while LoadCheck()==0: #not great way to check but works
        SleepValue = 0.5 #TEMP
        LoadTimer=LoadTimer+SleepValue
        time.sleep(SleepValue)
        log.debug('LoadTimer= '+str(LoadTimer))
        log.info("Waiting for page load")
        if LoadTimer>=15 and LoadRefresh==1:
            keyboard.press_and_release('F5')
            log.info('Refreshing page')
            LoadRefresh=0
        if LoadTimer>=35:
            log.info('Skipping loading...')
            break
        
    for each in ScreenRegionDict:
        log.debug("Image Creation: "+each)
        x=ScreenRegionDict[each] #Redundant line but looks better ? + less listing through dict ? performance gains ?
        ScreenshotRegion(each+MachineName+str(HourValue),x[0],x[1],x[2],x[3])

    MouseCur=mouse.get_position() #Close Window (not pretty simualing mouse clicks but i dont want to mess with os.kill and os.pid)
    mouse.move(10,10)
    mouse.click("middle")
    mouse.move(MouseCur[0],MouseCur[1])
    
##############################################################################################
def ExcelOutput(MachineName,HourValue,ShiftCheck,PathToFile): #MAYBE REWRITE TO THREADING FOR ALL VALUES TO GET WRITE AT SAME TIME
    log = logging.getLogger('ExcelOutput') # FULL TIME 1.4S , CPU TIME 0.34S
    if TimeDEBUG==1:
        StartTime=time.time()
        StartTimeCPU=time.process_time()
    log.info("Applying data to excel")  #Excel Start
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
            case "OEE" | "Scrap" : SheetDict[each][1] = ImageProcess(each+MachineName+str(HourValue),-1,tescfg=tessadvnumber_config) #NOT TRUE , NEED TO MONITOR DATA -> Extra removing % , it safer to detect and then remove than to block with blacklist
            case "Product" : SheetDict[each][1] = ImageProcess(each+MachineName+str(HourValue),-5) #Should be good , need more data
            case "OK" : #Hardcoded cleaning of value + Maybe still broken + Not pretty at all
                x = ImageProcess(each+MachineName+str(HourValue),4,tescfg=tessnumber_config)
                if x=="error" or x=="": #Not pretty handling
                    SheetDict[each][1] = "error"
                    continue  #Currently dont write error to excel just skips whole OK check , better than breaking code without logs :)
                while x.isdigit() == False : #Still too many infinite loops
                    log.debug("OK before:"+x)
                    y=negative(len(x)-1) 
                    #x = x[y:]
                    x = x[:-y]
                    log.debug("OK after :"+x)
                    #time.sleep(0.1)   #Maybe still broken , need to log data
                SheetDict[each][1] = x
            case 'NOK': # NEED TO FINISH , BAD REPORTING OF 0,1 
                SheetDict[each][1] = ImageProcess(each+MachineName+str(HourValue),-1,tescfg=tessnumber_config) #Need to log data
            case _ : SheetDict[each][1] = ImageProcess(each+MachineName+str(HourValue),-1)

        log.debug(str(each)+" : "+str(SheetDict[each][0]+' : '+str(SheetDict[each][1])))
        sheet[str(SheetDict[each][0])]=str(SheetDict[each][1])   #Writing into excel
        DeleteFile(str(each+MachineName+str(HourValue))+".png")
    # !!MESSY!!

    try: #Saving Excel Document #Very barebones 
        wb.save(filename="AutoData_gen.xlsx")  #NEED TO ADD TO CONFIG LATER
    except:
        log.warning("Excel edit failed") 
    else:
        log.info('Excel edit success')
    if TimeDEBUG==1:
        EndTime=time.time() - StartTime
        EndTimeCPU=time.process_time() - StartTimeCPU
        log.debug('Execution time FULL - EXCEL OUTPUT :'+str(EndTime))
        log.debug('Execution time CPU - EXCEL OUTPUT :'+str(EndTimeCPU))
##############################################################################################
def ExcelUpdate(File,PreSleep=1,PostSleep=0.1,Save=False):
    os.startfile(File) #NOT PRETTY WAY TO UPDATE DATA
    WindowFullscreen(PreSleep,PostSleep)
    if Save == True : keyboard.press_and_release('Ctrl+S')
    MouseCur=mouse.get_position() #Close Window (os.kill dont work without admin privileges)
    mouse.move(1900,10)
    mouse.click("left")
    mouse.move(MouseCur[0],MouseCur[1])
##############################################################################################

