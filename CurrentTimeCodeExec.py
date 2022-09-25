#Main Code Logic -Threading+Scheduling
##############################################################################################
#Imports
import ConfigHandler as Config
import time
import logging
import CurrentMainV2
import threading
import ShiftStart
import os
import sys 
##############################################################################################
# Declarations
logging.getLogger().setLevel(logging.DEBUG)
ThreadDict=Config.ConfigRead('TIMECODE','ThreadDict','dict')
TimeDebug = Config.ConfigRead('MAIN','TimeDebug','bool')
##############################################################################################
#print('sys.argv[0] =', sys.argv[0])     #BACKUP FOR GETTING FOLDER LOCATION        
pathname = os.path.dirname(sys.argv[0])  #FINDING WHERE IS THIS FILE LOCATED      
#print('path =', pathname)
#print('full path =', os.path.abspath(pathname))
##############################################################################################
class TimeValue():
    log = logging.getLogger('TimeValue')
    ShiftCheck=0
    Sleep=0
    def TimeCaseCheck(x):  # Shift 3: 6-14,14-22,22-6 Shift 4:6-18,18-6
        match x:        #Need to check intervals when webfactory is switching on 3shift
            case '00:59:54'|'00:59:55': TimeValue.ShiftCheck=0  ; return 1,
            case '01:59:54'|'01:59:55': TimeValue.ShiftCheck=0  ; return 2, #One time I saw time script missing exactly 1 sec needed to switch
            case '02:59:54'|'02:59:55': TimeValue.ShiftCheck=0  ; return 3, #So i added second case to have 2sec catch window
            case '03:59:54'|'03:59:55': TimeValue.ShiftCheck=0  ; return 4, #Hopefully suffice
            case '04:59:54'|'04:59:55': TimeValue.ShiftCheck=0  ; return 5,
            case '05:49:49'|'05:49:50': TimeValue.ShiftCheck=0  ; return 6, #We never work continuous on 6AM
            case '06:59:54'|'06:59:55': TimeValue.ShiftCheck=0  ; return 7,
            case '07:59:54'|'07:59:55': TimeValue.ShiftCheck=0  ; return 8,
            case '08:59:54'|'08:59:55': TimeValue.ShiftCheck=0  ; return 9,
            case '09:59:54'|'09:59:55': TimeValue.ShiftCheck=0  ; return 10,
            case '10:59:54'|'10:59:55': TimeValue.ShiftCheck=0  ; return 11,
            case '11:59:54'|'11:59:55': TimeValue.ShiftCheck=0  ; return 12,
            case '12:59:54'|'12:59:55': TimeValue.ShiftCheck=0  ; return 13,
            case '13:49:40'|'13:49:41': TimeValue.ShiftCheck=0  ; return 14, #TEMP FIX TO BAD READS
            #case '13:49:49'|'13:49:50': TimeValue.ShiftCheck=8  ; return 14,
            #case '13:58:54'|'13:58:55': TimeValue.ShiftCheck=12 ; return 14, #OEE ALWAYS MISSING
            case '14:59:54'|'14:59:55': TimeValue.ShiftCheck=0  ; return 15,
            case '15:59:54'|'15:59:55': TimeValue.ShiftCheck=0  ; return 16,
            case '16:59:54'|'16:59:55': TimeValue.ShiftCheck=0  ; return 17,
            case '17:49:49'|'17:49:50': TimeValue.ShiftCheck=12 ; return 18,
            case '17:58:54'|'17:58:55': TimeValue.ShiftCheck=8  ; return 18,
            case '18:59:54'|'18:59:55': TimeValue.ShiftCheck=0  ; return 19,
            case '19:59:54'|'19:59:55': TimeValue.ShiftCheck=0  ; return 20,
            case '20:59:54'|'20:59:55': TimeValue.ShiftCheck=0  ; return 21,
            case '21:49:40'|'21:49:50': TimeValue.ShiftCheck=8  ; return 22, #TEMP FIX TO BAD READS
            #case '21:49:49'|'21:49:50': TimeValue.ShiftCheck=8  ; return 22,
            #case '21:58:54'|'21:58:55': TimeValue.ShiftCheck=12 ; return 22, #OEE ALWAYS MISSING 
            case '22:59:54'|'22:59:55': TimeValue.ShiftCheck=0  ; return 23,
            case '23:59:54'|'23:59:55': TimeValue.ShiftCheck=0  ; return 24,
            case _ : TimeValue.ShiftCheck=0 ; return 0, #0

    def __run__(): #MESSY
        if OneCycle==1: #TEMP
            return TimeValue.TimeCaseCheck('05:49:50')[0]  #TEMP
        x=time.strftime("%H:%M:%S",time.localtime())
        if 48 < int(x[3:5]) < 50 or int(x[3:5]) >= 58: #Minute check
            TimeValue.log.debug("One sec cycle:"+x)
            y=TimeValue.TimeCaseCheck(x)[0]
            if y>0:
                TimeValue.log.debug("Writing hour:"+str(y))
                return int(y) 
            else:
                TimeValue.Sleep=1
                return 0     
        TimeValue.log.debug("One Min Cycle : "+x)
        TimeValue.Sleep=60
        return 0
##############################################################################################
def ActualMachineIndexing():
    log = logging.getLogger('ActualMachineIndexing')
    ActualMachine={"Machine":[],"URL":[],"ShiftCheck":[]} #USED LOCALLY FOR INDEXING
    for i in range(len(ThreadDict['MachineName'])): #Gathering which machines i want to run
        if ThreadDict['MachineActive'][i] == 1 and TimeValue.ShiftCheck == ThreadDict['ShiftCheck'][i] or TimeValue.ShiftCheck == 0 and ThreadDict['MachineActive'][i] == 1: #Split active check before or leave it like this
            ActualMachine['Machine'] = ActualMachine['Machine'] + [ThreadDict['MachineName'][i]]
            ActualMachine['URL'] = ActualMachine['URL'] + [ThreadDict['MachineURL'][i]]
            ActualMachine['ShiftCheck'] = ActualMachine['ShiftCheck'] + [ThreadDict['ShiftCheck'][i]]
            log.info('Machine queued: '+str(ThreadDict['MachineName'][i]))
        else : log.debug('Machine skipped: '+str(ThreadDict['MachineName'][i]))
    return ActualMachine
##############################################################################################
def TimeLoop():
    log = logging.getLogger('TimeLoop')
    log.info('Time Loop Start')
    while TimeValue.__run__() == 0: #Waiting for time trigger , maybe not best way for checking when value change
        time.sleep(TimeValue.Sleep) 
    HourVal=str(TimeValue.__run__())
    log.debug("Process hour aquired")
    
    if TimeDebug == True :
        StartTimeLoop=time.time()
        StartTimeCPULoop=time.process_time()
    
    ActiveMachine=ActualMachineIndexing()
    
    for i in range(len(ActiveMachine['Machine'])): #Not cleanest thing but works
        log.info("Starting data gather for: "+ActiveMachine['Machine'][i])
        log.debug('Thread 1 starting')
        thread1=threading.Thread(target=CurrentMainV2.ImageGrab(MachineName=ActiveMachine['Machine'][i],HourValue=HourVal,MachineURL=ActiveMachine["URL"][i]))
        thread1.start()
        thread1.join()
        log.debug('Thread 1 finished')
        log.debug('Thread 2 starting')
        thread2=threading.Thread(target=CurrentMainV2.ExcelOutput(MachineName=ActiveMachine['Machine'][i],HourValue=HourVal,ShiftCheck=ActiveMachine['ShiftCheck'][i],PathToFile=pathname))
        thread2.start()
    log.info("Time Loop Finished")
    
    if TimeDebug == True :
        EndTimeLoop=time.time() - StartTimeLoop
        EndTimeCPULoop=time.process_time() - StartTimeCPULoop
        log.debug('Execution time FULL - TIME LOOP :'+str(EndTimeLoop))
        log.debug('Execution time CPU - TIME LOOP :'+str(EndTimeCPULoop))
##############################################################################################
OneCycle=1 #USED FOR DEBUG
StartCycle=1
while True: #Main Cycle ...
    if StartCycle==1:
        MachineCreationDict=ActualMachineIndexing()
        ShiftStart.FileCreation.__run__(True,True,0,0,MachineCreationDict,pathname)
        StartCycle=0
    logging.info('Program Cycle started')
    thread=threading.Thread(target=TimeLoop())
    thread.start()
    thread.join()
    CurrentMainV2.ExcelUpdate('AutoData_gen.xlsx') #Moved update to end of cycle 
    logging.info('Program Cycle ended')
    if OneCycle==1: #TEMP
        break #TEMP
