#Main Code Logic -Threading+Scheduling
########################
#Imports
import time
import logging
import CurrentMainV2
import threading
########################
logging.getLogger().setLevel(logging.DEBUG) 
########################
class TimeValue():
    ShiftCheck=0
    Sleep=0
    def TimeCaseCheck(x):  # Shift 3: 6-14,14-22,22-6 Shift 4:6-18,18-6
        match x:        #Need to check intervals when webfactory is switching on 3shift
            case '00:59:55': TimeValue.ShiftCheck=0  ; return 1,
            case '01:59:55': TimeValue.ShiftCheck=0  ; return 2,
            case '02:59:55': TimeValue.ShiftCheck=0  ; return 3,
            case '03:59:55': TimeValue.ShiftCheck=0  ; return 4,
            case '04:59:55': TimeValue.ShiftCheck=0  ; return 5,
            case '05:49:50': TimeValue.ShiftCheck=0  ; return 6, #We never work continuous on 6AM
            case '06:59:55': TimeValue.ShiftCheck=0  ; return 7,
            case '07:59:55': TimeValue.ShiftCheck=0  ; return 8,
            case '08:59:55': TimeValue.ShiftCheck=0  ; return 9,
            case '09:59:55': TimeValue.ShiftCheck=0  ; return 10,
            case '10:59:55': TimeValue.ShiftCheck=0  ; return 11,
            case '11:59:55': TimeValue.ShiftCheck=0  ; return 12,
            case '12:59:55': TimeValue.ShiftCheck=0  ; return 13,
            case '13:49:50': TimeValue.ShiftCheck=8  ; return 14,
            case '13:58:55': TimeValue.ShiftCheck=12 ; return 14,
            case '14:59:55': TimeValue.ShiftCheck=0  ; return 15,
            case '15:59:55': TimeValue.ShiftCheck=0  ; return 16,
            case '16:59:55': TimeValue.ShiftCheck=0  ; return 17,
            case '17:49:50': TimeValue.ShiftCheck=12 ; return 18,
            case '17:58:55': TimeValue.ShiftCheck=8  ; return 18,
            case '18:59:55': TimeValue.ShiftCheck=0  ; return 19,
            case '19:59:55': TimeValue.ShiftCheck=0  ; return 20,
            case '20:59:55': TimeValue.ShiftCheck=0  ; return 21,
            case '21:49:50': TimeValue.ShiftCheck=8  ; return 22,
            case '21:58:55': TimeValue.ShiftCheck=12 ; return 22,
            case '22:59:55': TimeValue.ShiftCheck=0  ; return 23,
            case '23:59:55': TimeValue.ShiftCheck=0  ; return 24,
            case _ : TimeValue.ShiftCheck=0 ; return 0, #0

    def __run__(): #MESSY
        #return TimeValue.TimeCaseCheck('06:59:55')[0]  #DEBUG ALL MACHINES 
        x=time.strftime("%H:%M:%S",time.localtime())
        if 48 < int(x[3:5]) < 50 or int(x[3:5]) >= 58: #Minute check
            logging.debug("One sec cycle:"+x)
            y=TimeValue.TimeCaseCheck(x)[0]
            if y>0:
                logging.debug("Writing hour:"+str(y))               
                return int(y) 
            else:
                TimeValue.Sleep=1
                return 0     
        logging.debug("One Min Cycle : "+x)
        TimeValue.Sleep=60
        return 0

ThreadDict={'MachineName':['FILL','PETIG2'],'MachineURL':[1,2],'ShiftCheck':[-1,12]} #In future input from UI/cfg file

def TimeLoop():
    logging.info('Time Loop Start')
    while TimeValue.__run__() == 0: #Waiting for time trigger , maybe not best way for checking when value change
        time.sleep(TimeValue.Sleep)
    HourVal=str(TimeValue.__run__())
    logging.debug("Process hour aquired")

    ActualMachine={"Machine":[],"URL":[],"ShiftCheck":[]} 
    for i in range(len(ThreadDict['MachineName'])): # Gathering which machines i want to run 
        if TimeValue.ShiftCheck == ThreadDict['ShiftCheck'][i] or TimeValue.ShiftCheck == 0 and ThreadDict['ShiftCheck'][i] >= 0: #TEMP FIX TO DISABLE MACHINE
            ActualMachine['Machine'] = ActualMachine['Machine'] + [ThreadDict['MachineName'][i]]
            ActualMachine['URL'] = ActualMachine['URL'] + [ThreadDict['MachineURL'][i]]
            ActualMachine['ShiftCheck'] = ActualMachine['ShiftCheck'] + [ThreadDict['ShiftCheck'][i]]
            logging.debug('Machine queued: '+str(ThreadDict['MachineName'][i]))
        else : logging.debug('Machine skipped: '+str(ThreadDict['MachineName'][i]))
    
    for i in range(len(ActualMachine['Machine'])): # !!!! REWRITED LOOPING NEED TO TEST !!!!
        logging.info("Starting data gather for: "+ActualMachine['Machine'][i])
        logging.debug('Thread 1 starting')
        thread1=threading.Thread(target=CurrentMainV2.ImageGrab(MachineName=ActualMachine['Machine'][i],HourValue=HourVal,MachineURL=ActualMachine["URL"][i]))
        thread1.start()
        thread1.join()
        logging.debug('Thread 1 finished')
        logging.debug('Thread 2 starting')
        thread2=threading.Thread(target=CurrentMainV2.ExcelOutput(MachineName=ActualMachine['Machine'][i],HourValue=HourVal,ShiftCheck=ActualMachine['ShiftCheck'][i]))
        thread2.start()
    logging.info("Time Loop Finished")

while True: #Main Cycle ...
    logging.info('Program Cycle started')
    thread=threading.Thread(target=TimeLoop())
    thread.start()
    thread.join()
    logging.info('Program Cycle ended')
    #break #TEMP

