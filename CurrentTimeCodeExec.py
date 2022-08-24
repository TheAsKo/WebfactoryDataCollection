#Main Code Logic -Threading+Scheduling
import time
import logging
import os
import CurrentMainV2
import threading
logging.getLogger().setLevel(logging.DEBUG) #Need to add to GUI as check + need to clean/add logging


class TimeValue():
    ShiftCheck=0
    OneMin=0
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

    def __run__():  #MESSY 
        while TimeValue.OneMin==0 : #UNUSED BUT MAYBE REPURPOSED LATER
            x=time.strftime("%H:%M:%S",time.gmtime())#time.ctime()[11:19] #Maybe switch to time.strtime()
            if 48 < int(x[3:5]) < 50 or int(x[3:5]) >= 58: #Minute check
                while True: #Timing will be needed to be rewritten when I figure out threading
                    #x=time.ctime()[11:19]
                    logging.debug("One sec cycle:"+x)
                    y=TimeValue.TimeCaseCheck(x)[0]
                    if y>0:
                        logging.debug("Writing hour:"+str(y))               
                        #f = open("DataTransfer.txt", "w")
                        #f.write(str(y))
                        #f.close()
                        return int(y) 
                    else:
                        return 0
            logging.debug("One Min Cycle : "+x)
            time.sleep(60)

ThreadDict={'MachineName':['FILL','PETIG2'],'MachineURL':[1,2],'ShiftCheck':[8,12]}

# If hour is OK and machine is in correct shiftcheck then do correct cycle
# 
# Case 1 = both machines runs
# Case 2 = only 8H machine run
# Case 3 = only 12h machine run
#
# Case V2 = only 8H or 12H machines run (with checking)
# Case V2 _ = default, all machines run
#
def TimeLoop():
    logging.info('Time Loop Start')
    while TimeValue.__run__() == 0:
        time.sleep(1)
    HourVal=int(TimeValue.__run__())
    logging.debug("Process hour aquired")
    y=0
    run=1
    while len(ThreadDict['MachineName'])>y and run==1: ### ULTRA MESSY
        match TimeValue.ShiftCheck:
            case 8 | 12:   #Maybe could be compressed to one cycle after choosing which machines runs when ?
                for each in ThreadDict['MachineName']: #### !!!! REWRITED LOOPING NEED TO TEST !!!!
                    logging.debug(len(ThreadDict['MachineName']))
                    logging.debug(str(TimeValue.ShiftCheck)+" = "+str(ThreadDict['ShiftCheck'][y]))
                    logging.debug('each :'+str(each))
                    if TimeValue.ShiftCheck == ThreadDict['ShiftCheck'][y]:
                        logging.info("Starting data gather for: "+ThreadDict['MachineName'][y])
                        thread1=threading.Thread(target=CurrentMainV2.ImageGrab(MachineName=ThreadDict['MachineName'][y],HourValue=str(HourVal),MachineURL=ThreadDict['MachineURL'][y]))
                        logging.debug('Thread 1 starting')
                        thread1.start()
                        thread1.join()
                        logging.debug('Thread 1 finished')
                        thread2=threading.Thread(target=CurrentMainV2.ExcelOutput(MachineName=ThreadDict['MachineName'][y],HourValue=str(HourVal)))
                        logging.debug('Thread 2 starting')
                        thread2.start()
                        y=y+1
                    else: 
                         y=y+1
                         logging.debug('skipping machine')
                          
            case _:
                logging.info("Starting data gather for: "+ThreadDict['MachineName'][y])
                thread1=threading.Thread(target=CurrentMainV2.ImageGrab(ThreadDict['MachineName'][y],HourValue=HourVal,MachineURL=ThreadDict['MachineURL'][y]))
                logging.debug('Thread1 starting')
                thread1.start()
                thread1.join()
                logging.debug('Thread 1 finished')
                thread2=threading.Thread(target=CurrentMainV2.ExcelOutput(ThreadDict['MachineName'][y],HourValue=HourVal))
                logging.debug('Thread 2 starting')
                thread2.start()
                y=y+1
    run=0
    TimeValue.OneMin=0 #UNUSED BUT MAYBE REPURPOSED LATER
    y=0
    HourVal=0
    logging.info("Time Loop Finished")

while True:
    logging.info('Program Cycle started')
    thread=threading.Thread(target=TimeLoop())
    thread.start()
    thread.join()
    logging.info('Program Cycle ended')
