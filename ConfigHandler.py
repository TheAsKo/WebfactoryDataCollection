import configparser
import logging
import json

config=configparser.ConfigParser()
configparser.BasicInterpolation()
logging.getLogger().setLevel(logging.DEBUG)

def DefaultConfigWrite():
    config['MAIN'] = {'OnlyRootDebug':'True',
                    'tess_cmd':'C:/Users/nsz.fu.montaz/AppData/Local/Tesseract-OCR/tesseract.exe',
                    'tessdefault_config':'-c tessedit_char_whitelist 0123456789., tessedit_char_blacklist _',
                    'tessnumber_config':'--psm 3 --oem 3 -c tessedit_char_whitelist=0123456789/-., tessedit_char_blacklist _',
                    'DeleteImagesAfterUsage':'True',
                    'LoggerLevel':'DEBUG'}
    config['MAIN']['URLList'] = '["https://wf-nsm.neuman.at/clients/wf-login/#/","http://wf-nsm.neuman.at/clients/wf-mes/sk/#/wfmes/view/(mainview:msc/349)","http://wf-nsm.neuman.at/clients/wf-mes/sk/#/wfmes/view/(mainview:msc/346)"]'
    config['MAIN']['ScreenRegionDict'] = '{"OEE":[600,230,210,70],"OK":[1300,680,260,100],"NOK":[1680,680,190,80],"Product":[1020,620,600,50],"Scrap":[1700,230,120,40],"Norm":[1020,670,120,50]}'
    config['MAIN']['SheetDict'] = '{"TimeFlag":["A2","1"],"TimeSend":["B2","1"],"OEE":["C2","1"],"OK":["D2","1"],"NOK":["E2","1"],"Product":["F2","1"],"Scrap":["G2","1"],"Norm":["H2","1"]}'
    config['MAIN']['LoadCheck'] = '{"LoadCheck":[0,980,180,40],"ScrapLoadCheck":[1820,230,70,70]}'

    config['TIMECODE'] = {}
    config['TIMECODE']['ThreadDict'] = '{"MachineName":["FILL","PETIG2"],"MachineURL":[1,2],"ShiftCheck":[8,12],"MachineActive":[0,1]}'

    with open('config.ini', 'w') as configfile:
        config.write(configfile)

def ConfigRead(value1,value2,Type=None):
    config.read('config.ini')
    ValueX=config[value1]
    match Type:
        case 'dict' | 'list' : 
            try : 
                return json.loads(ValueX[value2])
            except:
                logging.warning("Requested value "+value2+" failed to load as dictionary/list!")
        case 'int' : 
            try:
                return ValueX.getint(value2)
            except:
                logging.warning("Requested value "+value2+" is not integer!")
        case 'float' : 
            try:
                return ValueX.getfloat(value2)
            except:
                logging.warning("Requested value "+value2+" is not float!")
        case 'bool' : 
            try:
                return ValueX.getboolean(value2)
            except:
                logging.warning("Requested value "+value2+" is not bool!")
        case _ : 
            try:
                return ValueX[value2]
            except:
                logging.warning("Requested value "+value2+" failed to load!")

def ConfigWrite(data1,data2,value,Type=None):
    match Type:
        case 'list' | 'dict':
            logging.critical('Not Finished') #NEED TO FINISH
        case 'int' | 'float' | 'bool' | 'str' | _ :
            try :
                config[data1][data2]=str(value)
            except:
                logging.warning('Updating value of '+data2+' failed!')
            else:
                logging.debug('Edited value of '+data2+' to: '+value)
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

#DefaultConfigWrite()