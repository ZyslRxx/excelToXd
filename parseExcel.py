#coding:utf-8
import xlrd
import sys
import os
import struct
import logging
import logging.config

logging.config.fileConfig('logger.conf')
logger = logging.getLogger('example01')

class ParseExcel():
    """author zhangys 
       firstversion 0.1
       date 2015-12-22
       仿造lucidworld 使用python根据excel生成xd文件
    """
    #logging.basicConfig(level=logging.DEBUG,
    #            format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
    #            datefmt='%a, %d %b %Y %H:%M:%S',
    #            filename='parse.log',
    #            filemode='w')

    def __init__(self,filename):
        self.excelname = filename
        if not os.path.exists(filename):
            logger.error("{} is not exist".format(filename))
            sys.exit(-1)
        #加载excel
        self.sheetVo = xlrd.open_workbook(filename)


    def sheetToXd(self,sheet):
        """ every sheet paser until sheetname 'end'"""
        lines = sheet.nrows
        if lines < 5:
            logger.error("{0} sheet lines is only {1},no less than 5 lines".format(sheet.name,lines))
            sys.exit(-1)

        #首行第二列版本号，第二行第二列var
        version = sheet.row(0)[1]
        var = sheet.row(1)[1]
        #'I'表示int,'!'表示使用big-endian,用4个字节存储转成字节流
        data = struct.pack('!I',int(version.value))
        logging.debug('var :{}'.format(var.value))
        uvar = ''
        if  var.value:
            #uvar = unicode.encode(var.value,'utf-8')
            pass
        
        uvar = struct.pack('H',len(uvar))+uvar
        data += uvar
        tmpData = ''
        self.types = sheet.row(3)
        for line in range(4,lines):
            lineData = sheet.row(line)
            tmpData += self.parseData(lineData)
        uDLen = struct.pack('!I',lines-4)
        data += uDLen
        data += tmpData
        
        logger.debug("total bytes : {0},tmpdata length:{1}".format(len(data),len(tmpData)))

        xdName = self.excelname.split('.')[0]+"_"+sheet.name+'.xd'

        with open(xdName,'wb') as f:
            f.write(data)
    
    def parseData(self,lineData):
        tmpData = ''
        """ every line parse"""

        for tp,d in zip(self.types,lineData):
            if not  tp.value:
                logger.debug("tp.value:{}".format(tp.value))
                break
            bData = None

            #type类型int,String,byte,short.float,long
            if 'int' == tp.value:
                bData = struct.pack('!I',int(d.value))
            elif 'String' == tp.value:
                if isinstance(d.value,float):
                    d.value = str(int(d.value))
                bData = d.value.encode("utf-8")
                bLen = struct.pack('!H',len(bData))
                bData = bLen+bData
                logger.debug("String data :{}".format(bData))

            elif 'byte' == tp.value:
                bData = struct.pack('B',int(d.value))
            elif 'short' == tp.value:
                bData = struct.pack('!H',int(d.value))
            elif 'float' == tp.value:
                bData = struct.pack('!f',float(d.value))
            elif 'long' == tp.value:
                bData = struct.pack('!L',long(d.value))
            else :
                logging.error("type:{} is not exist".format(tp.value))

            tmpData += bData

        return tmpData

    def parseSheets(self):
        sheets = self.sheetVo.sheets()
        
        for sheet in sheets:
            self.sheetToXd(sheet)

if __name__ == '__main__':
    excel = ParseExcel('equip-cfg.xls')

    excel.parseSheets()



 


        


