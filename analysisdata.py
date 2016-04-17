# -*- coding=utf-8 -*-

import os
import time
import urllib, urllib2
import random
import profile
import pstats
import xlsxwriter
import xlrd


'''
1. 读取已经分析出的xlsx表格模块制作，读取xlsx表之后，返回对应数据结构
2. 利用基准数据结构，分析待检测数据项。
3. 产出分析报表
'''

ignore = [u"rpyc"]


class ExcelReader():
    def __init__(self,excelfilename):

        self.excelfilename = excelfilename
        # self.excelsheetnames = excelsheetnames
        self.loadExcel(self.excelfilename)
        pass

    def loadExcel(self, filename):
        try:
            self.excelFile = xlrd.open_workbook(filename)
            self.excelSheets = self.excelFile.sheet_names()
        except Exception, e:
            print str(e)

    def getSheet(self, sheet):
        try:
            if isinstance(sheet, str):
                # print "read sheet by utf-8 str"
                return self.excelFile.sheet_by_name(sheet.decode("utf"))
            elif isinstance(sheet, unicode):
                # print "read sheet by unicode str"
                return self.excelFile.sheet_by_name(sheet)

            return self.excelFile.sheet_by_index(sheet)
        except:
            return None

    def getData2(self):
        # use this funciton to get the dict which is containning the gmcommands paramaters..

        dicttoreturn = {}

        for i in self.excelSheets:
            if i is not None:
                # print isinstance(i, unicode)
                list = {}
                current_sheet = self.getSheet(i)
                # print current_sheet,"test"
                colnames = current_sheet.row_values(0)  # get the colnames
                # print colnames,"test2"
                #ncols = current_sheet.ncols  # lie
                nrows = current_sheet.nrows  # hang

                for rownum in range(1, nrows):
                    row = current_sheet.row_values(rownum)
                    if row:
                        command_info = {}
                        for j in range(len(colnames)):
                            command_info[colnames[j]] = row[j]
                        list[int(row[0])] = command_info

                #print list
                dicttoreturn[i] = list

        return dicttoreturn

    def getData(self):
        dicttoreturn = {}

        # print isinstance(i, unicode)
        list = {}
        current_sheet = self.getSheet(0)
        # print current_sheet,"test"
        colnames = current_sheet.row_values(0)  # get the colnames
        # print colnames,"test2"
        #ncols = current_sheet.ncols  # lie
        nrows = current_sheet.nrows  # hang

        for rownum in range(1, nrows):
            row = current_sheet.row_values(rownum)
            #print row

            # 在这里做一些筛选，去除需要忽略的函数名称和辅助模块等。

            if row:
                #print row[6]
                continueflag = False
                for i in ignore:
                    if i in row[6]:
                        continueflag = True

                        print row[6] + "  has been ignored..."

                        break
                if continueflag is False:
                    command_info = {}
                    for j in range(len(colnames)):
                        command_info[colnames[j]] = row[j]
                    list[int(rownum)] = command_info

        return list

class AnalysierBase():
    '''
    获取base数据。利用base数据分析待分析数据。 随后产出分析表。
    '''
    def __init__(self,baseDataFilename,targetDataFilename,outputFilename):
        '''
        利用上面的预读类库，读取excel 并将转存，留待分析
        :param baseDataFilename:  string filename .xlsx
        :param targetDataFilename: string filename .xlsx
        :param outputFilename: string filename .xlsx
        :return no return:
        '''
        self.baseData = self.loadbase(baseDataFilename)
        self.targetData = self.loadbase(targetDataFilename)

        self.workbook = xlsxwriter.Workbook(outputFilename)
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.set_column('A:B',8)
        self.worksheet.set_column('C:C',80)
        self.worksheet.set_column('D:E',8)
        self.worksheet.set_column('F:K',14)

    def loadbase(self,DataFilename):
        '''
        利用excelreader类，将excel读取到内存。
        :param DataFilename:
        :return:
        '''
        datard = ExcelReader(DataFilename)
        return datard.getData()

    def doanalysis(self):
        list = []
        for (basekey,basevalue) in self.baseData.items():
            for (targetkey,targetvalue) in self.targetData.items():
                if targetvalue.get(u'filename:lineno(function)') == basevalue.get(u'filename:lineno(function)'):
                    #print targetvalue
                    #print "id,函数名,total/ncalls,cumtime/ccalls,total/ncall偏离值,cumtime/calls偏离值"
                    #if u"rpyc" in targetvalue.get(u'filename:lineno(function)') or
                    line = [targetkey,
                            targetvalue.get(u'filename:lineno(function)'),
                            targetvalue.get(u'ncalls'),
                            targetvalue.get(u'ccalls'),
                            targetvalue.get(u'totaltime'),
                            targetvalue.get(u'totaltime/ncalls'),
                            targetvalue.get(u'cumtime'),
                            targetvalue.get(u'cumtime/ccalls'),
                            float(targetvalue.get(u'totaltime/ncalls')) - float(basevalue.get(u'totaltime/ncalls')),
                            float(targetvalue.get(u'cumtime/ccalls')) - float(basevalue.get(u'cumtime/ccalls'))]
                    list.append(line)

        return list

    def print_title(self):
        self.worksheet.write(0,0,"ID")
        self.worksheet.write(0,1,"oirginID")
        self.worksheet.write(0,2,"functionInfo")
        self.worksheet.write(0,3,"ncalls")
        self.worksheet.write(0,4,"ccalls")
        self.worksheet.write(0,5,"totaltime")
        self.worksheet.write(0,6,"totaltime/ncalls")
        self.worksheet.write(0,7,"cumtime")
        self.worksheet.write(0,8,"cumtime/ccalls")
        self.worksheet.write(0,9,"offset-totaltime/ncalls")
        self.worksheet.write(0,10,"offset-cumtime/ccalls")

    def datatoexcel(self):
        self.print_title()  # 为excel添加titile。
        list = self.doanalysis()
        for index,line in enumerate(list):
            self.worksheet.write(index+1,0,index)   # id
            self.worksheet.write(index+1,1,line[0]) # old id
            self.worksheet.write(index+1,2,line[1]) # function info
            self.worksheet.write(index+1,3,line[2]) # ncalls
            self.worksheet.write(index+1,4,line[3]) # ccalls
            self.worksheet.write(index+1,5,line[4]) # totaltime
            self.worksheet.write(index+1,6,line[5]) # total/ncalls
            self.worksheet.write(index+1,7,line[6]) # cumtime
            self.worksheet.write(index+1,8,line[7]) # cumtime/ccalls
            self.worksheet.write(index+1,9,line[8]) # offset1
            self.worksheet.write(index+1,10,line[9]) # offset2

        self.workbook.close()
        pass

class AnalysiserAverage():
    def __init__(self,DataFileList):
        #此部分用于平均值分析。多个表格整合为一个表格产出，将所有表格中的数据，全部统计多次出现取平均值。

        #for Datafilename in DataFileList:
        self.baseData = self.loadbase(DataFileList)
        print self.baseData
        pass

    def loadbase(self,DataFilename):
        '''
        利用excelreader类，将excel读取到内存。
        :param DataFilename:
        :return:
        '''
        datard = ExcelReader(DataFilename)
        return datard.getData()



if __name__ == "__main__":
    # er = ExcelReader("test_prof_3v3_60s.xlsx")
    # file1 = open("temp","w")
    # print >> file1, er.getData()
    # data = er.getData()
    # name = data.get(1).get(u'filename:lineno(function)')
    # print u"acquire" in name
    # file1.close()
    #ab = AnalysierBase("test_prof_3v3_60s.xlsx","test_prof_3v3_60s.xlsx","test.xlsx")
    #ab.datatoexcel()
    aa = AnalysiserAverage("test_prof_3v3_60s.xlsx")
