# -*- coding=utf-8 -*-

import os
import sys
import time
import shutil
import urllib, urllib2
import random
import profile
import pstats
from pstats import Stats
import xlsxwriter
from analysisdata import ExcelReader
from analysisdata import AnalysierBase


sortby = "ncalls"



def f8(x):
    return "%10.8f" % x


def func_std_string(func_name):  # match what old profile produced
    if func_name[:2] == ('~', 0):
        # special case for built-in functions
        name = func_name[2]
        if name.startswith('<') and name.endswith('>'):
            return '{%s}' % name[1:-1]
        else:
            return name
    else:
        return "%s:%d(%s)" % func_name


def func_get_function_name(func):
    return func[2]


class profiletoexcel(Stats):
    def __init__(self, *args, **kwds):
        Stats.__init__(self, *args, **kwds)

        self.workbook = xlsxwriter.Workbook(args[0] + ".xlsx")
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.set_column('A:F', 15)
        self.worksheet.set_column('G:H', 90)

    def print_stats(self, *amount):

        # print >> self.stream,"#################"

        for filename in self.files:
            print >> self.stream, filename
        if self.files: print >> self.stream
        indent = ' ' * 2
        for func in self.top_level:
            print >> self.stream, indent, func_get_function_name(func), "level func"

        print >> self.stream, indent, self.total_calls, "function calls",
        if self.total_calls != self.prim_calls:
            print >> self.stream, "(%d primitive calls)" % self.prim_calls,
        print >> self.stream, "in %.3f seconds" % self.total_tt
        print >> self.stream
        width, list = self.get_print_list(amount)
        if list:
            self.print_title()
            for func in list:
                self.print_line(func)
            print >> self.stream
            print >> self.stream  # write data in to the pstats

            self.print_title2()
            for index, func in enumerate(list):
                ccalls, ncalls, totaltime, percall1, cumtime, percall2, funcinfo, callers = self.get_line(func)
                #print ncalls,totaltime,percall1,cumtime,percall2,funcinfo,callers
                self.worksheet.write(index + 1, 0, ccalls)
                self.worksheet.write(index + 1, 1, ncalls)
                self.worksheet.write(index + 1, 2, totaltime)
                self.worksheet.write(index + 1, 3, percall1)
                self.worksheet.write(index + 1, 4, cumtime)
                self.worksheet.write(index + 1, 5, percall2)
                self.worksheet.write(index + 1, 6, funcinfo)
                self.worksheet.write(index + 1, 7, str(callers))

        return self

    def print_title2(self):
        # print 'ncalls  tottime  percall  cumtime  percall  filename:lineno(function)',
        self.worksheet.write(0, 0, "ccalls")
        self.worksheet.write(0, 1, "ncalls")
        self.worksheet.write(0, 2, "totaltime")
        self.worksheet.write(0, 3, "totaltime/ncalls")
        self.worksheet.write(0, 4, "cumtime")
        self.worksheet.write(0, 5, "cumtime/ccalls")
        self.worksheet.write(0, 6, "filename:lineno(function)")
        self.worksheet.write(0, 7, "callers")

    def get_line(self, func):  # hack : should print percentages
        cc, nc, tt, ct, callers = self.stats[func]
        # print cc, nc, tt, ct, callers
        ncalls = str(nc)
        ccalls = str(cc)

        totaltime = f8(tt)
        if nc == 0:
            percall1 = " "
        else:
            percall1 = f8(float(tt) / nc)
        cumtime = f8(ct)
        if cc == 0:
            percall2 = " "
        else:
            percall2 = f8(float(ct) / cc)

        funcinfo = func_std_string(func)

        return ccalls, ncalls, totaltime, percall1, cumtime, percall2, funcinfo, callers


# def reloadProfileFile():
# # 将原始的profile文件转储到txt文件中，留待解析操作。
#     #原理为，将原来的标准输出重定向到pstats的输出流中。
#
#     newfile = open("pstats", "w+")
#     stats = pstats.Stats("test_prof_3v3_60s", stream=newfile)
#     ty = "cumulative"
#     stats.strip_dirs().sort_stats(ty).print_stats()
#     newfile.close()
#
#     pass
class exportAnalysis():
    def __init__(self):

        self.inputfilelist = []
        self.outputfilelist = []
        self.makefilenameExport()
        self.readfilenameExport()

        pass
    def makefilenameExport(self):

        path = os.path.abspath(os.path.dirname(sys.argv[0]))
        print path
        pathinput = path + "\\input_stuck\\"
        pathoutput = path + "\\output_pstats\\"
        for root,dirs,files in os.walk(pathinput):
            for filename in files:
                self.inputfilelist.append(pathinput+str(filename))
                self.outputfilelist.append(pathoutput+str(filename))

        pass

    def readfilenameExport(self):
        er = ExcelReader("filenameExport.xlsx")
        er.getData2()

        # idct = {u'strategy1':
        #             {
        #                 1: {u'targetfilename': u'test_prof_3v3.xlsx', u'basefilename': u'test_prof_3v3.xlsx', u'id': 1.0}
        #             },
        #         u'strategy2':
        #             {
        #                 1: {u'basefilename': u'test_prof_3v3.xlsx', u'id': 1.0}
        #             },
        #         u'profilefilename':
        #             {
        #                 1: {u'ID': 1.0, u'profilefilename': u'test_prof_3v3.xlsx'}
        #             }
        # }

        self.profilefilenamesheet = self.inputfilelist
        self.strategy1sheet = er.getData2().get(u'strategy1')
        self.strategy2sheet = er.getData2().get(u'strategy2')

        #print self.strategy1sheet
        self.profilefilenames = []
        for  value in self.profilefilenamesheet:
            self.profilefilenames.append(value)

        self.strategy1 = []
        for (key, value) in self.strategy1sheet.items():
            #line = [value.get(u'basefilename'),value.get(u'targetfilename'),value.get(u'outputfilename')]
            self.strategy1.append(value)  # a dict inside it

        self.strategy2 = []
        for (key, value) in self.strategy2sheet.items():
            self.strategy2.append(value.get(u'basefilename'))

            #print self.profilefilenames
            #print self.strategy1
            #print self.strategy2
            #print self.profilefilenamesheet,self.stragegy1sheet,self.stragegy2sheet


    def reloadtoexcle(self):
        #增加读取文件导表模块，实现文件名可配置。
        try:
            for filenameindex, filename in enumerate(self.profilefilenames):
                newfilename = self.profilefilenames[filenameindex] + ".pstats"
                newfilename = newfilename.replace("input_stuck","output_pstats")
                newfile = open(newfilename, "w+")
                stats = profiletoexcel(filename, stream=newfile)
                stats.sort_stats(sortby).print_stats()
                newfile.close()
        except Exception, e:
            print str(e)


    def analysisdata(self):
        try:
            for oneline in self.strategy1:
                ab = AnalysierBase(oneline.get(u'basefilename'), oneline.get(u'targetfilename'),
                                   oneline.get(u'outputfilename'))
                ab.datatoexcel()
        except Exception, e:
            print str(e)

    def changeexcelfile(self):
        path = os.path.abspath(os.path.dirname(sys.argv[0]))
        print path
        pathinput = path + "\\input_stuck\\"
        pathexceloutput = path + "\\output_excel\\"
        for root,dirs,files in os.walk(pathinput):
            for filename in files:
                if ".xlsx" in filename:
                    try:

                        shutil.copy(pathinput+str(filename),pathexceloutput+str(filename))
                        os.remove(pathinput+str(filename))
                    except Exception, e:
                        os.remove(pathinput+str(filename))
                        print str(e)



if __name__ == "__main__":
    print "reading the export excel......"
    ea = exportAnalysis()
    print "export all base excel data......"


    ea.reloadtoexcle()
    print "compare......"
    ea.analysisdata()
    print "success.."
    time.sleep(3)
    ea.changeexcelfile()


    # redmine_remind()
    #print getUndoneReqList()
    #stats = pstats.Stats("test_prof_3v3_60s")  #构造函数，用来构造Stats对象
    #calls, cumulative, file, line, module, name, nfl, pcalls, stdname, time
    #strip = stats.strip_dirs()# 出去文件名前面的路径信息
    ##sort  = strip.sort_stats("stdname")# 对strip进行排序

    #stats.print_stats()
    #print stats.files  # 得到文件信息
    #print stats.top_level

    ######reloadtoexcle()

    #readfilenameExport()
    pass


    # strip_dirs()                      用以除去文件名前名的路径信息。
    # add(filename,[…])                 把profile的输出文件加入Stats实例中统计
    # dump_stats(filename)              把Stats的统计结果保存到文件
    # sort_stats(key,[…])               最重要的一个函数，用以排序profile的输出
    # reverse_order()                   把Stats实例里的数据反序重排
    # print_stats([restriction,…])      把Stats报表输出到stdout
    # print_callers([restriction,…])    输出调用了指定的函数的函数的相关信息
    # print_callees([restriction,…])    输出指定的函数调用过的函数的相关信息


    # ‘ncalls’                  被调用次数
    # ‘cumulative’              函数运行的总时间
    # ‘file’                    文件名
    # ‘module’                  文件名
    # ‘pcalls’                  简单调用统计（兼容旧版，未统计递归调用）
    # ‘line’                    行号
    # ‘name’                    函数名
    # ‘nfl’                     Name/file/line
    # ‘stdname’                 标准函数名
    # ‘time’                    函数内部运行时间（不计调用子函数的时间）


    # a = []
    # str.print_title()
    # str.print_stats(2)
    #stats.dump_stats("test1")

