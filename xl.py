# -*- coding: utf-8 -*-
'''
Created on Feb 6, 2013

@author: Hugo
'''
import xlrd, xlwt
from xlutils.copy import copy
import os
from decimal import Decimal

#xlrd.Book.encoding = "gbk"
location = './/Data/'
location_ws = './/Data2/'
stand_cells_conf = './conf/standcells.conf'
cell_cells_conf = './conf/cell_cells.conf'
ws_cells_conf = './conf/ws_cells.conf'
cell_custom_conf = './conf/cell_customquery.conf'
ws_custom_conf = './conf/ws_customquery.conf'

years = []
years_ws = []
dep_list = os.listdir(location)
for year in os.listdir(location + '\\' + dep_list[0]):
    years.append(year[0:-4])

dep_list_ws = os.listdir(location_ws)
for year in os.listdir(location + '\\' + dep_list[0]):
    years_ws.append(year[0:-4])
    
# Create a font to use with the style
style = xlwt.XFStyle()

font0 = xlwt.Font()
font0.name = u'微软雅黑'
font0.colour_index = 0
font0.bold = True

pattern = xlwt.Pattern()  # Create the Pattern
pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
pattern.pattern_fore_colour = 2  # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...

style.pattern = pattern  # Add Pattern to Style
style.font = font0


#Init Standard Query report        
if os.path.exists('./report/Stand_report.xls'):
    os.remove('./report/Stand_report.xls')
worksheet_stand_report = xlwt.Workbook(encoding='gbk')
worksheet_stand_report.add_sheet(u'第一表', cell_overwrite_ok=True)
report_table = worksheet_stand_report.get_sheet(0)
report_table.write(1, 0, u'第一表')
report_table.write(3, 0, u'单位')                        
worksheet_stand_report.save('./report/Stand_report.xls')


#Init Custom Query Report
if os.path.exists('./report/Cell_Custom_report.xls'):
    os.remove('./report/Cell_Custom_report.xls')
cell_custom_report = xlwt.Workbook(encoding='gbk')
cell_custom_report.add_sheet(u'默认表', cell_overwrite_ok=True)
cell_custom_report.save('./report/Cell_Custom_report.xls')

#if os.path.exists('./report/WS_Custom_report.xls'):
#    os.remove('./report/WS_Custom_report.xls')
#ws_custom_report = xlwt.Workbook(encoding='gbk')
#ws_custom_report.add_sheet(u'默认表', cell_overwrite_ok=True)
#ws_custom_report.save('./report/WS_Custom_report.xls')


#Init index of cells.conf
def index_init(filename):
    index = {}
    f = file(filename, 'r')
    #t = len(f.readlines())
    i = 0
    #f.readline(100)
    first = f.readline().split('.')
    index[first[0].decode('gbk')] = 0
    
    while True:    
        record = f.readline()
    #    print record
        if len(record) == 0:
            break
        else:
            if index.has_key(record.split('.')[0].decode('gbk')):
    #            print i, 'same', record.decode('gbk')
                i += 1 
            else:
    #            print i, 'diff', record.decode('gbk')
                index[record.split('.')[0].decode('gbk')] = i + 1
                i += 1    
    f.close
    return index


#Get_customquery cells
def get_customrecords(xy):
#    a = '测试:2011.第一表.A1+2012.第二表.A2-2011.第二表.A3+2011.第二表.A4-2011.第二表.A5-2011.第二表.A6+2011.第二表.A7'
#    print xy
    record = xy.split(':')
    record_xy = [record[0]]
    cell = record[1].split('+')
    for a in cell:
#        print a
        if  a.find('-') != -1:            
#            print 'find'
            i = 0            
            for b in a.split('-'):
                record_xy.append(b)                
                if i != len(a.split('-')) - 1:
                    record_xy.append('-')
                else:
                    record_xy.append('+')
                i += 1
        else:
#            print 'no found'
            record_xy.append(a)
            record_xy.append('+')
    
#    for i in record_xy[0:-1]:
#        print  'dsa=', i
    return record_xy[0:-1]

    
#Get cells physical location of Excel.

def get_standcell_xy(xy):
#    try:
        result = []
        record = xy.split('=')
        
        query_id = record[0].split('.')[1]
        query_name = record[3]
#print phyical location of record    
        for lens in range(1, len(record) - 1):
            table_cell = record[lens].split('->')            
            cell = table_cell[1].split(':')        
            cell_x = ord(cell[0]) - 65
            cell_y = cell[1]
            result.append(xy.split('.')[0])
            result.append(int(table_cell[0]) - 1)
            result.append(int(cell_x))
            result.append(int(cell_y) - 1)
        result.append(query_id)
        result.append(query_name)
#        print 'aaaa', result
        return result
#    except:
#        print 'Get cells xy failed'

def get_allcells_xy(xy):
#    try:
        result = []
        record = xy.split('=')
        query_name = record[0].split('.')[0]

#print phyical location of record    
        for lens in range(1, len(record)):
            table_cell = record[lens].split('->')            
            cell = table_cell[1].split(':')        
            cell_x = ord(cell[0]) - 65
            cell_y = cell[1]
#            print cell, cell_x, cell_y, int(table_cell[0]) - 1
            result.append(query_name)
            result.append(int(table_cell[0]) - 1)
            result.append(int(cell_x))
            result.append(int(cell_y) - 1)
#            print result
        return result
#    except:
#        print 'Get cells xy failed'        

def get_customcells_xy(xy):
#    try:
        result = []
        record = xy.split('=')
        query_name = record[0].split('.')[0]

#print phyical location of record    
        for lens in range(1, len(record)):
            table_cell = record[lens].split('->')            
            cell = table_cell[1].split(':')        
            cell_x = ord(cell[0]) - 65
            cell_y = cell[1]
#            print cell, cell_x, cell_y, int(table_cell[0]) - 1
            result.append(query_name)
            result.append(int(table_cell[0]) - 1)
            result.append(int(cell_x))
            result.append(int(cell_y) - 1)
#            print result
        return result
#    except:
#        print 'Get cells xy failed'        
        
        
#Load standard query configure from standcells.conf

def StandQuery(*query_content):
     
# Create a font to use with the style
    font0 = xlwt.Font()
    font0.name = u'微软雅黑'
    font0.colour_index = 2
    font0.bold = True
    style = xlwt.XFStyle()
    style.font = font0
    

    for name_id in (query_content):
        f = file(stand_cells_conf, 'r')
        length = 0
        while True:        
            line = f.readline()
            
            if len(line) == 0:
                break
            if line.decode('gbk').split('=')[0].startswith(name_id) != 1:
#                print 'diff'
                pass
            else:
#                print 'same'
                record = get_standcell_xy(line.decode('gbk'))
#                print record
                dep_list = os.listdir(location)
                length += 4
#                print length
    #Write data into report.xls                
#                try:
                report_xls = xlrd.open_workbook('./report/Stand_report.xls')
                
                for k in range(0, len(report_xls.sheet_names())):
#                    print k, len(report_xls.sheet_names())
#                    print record[0], report_xls.sheet_names()[k].decode('utf-8') 
                    if record[0] == report_xls.sheet_names()[k]  :
    
                        table = report_xls.sheet_by_name(record[0])
    
                        for i in range(0, len(dep_list)):
                            data_source_before = location + dep_list[i] + '\\' + years[0] + '.xls'
                            data_source_after = location + dep_list[i] + '\\' + years[1] + '.xls'
    
                            worksheet_before = xlrd.open_workbook(data_source_before)
                            worksheet_after = xlrd.open_workbook(data_source_after)
            
                            table_before = worksheet_before.sheet_by_index(int(record[1]))
                            table_after = worksheet_after.sheet_by_index(int(record[5]))
    
                            a = table_before.cell(record[3], record[2]).value
                            b = table_after.cell(record[7], record[6]).value
                            if a == '':
                                a = a.replace('', '0')
                            if b == '':
                                b = b.replace('', '0')
                            report_table = worksheet_stand_report.get_sheet(k)
                            
#Deal with ncols >250
                            if length <= 16:
                                length_y = table.ncols                                                                     
                                report_table.write(2 , 0 + length_y , record[8] + '-' + record[9])                            
                                report_table.write(3 , 0 + length_y , years[0])
                                report_table.write(3 , 1 + length_y , years[1])
                                report_table.write(3 , 2 + length_y , u'合计', style)
                                report_table.write(4 + i , 0, dep_list[i])
                                report_table.write(4 + i , 0 + length_y, int(a))
                                report_table.write(4 + i , 1 + length_y, int(b))
                                report_table.write(4 + i , 2 + length_y, int(b) - int(a))
                                report_table.write(4 + i , 3 + length_y, ' ')
                                worksheet_stand_report.save('./report/Stand_report.xls')
#                                print length, 'oooo'
                            elif length % 16 > 0:                                
                                length_x = (length / 16) * (len(dep_list) + 3) 
                                length_y = length % 16 - 3
#                                print length, length_x, length_y
                                
                                report_table.write(2 + length_x, 0 + length_y, record[8] + '-' + record[9])                            
                                report_table.write(3 + length_x, 0 + length_y, years[0])
                                report_table.write(3 + length_x, 1 + length_y, years[1])
                                report_table.write(3 + length_x, 2 + length_y, u'合计', style)
                                report_table.write(4 + i + length_x , 0, dep_list[i])
                                report_table.write(4 + i + length_x, 0 + length_y, int(a))
                                report_table.write(4 + i + length_x, 1 + length_y, int(b))
                                report_table.write(4 + i + length_x, 2 + length_y, int(b) - int(a))
                                report_table.write(4 + i + length_x, 3 + length_y, ' ')
                                worksheet_stand_report.save('./report/Stand_report.xls')
                            else:
#                                print 'end of line'
                                length_x = (length / 16 - 1) * (len(dep_list) + 3)
                                length_y = 16 - 3
#                                print length, length_x, length_y
                                
                                report_table.write(2 + length_x, 0 + length_y, record[8] + '-' + record[9])                            
                                report_table.write(3 + length_x, 0 + length_y, years[0])
                                report_table.write(3 + length_x, 1 + length_y, years[1])
                                report_table.write(3 + length_x, 2 + length_y, u'合计', style)
                                report_table.write(4 + i + length_x , 0, dep_list[i])
                                report_table.write(4 + i + length_x, 0 + length_y, int(a))
                                report_table.write(4 + i + length_x, 1 + length_y, int(b))
                                report_table.write(4 + i + length_x, 2 + length_y, int(b) - int(a))
                                report_table.write(4 + i + length_x, 3 + length_y, ' ')
                                worksheet_stand_report.save('./report/Stand_report.xls')
                                
                        
                        break
                
                    elif k == int(len(report_xls.sheet_names()) - 1):
                        worksheet_stand_report.add_sheet(record[0], cell_overwrite_ok=True)
                        report_table = worksheet_stand_report.get_sheet(k + 1)
                        report_table.write(2 , 1 , record[8] + '-' + record[9])
                        report_table.write(1, 0, record[0])
                        report_table.write(3, 0, u'单位')                        
                        report_table.write(3, 1, years[0])
                        report_table.write(3, 2, years[1])
                        report_table.write(3, 3, u'合计', style)

                        for i in range(0, len(dep_list)):
                            data_source_before = location + dep_list[i] + '\\' + years[0] + '.xls'
                            data_source_after = location + dep_list[i] + '\\' + years[1] + '.xls'
                            
                            worksheet_before = xlrd.open_workbook(data_source_before)
                            worksheet_after = xlrd.open_workbook(data_source_after)
            
                            table_before = worksheet_before.sheet_by_index(int(record[1]))
                            table_after = worksheet_after.sheet_by_index(int(record[5]))
                            
                            a = table_before.cell(record[3], record[2]).value
                            b = table_after.cell(record[7], record[6]).value
                            if a == '':
                                a = a.replace('', '0')
                            if b == '':
                                b = b.replace('', '0')
                            
                            report_table = worksheet_stand_report.get_sheet(k + 1)
                            report_table.write(4 + i , 1 , int(a))
                            report_table.write(4 + i , 2 , int(b))
                            report_table.write(4 + i , 3 , int(b) - int(a))
                            report_table.write(4 + i , 4 , ' ')
                            report_table.write(4 + i, 0, dep_list[i])
                            worksheet_stand_report.save('./report/Stand_report.xls')
    
#                except:
#                    print 'Some error occurred'
                            
                worksheet_stand_report.save('./report/Stand_report.xls')
    f.close()
#    except:
#        print 'Stand Query failed'
#Define Custom Query Model

def CellCustomQuery(query_content):

# Create a font to use with the style
    font0 = xlwt.Font()
    font0.name = u'微软雅黑'
    font0.colour_index = 2
    font0.bold = True
    style = xlwt.XFStyle()
    style.font = font0

    record_name = get_customrecords(query_content)[0]
    record_custom = get_customrecords(query_content)[1:]
#    print record_name, record_custom
    table_exist = 0
    custom_report = xlrd.open_workbook('./report/Cell_Custom_report.xls', formatting_info=True)
    for table_exist in range(0, len(custom_report.sheet_names())):
#        print table_exist, record_name, custom_report.sheet_names()[table_exist]
        if custom_report.sheet_names()[table_exist] == record_name:
            write_custom(record_custom, table_exist)
            break
        elif table_exist == len(custom_report.sheet_names()) - 1:
            cell_custom_report.add_sheet(record_name, cell_overwrite_ok=True)
            cell_custom_report.save('./report/Cell_Custom_report.xls')
            write_custom(record_custom, len(custom_report.sheet_names()))
            break
#        table_exist += 1      
        
        
#
def WSCustomQuery(v, xy, year, threshold):
    
    index = index_init(ws_cells_conf)
    record_custom = get_customrecords(xy)[1:]
    record_name = get_customrecords(xy)[0]
#    print record_name.decode('utf-8')
    for i in range(0, len(dep_list_ws)):
    
        datasource = location_ws + dep_list_ws[i].decode('gbk') + '//' + year + '.xls'
        if v == 0:
#            print dep_list_ws[i].decode('gbk')
            ws_report_xls = xlrd.open_workbook(datasource, formatting_info=True)
            ws_custom_xls = copy(ws_report_xls)
        else:
            ws_report_xls = xlrd.open_workbook('./report/' + dep_list_ws[i].decode('gbk') + '_' + record_name + '_WS_custom_report.xls', formatting_info=True)
            ws_custom_xls = copy(ws_report_xls)
            
        for j in range(0, len(record_custom), 2):
#            print j
            f = file(ws_cells_conf, 'r')
            
            for line in f.readlines()[index[record_custom[j]]:]:
                f1 = file(ws_cells_conf, 'r')
                for k in f1.readlines()[index[ record_custom[len(record_custom) - 1]]:]:
#                        print record_custom[len(record_custom) - 1] + '.' + line.split('=')[0].split('.')[1]
                    
                    if k.decode('gbk').startswith(record_custom[len(record_custom) - 1] + '.' + line.split('=')[0].split('.')[1]):
                        record_result = get_allcells_xy(k)
#                        print k.decode('gbk'), record_result
                        break
                    else:
                        continue
                f1.close()
                                
                if len(line) == 0:
                    break                
                
                if line.decode('gbk').split('=')[0].startswith(record_custom[j]) != 1:                    
#                    print 'diff'
                    break
                else:
#                    print 'same'
                    record = get_allcells_xy(line)                    
#                    print record_custom[j], record                
                    
                    ws_worksheet = xlrd.open_workbook(datasource)
                    table = ws_worksheet.sheet_by_index(record[1] + 1)
                    b = table.cell(record[3], record[2]).value
                    if b == '':
                        b = b.replace('', '0')
                    elif b == '-':
                        b = b.replace('-', '0')
#                    print b
                    ws_report_table = ws_custom_xls.get_sheet(len(ws_worksheet.sheet_names()) - 1)
#                    ws_report_table.get
                    if j == 0:
                        result = int(b)
                        ws_report_table.write(record_result[3], record_result[2], result)
                        ws_custom_xls.save('./report/' + dep_list_ws[i].decode('gbk') + '_' + record_name + '_WS_custom_report.xls')
                    elif j < len(record_custom) :
                        ws_result = xlrd.open_workbook('./report/' + dep_list_ws[i].decode('gbk') + '_' + record_name + '_WS_custom_report.xls')
                        table_result = ws_result.sheet_by_index(len(ws_result.sheet_names()) - 1)
                        c = table_result.cell(record_result[3], record_result[2]).value
#                        print 'b-c:' , b, c
                        if c == '':
                            c = c.replace('', '0')
                        elif b == '-':
                            c = c.replace('-', '0')
                        
                        if record_custom[j - 1] == '+':
                            
                            result = int(b) + int(c)
                        else:
                    
                            result = int(c) - int(b)
                            
                        if result > threshold:                            
                            ws_report_table.write(record_result[3], record_result[2], result, style)
                            ws_custom_xls.save('./report/' + dep_list_ws[i].decode('gbk') + '_' + record_name + '_WS_custom_report.xls')
                        else:
                            ws_report_table.write(record_result[3], record_result[2], result)
                            ws_custom_xls.save('./report/' + dep_list_ws[i].decode('gbk') + '_' + record_name + '_WS_custom_report.xls')
#                ws_custom_xls.save('./report/' + dep_list_ws[i].decode('gbk') + '_WS_custom_report.xls')
        
                
def write_custom(record_custom, table_index):        
    report_table_custom = cell_custom_report.get_sheet(table_index)
    report_table_custom.write(1, 1 + len(record_custom) / 2 + 1, u'合计')
#    print record_custom
    for k in range(0, len(dep_list)):
#        print k
        result = 0
        report_table_custom.write(2 + k, 0, dep_list[k])
        for i in range(0, len(record_custom) + 1, 2):
#            print i
            f = file(cell_cells_conf, 'r')
            while True:                            
                line = f.readline()
                if len(line) == 0:
                    break
#                print line.decode('gbk').split(':')[0], record_custom[i][5:].replace('\n', '') + '='
                if line.decode('gbk').split(':')[0].startswith(record_custom[i][5:].replace('\n', '') + '=') != 1:
#                    print 'diff'
                    pass
                else:
#                    print 'same'
                    record = get_customcells_xy(line.decode('gbk'))           
#                    print record
                    data_source = location + dep_list[k].decode('gbk') + '\\' + record_custom[i][0:4] + '.xls'
                    worksheet_custom = xlrd.open_workbook(data_source)
                    table_after = worksheet_custom.sheet_by_index(record[1])
                    b = table_after.cell(record[3], record[2]).value
#                    print b                    
                    report_table_custom.write(1, 1 + i / 2, record_custom[i])
                    report_table_custom.write(2 + k, 1 + i / 2, b)
                    cell_custom_report.save('./report/Cell_Custom_report.xls')
                    if b == '':
                        b = b.replace('', '0')
#                    print i
                    if i == 0:
                        result = b
#                        print b, result
                    elif i < len(record_custom) :
                            if record_custom[i - 1] == '+':
                                result += int(b)
                            else:
                                result -= int(b)
#                            print b, record_custom[i - 1], result
                    
                    report_table_custom.write(2 + k, 1 + len(record_custom) / 2 + 1, result)
                    cell_custom_report.save('./report/Cell_Custom_report.xls')
                    break                    
                    
            f.close()
#            break            
        cell_custom_report.save('./report/Cell_Custom_report.xls')


def PercentQuery(v, check_percent, xy):
    
    index = index_init(cell_cells_conf)
    
    for i in range(0, len(dep_list)):
        t = 0
#        print location + dep_list[i].decode('gbk') + '\\' + years[0] + '.xls'
        data_source_before = location + dep_list[i] + '//' + years[0] + '.xls'
        data_source_after = location + dep_list[i] + '//' + years[1] + '.xls'
        if v == 0:
            report_xls = xlrd.open_workbook(data_source_after, formatting_info=True)
            per_report_xls = copy(report_xls)
        else:
            report_xls = xlrd.open_workbook('./report/' + dep_list[i] + '_Percent_report.xls', formatting_info=True)
            per_report_xls = copy(report_xls)
            
        worksheet_before = xlrd.open_workbook(data_source_before)
        worksheet_after = xlrd.open_workbook(data_source_after)
        
        f = file(cell_cells_conf, 'r')
#        while False:
        for line in f.readlines()[index[xy]:]:
#            print t
#            line = f.readline()
            if len(line) == 0:
                break
#            print line.decode('gbk').split('=')[0], xy
            if line.decode('gbk').split('=')[0].startswith(xy) != 1:
    #            print 'diff'
                break
            else:
    #            print 'same'
                record = get_allcells_xy(line)
#                print record

                table_before = worksheet_before.sheet_by_index(int(record[1]))
                table_after = worksheet_after.sheet_by_index(int(record[1]))
        
                a = table_before.cell(record[3], record[2]).value
                b = table_after.cell(record[3], record[2]).value
#                print a, b
                if a == '':
                    a = a.replace('', '0')
                elif a == '-':
                    a = a.replace('-', '0')
                if b == '':
                    b = b.replace('', '0')
                elif b == '-':
                    b = b.replace('-', '0')
                
#                print a, b, int(b) - int(a)
                report_table = per_report_xls.get_sheet(record[1])
                if int(a) == 0:
                    if int(b) != 0:
                        report_table.write(record[3], record[2], 100, style)
                        per_report_xls.save('./report/' + dep_list[i] + '_Percent_report.xls')
                    else:
                        report_table.write(record[3], record[2], 0)
                        per_report_xls.save('./report/' + dep_list[i] + '_Percent_report.xls')
                elif (int(b) - int(a)) > 0:
                    b_big_a = round(Decimal(int(b) - int(a)) / Decimal(a) * 100, 2)
                
                    if b_big_a >= check_percent:    
                        report_table.write(record[3], record[2], b_big_a, style)
                        per_report_xls.save('./report/' + dep_list[i] + '_Percent_report.xls')
                    else:
                        report_table.write(record[3], record[2], b_big_a)
                        per_report_xls.save('./report/' + dep_list[i] + '_Percent_report.xls')
                elif (int(b) - int(a)) < 0:
                    a_big_b = round(Decimal(int(a) - int(b)) / Decimal(a) * 100, 2)
                
                    if a_big_b >= check_percent:
                        report_table.write(record[3], record[2], -1 * a_big_b, style)
                        per_report_xls.save('./report/' + dep_list[i] + '_Percent_report.xls')
                    else:
                        report_table.write(record[3], record[2], -1 * a_big_b)
                        per_report_xls.save('./report/' + dep_list[i] + '_Percent_report.xls')
                else:
                    report_table.write(record[3], record[2], 0)
                    per_report_xls.save('./report/' + dep_list[i] + '_Percent_report.xls')
            t += 1            
        f.close()




