# -*- coding: utf-8 -*-
'''
Created on Feb 6, 2013

@author: Hugo
'''

import xlrd, xlwt
#import sys
import os

#1xlrd.Book.encoding = "gbk"
location = '.\\ECL\\Data\\'
years = ['2011', '2012']




#Get cells physical location of Excel.
def get_cell_xy(xy):
    try:
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
        
        return result
    except:
        print 'Get cells xy failed'
##Load cells configure from cells.conf
def read_cells_conf():
    
    try:
        f = file('./ECL/cells.conf', 'r')
        while True:
            line = f.readline()
            if len(line) == 0:
                break
            record = get_cell_xy(line)
            
        f.close()
    except:
        print 'Read cells configure failed'
    
#Load standard query configure from standcells.conf
def stand_query():

        
        if os.path.exists('./report.xls'):
            os.remove('./report.xls')
        worksheet_report = xlwt.Workbook(encoding='gbk')
         
# Create a font to use with the style
        font0 = xlwt.Font()
        font0.name = u'微软雅黑'
        font0.colour_index = 2
        font0.bold = True
        style = xlwt.XFStyle()
        style.font = font0
        
#Init report.xls        
        worksheet_report.add_sheet(u'第一表', cell_overwrite_ok=True)
        report_table = worksheet_report.get_sheet(0)
        report_table.write(1, 0, u'第一表')
        report_table.write(3, 0, u'单位')                        
        worksheet_report.save('report.xls')
        
        
        query_content = ['第一表.A1', '第6表']
        for name_id in (query_content):
            f = file('./ECL/standcells.conf', 'r')
            length = 0
            while True:        
                line = f.readline()
                
                if len(line) == 0:
                    break
                print name_id, line.decode('gbk').split('=')[0].encode('utf-8')
                print len(name_id), len(line.decode('gbk').split('=')[0].encode('utf-8'))
                
                if line.decode('gbk').split('=')[0].encode('utf-8').startswith(name_id) != 1:
                    print 'diff'
#                    pass
                else:
                    print 'same'
                    record = get_cell_xy(line)
#                    print record
                    dep_list = os.listdir(location)
                    length += 4
                    print length
        #Write data into report.xls                
    #                try:
                    report_xls = xlrd.open_workbook('./report.xls')
                    
                    for k in range(0, len(report_xls.sheet_names())):
    #                    print k, len(report_xls.sheet_names())
                        if record[0].decode('gbk').encode('utf-8') == report_xls.sheet_names()[k].decode('utf-8').encode('utf-8')  :
        
                            table = report_xls.sheet_by_name(record[0].decode('gbk').encode('utf-8'))
        
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
                                report_table = worksheet_report.get_sheet(k)
                                
    #Deal with ncols >250
                                if length <= 12:
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
                                    worksheet_report.save('report.xls')
                                    print length, 'oooo'
                                elif length % 12 > 0:                                
                                    length_x = (length / 12) * (len(dep_list) + 3) 
                                    length_y = length % 12 - 3
                                    print length, length_x, length_y
                                    
                                    report_table.write(2 + length_x, 0 + length_y, record[8] + '-' + record[9])                            
                                    report_table.write(3 + length_x, 0 + length_y, years[0])
                                    report_table.write(3 + length_x, 1 + length_y, years[1])
                                    report_table.write(3 + length_x, 2 + length_y, u'合计', style)
                                    report_table.write(4 + i + length_x , 0, dep_list[i])
                                    report_table.write(4 + i + length_x, 0 + length_y, int(a))
                                    report_table.write(4 + i + length_x, 1 + length_y, int(b))
                                    report_table.write(4 + i + length_x, 2 + length_y, int(b) - int(a))
                                    report_table.write(4 + i + length_x, 3 + length_y, ' ')
                                    worksheet_report.save('report.xls')
                                else:
                                    print 'end of line'
                                    length_x = (length / 12 - 1) * (len(dep_list) + 3)
                                    length_y = 12 - 3
                                    print length, length_x, length_y
                                    
                                    report_table.write(2 + length_x, 0 + length_y, record[8] + '-' + record[9])                            
                                    report_table.write(3 + length_x, 0 + length_y, years[0])
                                    report_table.write(3 + length_x, 1 + length_y, years[1])
                                    report_table.write(3 + length_x, 2 + length_y, u'合计', style)
                                    report_table.write(4 + i + length_x , 0, dep_list[i])
                                    report_table.write(4 + i + length_x, 0 + length_y, int(a))
                                    report_table.write(4 + i + length_x, 1 + length_y, int(b))
                                    report_table.write(4 + i + length_x, 2 + length_y, int(b) - int(a))
                                    report_table.write(4 + i + length_x, 3 + length_y, ' ')
                                    worksheet_report.save('report.xls')
                                    
                            
                            break
                    
                        elif k == int(len(report_xls.sheet_names()) - 1):
                            worksheet_report.add_sheet(record[0], cell_overwrite_ok=True)
                            report_table = worksheet_report.get_sheet(k + 1)
                            report_table.write(2 , 1 , record[8] + '-' + record[9])
                            report_table.write(1, 0, record[0])
                            report_table.write(3, 0, u'单位')                        
                            report_table.write(3, 1, years[0])
                            report_table.write(3, 2, years[1])
                            report_table.write(3, 3, u'合计', style)
                            report_table.write(4 + i, 0, dep_list[i])
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
                                
                                report_table = worksheet_report.get_sheet(k + 1)
                                report_table.write(4 + i , 1 , int(a))
                                report_table.write(4 + i , 2 , int(b))
                                report_table.write(4 + i , 3 , int(b) - int(a))
                                report_table.write(4 + i , 4 , ' ')
                                worksheet_report.save('report.xls')
        
                                
        
    #                except:
    #                    print 'Some error occurred'
                                
                    worksheet_report.save('report.xls')

    
                    
        f.close()
#    except:
#        print 'Stand Query failed'

#read_cells_conf()
stand_query()



        
        
        
        
#    read_stand_query()

