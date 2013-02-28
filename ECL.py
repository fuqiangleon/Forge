# -*- coding: utf-8 -*-
#import data
import wx
import xl
import time
import os
import platform
import hashlib
import base64
import datetime





class MainPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        grid = wx.GridBagSizer(hgap=0, vgap=0)
        grid_sub = wx.GridBagSizer(hgap=0, vgap=0)
        vSizer = wx.BoxSizer(wx.VERTICAL)

        self.quote_all = wx.StaticText(self, label=u"所有记录 ", style=wx.ALIGN_CENTER, size=(120, 20))
        grid.Add(self.quote_all, pos=(1, 5))
        
        self.quote_selected = wx.StaticText(self, label=u"查询记录 ", style=wx.ALIGN_CENTER, size=(120, 20))
        grid.Add(self.quote_selected, pos=(1, 12))

        self.list_all = wx.ListBox(self, -1, size=(120, 250), choices='', style=wx.LB_EXTENDED)
        grid.Add(self.list_all, pos=(2, 5))
        
        self.list_selected = wx.ListBox(self, -1, size=(120, 250), style=wx.LB_EXTENDED)
        grid.Add(self.list_selected, pos=(2, 12))

        self.query_add = wx.Button(self, label=u'>>添加>>', size=(90, 30))
        grid_sub.Add(self.query_add, pos=(3, 5), flag=wx.ALIGN_TOP)
        self.Bind(wx.EVT_BUTTON, self.AddRecord, self.query_add)
#        grid_sub.Add((10, 5))
        self.query_del = wx.Button(self, label=u'<<删除<<', size=(90, 30))
        grid_sub.Add(self.query_del, pos=(5, 5), flag=wx.ALIGN_CENTER_VERTICAL)
        self.Bind(wx.EVT_BUTTON, self.DelRecord, self.query_del)

        self.query_reset = wx.Button(self, label=u'<<全选>>', size=(90, 30))
        grid_sub.Add(self.query_reset, pos=(7, 5), flag=wx.ALIGN_BOTTOM)
        self.Bind(wx.EVT_BUTTON, self.ResetRecord, self.query_reset)
        
        self.query_button = wx.Button(self, label=u'查询', size=(100, 50))
        grid.Add(self.query_button, pos=(2, 20), flag=wx.ALIGN_CENTER_VERTICAL)

        grid.Add(grid_sub, pos=(2, 6))
        vSizer.Add(grid, 1, wx.ALL, 0)
        
        self.SetSizerAndFit(vSizer)

            
    def AddRecord(self, event):       
        for v in self.list_all.GetSelections():
            self.list_selected.Append(self.list_all.GetString(v))
        
    def DelRecord(self, event):
        count = len(self.list_selected.GetSelections())
        for v in range(0, count):
#            print v
            self.list_selected.Delete(self.list_selected.GetSelections()[count - v - 1])

    def ResetRecord(self, event):
        v = 0
        if self.list_selected.GetCount() != 0:
            self.list_selected.Clear()
        else:
            try:
                while 1: 
                    self.list_selected.Append(self.list_all.GetString(v))
                    v += 1
            except:
                print 'error'

    
class StandQuery(MainPanel):
    def __init__(self, parent):
        MainPanel.__init__(self, parent)
        self.Bind(wx.EVT_BUTTON, self.StandQuery, self.query_button)
        self.GetRecord()
    def GetRecord(self):
        self.list_all.Clear()
        f = file(xl.stand_cells_conf, 'r')
        while True:        
            line = f.readline()                
            if len(line) == 0:
                break            
            if self.list_all.FindString(line.split('.')[0].decode('gbk')) == -1:
                
                self.list_all.Append(line.split('.')[0].decode('gbk'))
            

    def StandQuery(self, event):
        v = 0  
        self.gauge = wx.Gauge(self, -1, len(self.list_all.GetSelections()), (0, 315), (785, 20))
        try: 
            while True:
                xl.StandQuery(self.list_selected.GetString(v))                
                self.gauge.SetValue(v + 1)
                v += 1
        except:
            if v == 0:
                
                dlg_warning = wx.MessageDialog(self, u'请选择要查询的记录', u'选择错误！', wx.OK)
                dlg_warning.ShowModal()
            else:        
                self.gauge.SetValue(len(self.list_all.GetSelections() * 10))
                time.sleep(1)
                dlg_over = wx.MessageDialog(self, u'查询成功完成', '', wx.OK)
                dlg_over.ShowModal()
                self.gauge.SetValue(0)            
                self.gauge.Destroy()
            
                        
class CustomQuery(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        # create some sizers
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        grid = wx.GridBagSizer(hgap=6, vgap=6)
        hSizer = wx.BoxSizer(wx.HORIZONTAL)

        self.name_label = wx.StaticText(self, label=u"名称", size=(-1, -1))
        self.year_label = wx.StaticText(self, label=u"年份", size=(-1, -1))
        self.table_label = wx.StaticText(self, label=u"表名")
#        self.cell_label = wx.StaticText(self, label=u"单元名")
        self.table_expr = wx.StaticText(self, label=u"公式")
        self.table_all = wx.StaticText(self, label=u"所有记录")
        self.table_selected = wx.StaticText(self, label=u"查询记录")
        
        self.name_box = wx.TextCtrl(self, size=(100, -1))
        self.year_box = wx.Choice(self, size=(60, -1), choices=xl.years)
        self.table_box = wx.TextCtrl(self, size=(60, -1))
#        self.cell_box = wx.TextCtrl(self, size=(60, -1))
        self.expr_box = wx.Choice(self, size=(35, -1), choices=['+', '-', ''])
        
        self.button_add = wx.Button(self, label=u'添加', size=(50, -1))
        self.button_del = wx.Button(self, label=u'删除', size=(50, -1))
        self.button_reset = wx.Button(self, label=u'重置', size=(50, -1))
        self.button_save = wx.Button(self, label=u'保存', size=(50, -1))
        
        self.list_all = wx.ListBox(self, -1, size=(470, 130), choices='', style=wx.LB_SINGLE | wx.LB_HSCROLL)
        self.list_query = wx.ListBox(self, -1, size=(470, 100), choices='', style=wx.LB_MULTIPLE | wx.LB_HSCROLL)
        self.button_query = wx.Button(self, -1, size=(100, 50), label=u'查询', style=wx.ALIGN_CENTER_VERTICAL)
        self.button_selected = wx.Button(self, -1, size=(-1, -1), label=u'↓↓', style=wx.ALIGN_CENTER_VERTICAL)
        self.button_unselected = wx.Button(self, -1, size=(-1, -1), label=u'↑↑', style=wx.ALIGN_CENTER_VERTICAL)
        
        self.Bind(wx.EVT_BUTTON, self.Addrecord, self.button_add)
        self.Bind(wx.EVT_BUTTON, self.Delrecord, self.button_del)
        self.Bind(wx.EVT_BUTTON, self.Resetrecord, self.button_reset)
        self.Bind(wx.EVT_BUTTON, self.Saverecord, self.button_save)
        self.Bind(wx.EVT_BUTTON, self.Selected, self.button_selected)
        self.Bind(wx.EVT_BUTTON, self.Unselected, self.button_unselected)
        self.Bind(wx.EVT_BUTTON, self.Query, self.button_query)
        
        grid.Add(self.name_label, pos=(0, 1))
        grid.Add(self.year_label, pos=(0, 3))
        grid.Add(self.table_label, pos=(0, 5))
#        grid.Add(self.cell_label, pos=(0, 7))
        grid.Add(self.table_expr, pos=(0, 9))
        
        grid.Add(self.name_box, pos=(0, 2))
        grid.Add(self.year_box, pos=(0, 4))
        grid.Add(self.table_box, pos=(0, 6))
#        grid.Add(self.cell_box, pos=(0, 8))
        grid.Add(self.expr_box, pos=(0, 10))
        
        grid.Add(self.button_add, pos=(0, 11))
        grid.Add(self.button_del, pos=(0, 12))
        grid.Add(self.button_reset, pos=(0, 13))
        grid.Add(self.button_save, pos=(0, 14))
        
        grid.Add(self.list_all, pos=(1, 2), span=(1, 10))
        grid.Add(self.list_query, pos=(3, 2), span=(1, 10))
        grid.Add(self.table_all, pos=(1, 1))
        grid.Add(self.table_selected, pos=(3, 1))
        
        grid.Add(self.button_selected, pos=(2, 3), span=(1, 2), flag=wx.ALIGN_CENTER)
        grid.Add(self.button_unselected, pos=(2, 6), span=(1, 2), flag=wx.ALIGN_CENTER)
        
        grid.Add(self.button_query, pos=(3, 12), span=(1, 2), flag=wx.ALIGN_CENTER)
        
        hSizer.Add(grid, 0, wx.ALL, 5)
        mainSizer.Add(hSizer, 0, wx.ALL, 5)
#        mainSizer.Add(self.list_query, 0, wx.LEFT)
        self.SetSizerAndFit(mainSizer)
        self.Initrecord()
        
    def Initrecord(self):
        f = open(xl.cell_custom_conf, 'r')
        while True:
            line = f.readline()
#            print line.decode('gbk')
            if len(line) == 0:
                break
            else:
                self.list_all.Append(line.decode('utf-8'))
        f.close()
        
    def Addrecord(self, event):
        if self.name_box.GetValue() == '':
            dlg_over = wx.MessageDialog(self, u'名称不能为空', '')
            result = dlg_over.ShowModal()
        elif self.table_box.GetValue() == '':
            dlg_over = wx.MessageDialog(self, u'表名不能为空', '')
            result = dlg_over.ShowModal()
        elif self.name_box.GetValue() == '':
            dlg_over = wx.MessageDialog(self, u'单元名不能为空', '')
            result = dlg_over.ShowModal()            
        else:   
            record = self.year_box.GetStringSelection() + '.' + self.table_box.GetValue() + self.expr_box.GetStringSelection()
            if self.expr_box.GetStringSelection() == '':
                dlg_over = wx.MessageDialog(self, u'确定是最后一个单元名吗？', '')
                result = dlg_over.ShowModal()
                if result == wx.ID_OK:
    #                print record
                    if self.list_all.GetStringSelection() != '':
                        old = self.list_all.GetStringSelection()
                        self.list_all.Delete(self.list_all.GetSelection())
                        new = old + record
                        self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                        
                    else:
                        new = self.name_box.GetValue() + ':' + record
                        self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                else:
                    pass       
            else:        
#                print record
                if self.list_all.GetStringSelection() != '':
                    old = self.list_all.GetStringSelection()
                    self.list_all.Delete(self.list_all.GetSelection())
                    new = old + record
                    self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                else:
                    new = self.name_box.GetValue() + ':' + record
                    self.list_all.Insert(new, self.list_all.GetSelection() + 1)

    def Saverecord(self, event):
        v = 0
        f = open(xl.cell_custom_conf, 'w')
        try: 
            while True:
#                print 
                f.write(self.list_all.GetString(v))
                v += 1
        except:
            if v == 0:
                dlg_warning = wx.MessageDialog(self, u'无可保存的记录', '错误！', wx.OK)
                dlg_warning.ShowModal()
            else:        
                time.sleep(1)
                dlg_over = wx.MessageDialog(self, u'保存完成', '', wx.OK)
                dlg_over.ShowModal()

        f.close()
        
        pass
    def Delrecord(self, event):
        pass
    def Resetrecord(self, event):
        pass
    def Selected(self, event):
        for v in self.list_all.GetSelections():
            self.list_query.Append(self.list_all.GetString(v))
        
    def Unselected(self, event):
        for v in self.list_query.GetSelections():
            self.list_query.Delete(v)
        
    def Query(self, event):
        v = 0        
        
        self.gauge = wx.Gauge(self, -1, len(self.list_query.GetSelections()), (0, 315), (785, 20))
        try: 
            while True:
#                print self.list_query.GetString(v)
                xl.WSCustomQuery(self.list_query.GetString(v))                
                self.gauge.SetValue(v + 1)
                v += 1
#                print v
        except:
            self.gauge.SetValue(len(self.list_query.GetSelections()))          
            
            if v == 0:
                dlg_warning = wx.MessageDialog(self, u'请选择要查询的记录', '选择错误！', wx.OK)
                dlg_warning.ShowModal()
            else:        
                time.sleep(1)
                dlg_over = wx.MessageDialog(self, u'查询成功完成', '', wx.OK)
                dlg_over.ShowModal()
                self.gauge.SetValue(0)            
                self.gauge.Destroy()
##        
class WSCustomQuery(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        # create some sizers
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        grid = wx.GridBagSizer(hgap=6, vgap=6)
        hSizer = wx.BoxSizer(wx.HORIZONTAL)

        self.name_label = wx.StaticText(self, label=u"名称", size=(-1, -1))
        self.year_label = wx.StaticText(self, label=u"年份", size=(-1, -1))
        self.table_label = wx.StaticText(self, label=u"表名")
#        self.cell_label = wx.StaticText(self, label=u"单元名")
        self.table_expr = wx.StaticText(self, label=u"公式")
        self.table_all = wx.StaticText(self, label=u"所有记录")
        self.table_selected = wx.StaticText(self, label=u"查询记录")
        
        self.name_box = wx.TextCtrl(self, size=(100, -1))
        self.year_box = wx.Choice(self, size=(60, -1), choices=xl.years)
        self.table_box = wx.TextCtrl(self, size=(60, -1))
#        self.cell_box = wx.TextCtrl(self, size=(60, -1))
        self.expr_box = wx.Choice(self, size=(35, -1), choices=['+', '-', ''])
        
        self.button_add = wx.Button(self, label=u'添加', size=(50, -1))
        self.button_del = wx.Button(self, label=u'删除', size=(50, -1))
        self.button_reset = wx.Button(self, label=u'重置', size=(50, -1))
        self.button_save = wx.Button(self, label=u'保存', size=(50, -1))
        
        self.list_all = wx.ListBox(self, -1, size=(470, 130), choices='', style=wx.LB_SINGLE | wx.LB_HSCROLL)
        self.list_query = wx.ListBox(self, -1, size=(470, 100), choices='', style=wx.LB_MULTIPLE | wx.LB_HSCROLL)
        self.button_query = wx.Button(self, -1, size=(100, 50), label=u'查询', style=wx.ALIGN_CENTER_VERTICAL)
        self.button_selected = wx.Button(self, -1, size=(-1, -1), label=u'↓↓', style=wx.ALIGN_CENTER_VERTICAL)
        self.button_unselected = wx.Button(self, -1, size=(-1, -1), label=u'↑↑', style=wx.ALIGN_CENTER_VERTICAL)
        
        self.percent_name = wx.StaticText(self, label=u'对比分析阀值', pos=(610, 50), size=(-1, 30))
        self.slade = wx.Slider(self, -1, 0, -100, 100, pos=(600, 80), size=(100, -1), style=wx.SL_HORIZONTAL | wx.SL_AUTOTICKS | wx.SL_LABELS)
        
        self.Bind(wx.EVT_BUTTON, self.Addrecord, self.button_add)
        self.Bind(wx.EVT_BUTTON, self.Delrecord, self.button_del)
        self.Bind(wx.EVT_BUTTON, self.Resetrecord, self.button_reset)
        self.Bind(wx.EVT_BUTTON, self.Saverecord, self.button_save)
        self.Bind(wx.EVT_BUTTON, self.Selected, self.button_selected)
        self.Bind(wx.EVT_BUTTON, self.Unselected, self.button_unselected)
        self.Bind(wx.EVT_BUTTON, self.Query, self.button_query)
        
        grid.Add(self.name_label, pos=(0, 1))
        grid.Add(self.year_label, pos=(0, 3))
        grid.Add(self.table_label, pos=(0, 5))
#        grid.Add(self.cell_label, pos=(0, 7))
        grid.Add(self.table_expr, pos=(0, 7))
        
        grid.Add(self.name_box, pos=(0, 2))
        grid.Add(self.year_box, pos=(0, 4))
        grid.Add(self.table_box, pos=(0, 6))
#        grid.Add(self.cell_box, pos=(0, 8))
        grid.Add(self.expr_box, pos=(0, 8))
        
        grid.Add(self.button_add, pos=(0, 9))
        grid.Add(self.button_del, pos=(0, 10))
        grid.Add(self.button_reset, pos=(0, 11))
        grid.Add(self.button_save, pos=(0, 12))
        
        grid.Add(self.list_all, pos=(1, 2), span=(1, 10))
        grid.Add(self.list_query, pos=(3, 2), span=(1, 10))
        grid.Add(self.table_all, pos=(1, 1))
        grid.Add(self.table_selected, pos=(3, 1))
        
        grid.Add(self.button_selected, pos=(2, 3), span=(1, 2), flag=wx.ALIGN_CENTER)
        grid.Add(self.button_unselected, pos=(2, 6), span=(1, 2), flag=wx.ALIGN_CENTER)
        
        grid.Add(self.button_query, pos=(3, 12), span=(1, 2), flag=wx.ALIGN_CENTER)
        
        hSizer.Add(grid, 0, wx.ALL, 5)
        mainSizer.Add(hSizer, 0, wx.ALL, 5)
#        mainSizer.Add(self.list_query, 0, wx.LEFT)
        self.SetSizerAndFit(mainSizer)
        self.Initrecord()
        
    def Initrecord(self):
        f = open(xl.ws_custom_conf, 'r')
        while True:
            line = f.readline()
#            print line.decode('gbk')
            if len(line) == 0:
                break
            else:
                self.list_all.Append(line.decode('utf-8'))
        f.close()
        
    def Addrecord(self, event):
        if self.name_box.GetValue() == '':
            dlg_over = wx.MessageDialog(self, u'名称不能为空', '')
            result = dlg_over.ShowModal()
        elif self.table_box.GetValue() == '':
            dlg_over = wx.MessageDialog(self, u'表名不能为空', '')
            result = dlg_over.ShowModal()            
        else:   
            record = self.table_box.GetValue() + self.expr_box.GetStringSelection()
            if self.expr_box.GetStringSelection() == '':
                dlg_over = wx.MessageDialog(self, u'确定是最后一个表名吗？', '')
                result = dlg_over.ShowModal()
                if result == wx.ID_OK:
    #                print record
                    if self.list_all.GetStringSelection() != '':
                        old = self.list_all.GetStringSelection()
                        self.list_all.Delete(self.list_all.GetSelection())
                        new = old + record
                        self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                        
                    else:
                        new = self.name_box.GetValue() + ':' + record
                        self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                else:
                    pass       
            else:        
#                print record
                if self.list_all.GetStringSelection() != '':
                    old = self.list_all.GetStringSelection()
                    self.list_all.Delete(self.list_all.GetSelection())
                    new = old + record
                    self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                else:
                    new = self.name_box.GetValue() + ':' + record
                    self.list_all.Insert(new, self.list_all.GetSelection() + 1)

    def Saverecord(self, event):
        v = 0
        f = open(xl.ws_custom_conf, 'w')
        try: 
            while True:
#                print 
                f.write(self.list_all.GetString(v))
                v += 1
        except:
            if v == 0:
                dlg_warning = wx.MessageDialog(self, u'无可保存的记录', '错误！', wx.OK)
                dlg_warning.ShowModal()
            else:        
                time.sleep(1)
                dlg_over = wx.MessageDialog(self, u'保存完成', '', wx.OK)
                dlg_over.ShowModal()

        f.close()
        
        pass
    def Delrecord(self, event):
        pass
    def Resetrecord(self, event):
        pass
    def Selected(self, event):
        for v in self.list_all.GetSelections():
            self.list_query.Append(self.list_all.GetString(v))
        
    def Unselected(self, event):
        for v in self.list_query.GetSelections():
            self.list_query.Delete(v)
        
    def Query(self, event):
        v = 0        
#        xl.WSCustomQuery(v, self.list_query.GetString(v), '2012', self.slade.GetValue())
        self.gauge = wx.Gauge(self, -1, len(self.list_query.GetSelections()), (0, 315), (785, 20))
        try: 
            while True:
#                print self.list_query.GetString(v)
                xl.WSCustomQuery(v, self.list_query.GetString(v), '2012', self.slade.GetValue())                
                self.gauge.SetValue(v + 1)
                v += 1
#                print v
        except:
            self.gauge.SetValue(len(self.list_query.GetSelections()))          
            
            if v == 0:
                dlg_warning = wx.MessageDialog(self, u'请选择要查询的记录', '选择错误！', wx.OK)
                dlg_warning.ShowModal()
            else:        
                time.sleep(1)
                dlg_over = wx.MessageDialog(self, u'查询成功完成', '', wx.OK)
                dlg_over.ShowModal()
                self.gauge.SetValue(0)            
                self.gauge.Destroy()


class CellCustomQuery(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        # create some sizers
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        grid = wx.GridBagSizer(hgap=6, vgap=6)
        hSizer = wx.BoxSizer(wx.HORIZONTAL)

        self.name_label = wx.StaticText(self, label=u"名称", size=(-1, -1))
        self.year_label = wx.StaticText(self, label=u"年份", size=(-1, -1))
        self.table_label = wx.StaticText(self, label=u"表名")
        self.cell_label = wx.StaticText(self, label=u"单元名")
        self.table_expr = wx.StaticText(self, label=u"公式")
        self.table_all = wx.StaticText(self, label=u"所有记录")
        self.table_selected = wx.StaticText(self, label=u"查询记录")
        
        self.name_box = wx.TextCtrl(self, size=(100, -1))
        self.year_box = wx.Choice(self, size=(60, -1), choices=xl.years)
        self.table_box = wx.TextCtrl(self, size=(60, -1))
        self.cell_box = wx.TextCtrl(self, size=(60, -1))
        self.expr_box = wx.Choice(self, size=(35, -1), choices=['+', '-', ''])
        
        self.button_add = wx.Button(self, label=u'添加', size=(50, -1))
        self.button_del = wx.Button(self, label=u'删除', size=(50, -1))
        self.button_reset = wx.Button(self, label=u'重置', size=(50, -1))
        self.button_save = wx.Button(self, label=u'保存', size=(50, -1))
        
        self.list_all = wx.ListBox(self, -1, size=(470, 130), choices='', style=wx.LB_SINGLE | wx.LB_HSCROLL)
        self.list_query = wx.ListBox(self, -1, size=(470, 100), choices='', style=wx.LB_MULTIPLE | wx.LB_HSCROLL)
        self.button_query = wx.Button(self, -1, size=(100, 50), label=u'查询', style=wx.ALIGN_CENTER_VERTICAL)
        self.button_selected = wx.Button(self, -1, size=(-1, -1), label=u'↓↓', style=wx.ALIGN_CENTER_VERTICAL)
        self.button_unselected = wx.Button(self, -1, size=(-1, -1), label=u'↑↑', style=wx.ALIGN_CENTER_VERTICAL)
        
        self.Bind(wx.EVT_BUTTON, self.Addrecord, self.button_add)
        self.Bind(wx.EVT_BUTTON, self.Delrecord, self.button_del)
        self.Bind(wx.EVT_BUTTON, self.Resetrecord, self.button_reset)
        self.Bind(wx.EVT_BUTTON, self.Saverecord, self.button_save)
        self.Bind(wx.EVT_BUTTON, self.Selected, self.button_selected)
        self.Bind(wx.EVT_BUTTON, self.Unselected, self.button_unselected)
        self.Bind(wx.EVT_BUTTON, self.Query, self.button_query)
        
        grid.Add(self.name_label, pos=(0, 1))
        grid.Add(self.year_label, pos=(0, 3))
        grid.Add(self.table_label, pos=(0, 5))
        grid.Add(self.cell_label, pos=(0, 7))
        grid.Add(self.table_expr, pos=(0, 9))
        
        grid.Add(self.name_box, pos=(0, 2))
        grid.Add(self.year_box, pos=(0, 4))
        grid.Add(self.table_box, pos=(0, 6))
        grid.Add(self.cell_box, pos=(0, 8))
        grid.Add(self.expr_box, pos=(0, 10))
        
        grid.Add(self.button_add, pos=(0, 11))
        grid.Add(self.button_del, pos=(0, 12))
        grid.Add(self.button_reset, pos=(0, 13))
        grid.Add(self.button_save, pos=(0, 14))
        
        grid.Add(self.list_all, pos=(1, 2), span=(1, 10))
        grid.Add(self.list_query, pos=(3, 2), span=(1, 10))
        grid.Add(self.table_all, pos=(1, 1))
        grid.Add(self.table_selected, pos=(3, 1))
        
        grid.Add(self.button_selected, pos=(2, 3), span=(1, 2), flag=wx.ALIGN_CENTER)
        grid.Add(self.button_unselected, pos=(2, 6), span=(1, 2), flag=wx.ALIGN_CENTER)
        
        grid.Add(self.button_query, pos=(3, 12), span=(1, 2), flag=wx.ALIGN_CENTER)
        
        hSizer.Add(grid, 0, wx.ALL, 5)
        mainSizer.Add(hSizer, 0, wx.ALL, 5)
#        mainSizer.Add(self.list_query, 0, wx.LEFT)
        self.SetSizerAndFit(mainSizer)
        self.Initrecord()
        
    def Initrecord(self):
        f = open(xl.cell_custom_conf, 'r')
        while True:
            line = f.readline()
#            print line.decode('gbk')
            if len(line) == 0:
                break
            else:
                self.list_all.Append(line.decode('utf-8'))
        f.close()
        
    def Addrecord(self, event):
        if self.name_box.GetValue() == '':
            dlg_over = wx.MessageDialog(self, u'名称不能为空', '')
            result = dlg_over.ShowModal()
        elif self.table_box.GetValue() == '':
            dlg_over = wx.MessageDialog(self, u'表名不能为空', '')
            result = dlg_over.ShowModal()
        elif self.name_box.GetValue() == '':
            dlg_over = wx.MessageDialog(self, u'单元名不能为空', '')
            result = dlg_over.ShowModal()            
        else:   
            record = self.year_box.GetStringSelection() + '.' + self.table_box.GetValue() + '.' + self.cell_box.GetValue() + self.expr_box.GetStringSelection()
            if self.expr_box.GetStringSelection() == '':
                dlg_over = wx.MessageDialog(self, u'确定是最后一个单元名吗？', '')
                result = dlg_over.ShowModal()
                if result == wx.ID_OK:
    #                print record
                    if self.list_all.GetStringSelection() != '':
                        old = self.list_all.GetStringSelection()
                        self.list_all.Delete(self.list_all.GetSelection())
                        new = old + record
                        self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                        
                    else:
                        new = self.name_box.GetValue() + ':' + record
                        self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                else:
                    pass       
            else:        
#                print record
                if self.list_all.GetStringSelection() != '':
                    old = self.list_all.GetStringSelection()
                    self.list_all.Delete(self.list_all.GetSelection())
                    new = old + record
                    self.list_all.Insert(new, self.list_all.GetSelection() + 1)
                else:
                    new = self.name_box.GetValue() + ':' + record
                    self.list_all.Insert(new, self.list_all.GetSelection() + 1)

    def Saverecord(self, event):
        v = 0
        f = open(xl.cell_custom_conf, 'w')
        try: 
            while True:
#                print 
                f.write(self.list_all.GetString(v))
                v += 1
        except:
            if v == 0:
                dlg_warning = wx.MessageDialog(self, u'无可保存的记录', '错误！', wx.OK)
                dlg_warning.ShowModal()
            else:        
                time.sleep(1)
                dlg_over = wx.MessageDialog(self, u'保存完成', '', wx.OK)
                dlg_over.ShowModal()

        f.close()
        
        pass
    def Delrecord(self, event):
        pass
    def Resetrecord(self, event):
        pass
    def Selected(self, event):
        for v in self.list_all.GetSelections():
            self.list_query.Append(self.list_all.GetString(v))
        
    def Unselected(self, event):
        for v in self.list_query.GetSelections():
            self.list_query.Delete(v)
        
    def Query(self, event):
        v = 0        
        
        self.gauge = wx.Gauge(self, -1, len(self.list_query.GetSelections()), (0, 315), (785, 20))
        try: 
            while True:
#                print self.list_query.GetString(v)
                xl.CellCustomQuery(self.list_query.GetString(v))                
                self.gauge.SetValue(v + 1)
                v += 1
#                print v
        except:
            self.gauge.SetValue(len(self.list_query.GetSelections()))          
            
            if v == 0:
                dlg_warning = wx.MessageDialog(self, u'请选择要查询的记录', '选择错误！', wx.OK)
                dlg_warning.ShowModal()
            else:        
                time.sleep(1)
                dlg_over = wx.MessageDialog(self, u'查询成功完成', '', wx.OK)
                dlg_over.ShowModal()
                self.gauge.SetValue(0)            
                self.gauge.Destroy()
##        
    
class PercentQuery(MainPanel):
    def __init__(self, parent):
    
        MainPanel.__init__(self, parent)
        
        self.percent_name = wx.StaticText(self, label=u'对比分析比例(%)', pos=(550, 20), size=(-1, 30))
#        self.percent_value = wx.TextCtrl(self, size=(65, -1), value='10', pos=(560, 50), style=wx.ALIGN_CENTER_HORIZONTAL)
        self.slade = wx.Slider(self, -1, 10, 1, 100, pos=(550, 50), size=(100, -1), style=wx.SL_HORIZONTAL | wx.SL_AUTOTICKS | wx.SL_LABELS)   
              
        self.Bind(wx.EVT_BUTTON, self.PerQuery, self.query_button)
        self.GetRecord()
            
    def GetRecord(self):
        self.list_all.Clear()
        f = file(xl.cell_cells_conf, 'r')
        while True:        
            line = f.readline()                
            if len(line) == 0:
                break            
            if self.list_all.FindString(line.split('.')[0].decode('gbk')) == -1:
                self.list_all.Append(line.split('.')[0].decode('gbk'))
               
    def PerQuery(self, event):
        try:
            percent_value = int(self.slade.GetValue())
            print percent_value
            self.gauge = wx.Gauge(self, -1, len(self.list_all.GetSelections()), (0, 315), (785, 20))
            v = 0
            for v in range(0, 100000):
                print 'v=', v
                
                xl.PercentQuery(v, percent_value, self.list_selected.GetString(v))
                self.gauge.SetValue(v + 1)
            self.gauge.SetValue(len(self.list_all.GetSelections() * 10))        
            time.sleep(1)
            dlg_over = wx.MessageDialog(self, u'查询成功完成', '', wx.OK)
            dlg_over.ShowModal()
            self.gauge.SetValue(0)
        except:
            dlg_warning = wx.MessageDialog(self, u'查询成功完成！', '', wx.OK)
            dlg_warning.ShowModal()
            self.gauge.SetValue(0)
#                    

class Frame(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(800, 430), pos=((1366 - 800) / 2, 100), style=wx.CAPTION | wx.CLOSE_BOX | wx.SYSTEM_MENU | wx.MINIMIZE_BOX)
        
#        wx.Frame.__init__(self, parent, title=title, size=(830, 450), pos=((1366 - 800) / 2, 100))
        self.dirname = ''
        filemenu = wx.Menu()
        helpmenu = wx.Menu()
        
        menuOpen = filemenu.Append(wx.ID_OPEN, u"&导入配置", " Open a file to import me!")
#        filemenu.AppendSeparator()
        menuExit = filemenu.Append(wx.ID_EXIT, u"&退出", " Terminate the program")
        
        menuHelp = helpmenu.Append(wx.ID_HELP, u"&帮助索引", " I Will add it later.")
        menuAbout = helpmenu.Append(wx.ID_ABOUT, u"&关于我", " Information about this program")
#        filemenu.AppendSeparator()
        
        menuba = wx.MenuBar()
        menuba.Append(filemenu, u"&文件")  # Adding the "filemenu" to the MenuBar
        menuba.Append(helpmenu, u"&帮助")  # Adding the "filemenu" to the MenuBar
        self.SetMenuBar(menuba)
        
        self.Bind(wx.EVT_MENU, self.OnOpen, menuOpen)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnHelp, menuHelp)
        
        
        self.CreateStatusBar()
#        self.Bind(wx.EVT_CLOSE, self.onExit, menuExit)
        self.__set_properties()
        
    def __set_properties(self):
        # begin wxGlade: mainWindow.__set_properties
        _icon = wx.EmptyIcon()
        _icon.CopyFromBitmap(wx.Bitmap("./favicon.ico", wx.BITMAP_TYPE_ANY))
        self.SetIcon(_icon)
        
    def OnAbout(self, e):
        # Create a message dialog box
        dlg = wx.MessageDialog(self, " A small program to analyze data in MS Excel", "About me", wx.OK)
        dlg.ShowModal()  # Shows it
        dlg.Destroy()  # finally destroy it when finished.

    def OnExit(self, e):
        self.Close(True)  # Close the frame.

    def OnOpen(self, e):
        """ Open a file"""
        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.*", wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            f = open(os.path.join(self.dirname, self.filename), 'r')
            self.control.SetValue(f.read())
            f.close()
        dlg.Destroy()
    def OnHelp(self, e):
        pass

if __name__ == '__main__':    
    app = wx.App(redirect=False)
    frame = Frame(None, u'ECL数据对比分析  2.0')
    
    nb = wx.Notebook(frame)
    
    namelist = [u'标准查询', u'比例查询', u'自定义单元格查询', u'自定义工作簿查询']
        
    #for name in namelist:
    nb.AddPage(StandQuery(nb), namelist[0])
    nb.AddPage(PercentQuery(nb), namelist[1])
    nb.AddPage(CellCustomQuery(nb), namelist[2])
    nb.AddPage(WSCustomQuery(nb), namelist[3])
    
    cert = platform.machine() + platform.node() + platform.platform() + platform.processor() + platform.release()
    cert_file = file('./conf/cert', 'w')
    cert_file.write(base64.encodestring(cert))
    cert_file.close()
    cert_md5 = hashlib.md5(base64.encodestring(cert))
#    print cert_md5.hexdigest()
    t = time.strftime('%Y,%m,%d', time.localtime(time.time()))
    y = int(t.split(',')[0])
    m = int(t.split(',')[1])
    d = int(t.split(',')[2])
    d1 = datetime.datetime(2013, 4, 1)
    d2 = datetime.datetime(y, m, d)
    licence = ''
    if (d1 - d2).days > 0:
        f1 = file('./conf/licence', 'r')
        licence_content = f1.readline()
        for i in range(len(licence_content) - 1, -1, -1):
            licence += licence_content[i]
            
        if licence == cert_md5.hexdigest():
            print 'Licence available'        
            frame.Show(True)
        else:
            print 'Licence error'
            frame.Close()
        f1.close   
        
        frame.Show(True)
    else:
        frame.Close()
    
        
app.MainLoop()
