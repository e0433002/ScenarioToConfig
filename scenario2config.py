#C:Python31python.exe
# -*- coding:utf-8 -*-

import sys
import xlrd

# Settings index
SETTINGS = 2; SETTINGS_VALUE = 3

# Event index
INDEX = 0; SEMI = 1; EVENT = 2; ALIAS = 3; REPORT = 4;
EVENT_ID = 5; DV = 6; DATA_TYPE = 8; CLASS = 9

class Event:
    def __init__(self, index, SEMI, event, alias, report,
                 eventid, dv, dataType, classes):
        self.index = index
        self.SEMI = SEMI
        self.event = event
        self.alias = alias
        self.report = report
        self.eventid = eventid
        self.dv = dv
        self.dataType = dataType
        self.classes = classes

    def deposit(self, amount):
        if amount <= 0:
            raise ValueError('must be positive')
        self.balance += amount

    def withdraw(self, amount):
        if amount <= self.balance:
            self.balance -= amount
        else:
            raise RuntimeError('balance not enough')

def main(argv):
    ad_wb = xlrd.open_workbook("test.xlsx") # get excel file
    sheet_0 = ad_wb.sheet_by_index(0)       # get sheet

    event_list = [] # event array

    for row in range(sheet_0.nrows):
        if sheet_0.cell_value(row, 0) == 'Index':
            if sheet_0.cell_value(row, 6) == 'Valid DV':
                # put all name & value in dictionary as {VID_NAME: VID_VALUE}
                dv_dict = {}
                valid_dv_index = row + 1 # skip title
                while True:
                    if valid_dv_index < sheet_0.nrows and sheet_0.cell_value(valid_dv_index, 6) != '':
                        dv_dict.update({sheet_0.cell_value(valid_dv_index, 6):
                            sheet_0.cell_value(valid_dv_index, 7) if sheet_0.cell_value(valid_dv_index, 7) != '' else 0})
                        valid_dv_index += 1
                    else:
                        break

            data_row = row + 1
            event = Event(sheet_0.cell_value(data_row, INDEX), sheet_0.cell_value(data_row, SEMI), # index, SEMI
                          sheet_0.cell_value(data_row, EVENT), sheet_0.cell_value(data_row, ALIAS), # event, alias
                          sheet_0.cell_value(data_row, REPORT), sheet_0.cell_value(data_row, EVENT_ID), # report, eventid
                 dv_dict, sheet_0.cell_value(data_row, DATA_TYPE), sheet_0.cell_value(data_row, CLASS)) # dv(vid, value), dataType, class
            event_list.append(event)

    # start writing GemDCConfig
    f = open('xxx.cfg', 'w', encoding = 'UTF-8')
    f.write('[GemDCConfig]\n')
    settings = []
    i = 1
    f.write("Settings=[")
    while True:
        f.write("{}={}".format(sheet_0.cell_value(i, 2), sheet_0.cell_value(i, 3)))
        if sheet_0.cell_value(i+1, SETTINGS) == '':
            break;
        else:
            f.write(', '); i += 1
    f.write(']\n\n')
    # end writing GemDCConfig

    # start writing Vids
    f.write('[Vids]\n')
    vid_set = set()
    for ev in event_list:
        for x in ev.dv: # ev.dv = Valid DV list
            # print('x={}, value={}'.format(x, int(ev.dv[x])))
            if x not in vid_set:
                f.write("Vid=[ID={}, Name={}, Type={}, Class={}]\n".format(int(ev.dv[x]), x, ev.dataType, ev.classes))
                vid_set.add(x)
    # end writing Vids

    f.close


if __name__ == "__main__":
    main(sys.argv)