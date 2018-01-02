#C:Python31python.exe
# -*- coding:utf-8 -*-

import sys
import xlrd

# Settings excel column index
SETTINGS = 2; SETTINGS_VALUE = 3

# Event variable excel column index
INDEX = 0; SEMI = 1; EVENT = 2; ALIAS = 3; REPORT = 4;
EVENT_ID = 5; DV = 6; DV_VALUE = 7; DATA_TYPE = 8; CLASS = 9

class Event:
    def __init__(self, index, SEMI, event, alias, report, eventid, dv, dataType, classes):
        self.index = index
        self.SEMI = SEMI
        self.event = event
        self.alias = alias
        self.report = report
        self.eventid = eventid
        self.dv = dv
        self.dataType = dataType
        self.classes = classes

INPUT_EXCEL = "test.xlsx"   # input excel file
OUTPUT_FILE = "xxx.cfg"     # output config file

def main(argv):
    ad_wb = xlrd.open_workbook(INPUT_EXCEL) # get excel file
    sheet_0 = ad_wb.sheet_by_index(0)       # get sheet

    event_list = []     # class Event array

    # make xlsx to Event array
    for row in range(sheet_0.nrows):
        if sheet_0.cell_value(row, INDEX) == 'Index':
            if sheet_0.cell_value(row, DV) == 'Valid DV':
                # put all name & value in dictionary as {VID_NAME: VID_VALUE}
                dv_dict = {}
                valid_dv_index = row + 1 # skip title
                while True:
                    if valid_dv_index < sheet_0.nrows and sheet_0.cell_value(valid_dv_index, DV) != '':
                        dv_dict.update({
                            sheet_0.cell_value(valid_dv_index, DV):
                            sheet_0.cell_value(valid_dv_index, DV_VALUE) if sheet_0.cell_value(valid_dv_index, DV_VALUE) != '' else 0})
                        valid_dv_index += 1
                    else:
                        break

            # convert excel Event to class Event
            data_row = row + 1
            event = Event(sheet_0.cell_value(data_row, INDEX), sheet_0.cell_value(data_row, SEMI), # index, SEMI
                          sheet_0.cell_value(data_row, EVENT), sheet_0.cell_value(data_row, ALIAS), # event, alias
                          sheet_0.cell_value(data_row, REPORT), sheet_0.cell_value(data_row, EVENT_ID), # report, eventid
                 dv_dict, sheet_0.cell_value(data_row, DATA_TYPE), sheet_0.cell_value(data_row, CLASS)) # dv(vid, value), dataType, class
            event_list.append(event)

    # start GemDCConfig
    f = open(OUTPUT_FILE, 'w', encoding = 'UTF-8')
    f.write('[GemDCConfig]\n')
    settings = []
    i = 1
    f.write("Settings=[")
    while True:
        f.write("{}={}".format(sheet_0.cell_value(i, EVENT), sheet_0.cell_value(i, ALIAS)))
        if sheet_0.cell_value(i+1, SETTINGS) == '':
            break;
        else:
            f.write(', '); i += 1
    f.write(']\n')
    # end GemDCConfig

    # start Vids
    f.write('\n[Vids]\n')
    vid_set = set()
    vid_list = list()
    for ev in event_list:
        for x in ev.dv: # ev.dv = Valid DV list
            if x not in vid_set:
                f.write("Vid=[ID={}, Name={}, Type={}, Class={}]\n".format(int(ev.dv[x]), x, ev.dataType, ev.classes))
                vid_set.add(x)
                vid_list.insert(int(ev.dv[x]), x)
    # end Vids

    # start Events
    f.write('\n[Events]\n')
    for ev in event_list:
        if ev.eventid != '':
            f.write("Event=[ID={}, Name={}, Enable=True]\n".format(int(ev.eventid), ev.alias))
        else:
            print('[{}] EventID can\'t not be null'.format(int(ev.index)))
    # end Events

    # start Report
    f.write('\n[Reports]\n')
    i = 1
    event_link_report_list = list()
    for ev in event_list:
        event_vid_list = list()
        for x in ev.dv: # ev.dv = Valid DV list
            event_vid_list.append(vid_list.index(x))
        event_link_report_list.append(ev.alias)
        ### argument ID is temp
        f.write("Report=[ID={}, Name={}{}, Vids={}]\n".format(i, ev.SEMI+'_' if ev.SEMI != '' else '', ev.alias, event_vid_list))
        i += 1
    # end Report

    # start ReportLinks
    f.write('\n[ReportLinks]\n')
    for ev in event_link_report_list:
        f.write("ReportLink=[Event={}, Reports=[{}]]\n".format(ev, event_link_report_list.index(ev)+1))
    # end ReportLinks

    f.close

if __name__ == "__main__":
    main(sys.argv)