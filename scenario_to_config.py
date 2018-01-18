#C:Python31python.exe
# -*- coding:utf-8 -*-

import sys
import os
import xlrd

# Settings excel column index
SETTINGS = 2; SETTINGS_VALUE = 3

# Event variable excel column index
INDEX = 0; SEMI = 1; EVENT = 2; ALIAS = 3; REPORT = 4;
EVENT_ID = 5; VALID_DV = 6; DV_ID = 7; DATA_TYPE = 8; CLASS = 9

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

def main(argv):
    if len(sys.argv) < 2:
        print('No input file')
        os.system("pause")

    INPUT_EXCEL = sys.argv[1]   # input excel file

    try:
        ad_wb = xlrd.open_workbook(INPUT_EXCEL) # get excel file
        sheet_0 = ad_wb.sheet_by_index(0)       # get sheet
    except Exception as ex:
        print(ex)
        os.system("pause")

    event_list = []     # class Event array

    # make xlsx to Event array
    for row in range(sheet_0.nrows):
        # scan all row to find which cell is 'Index'
        if sheet_0.cell_value(row, INDEX) == 'Index':
            # find cell 'Valid DV' by correspond position
            if sheet_0.cell_value(row, VALID_DV) == 'Valid DV':
                # put all name & value in dictionary as {VID_NAME: VID_VALUE}
                dv_dict = {}
                row_of_dv = row + 1 # [Valid DV & DV Value]'s row
                while True:
                    if row_of_dv < sheet_0.nrows and sheet_0.cell_value(row_of_dv, VALID_DV) != '':
                        dv_dict.update({
                            sheet_0.cell_value(row_of_dv, VALID_DV):
                            sheet_0.cell_value(row_of_dv, DV_ID) if sheet_0.cell_value(row_of_dv, DV_ID) != '' else 0})
                        row_of_dv += 1 # Load next row if not blank line
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
    vendor = sheet_0.cell_value(1, 0) if sheet_0.cell_value(1, 0) != '' else "Vendor"
    module = sheet_0.cell_value(1, 1) if sheet_0.cell_value(1, 1) != '' else "Module"
    OUTPUT_FILE = vendor + "_" + module + ".cfg"     # output config file
    print("out:"+OUTPUT_FILE)

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
    vid_dict = {} # {Valid DV: DV ID}
    for ev in event_list:
        for obj_name in ev.dv: # ev.dv = Valid DV list
            if obj_name not in vid_set:
                f.write("Vid=[ID={}, Name={}, Type={}, Class={}]\n".format(int(ev.dv[obj_name]), obj_name, ev.dataType, ev.classes))
                vid_set.add(obj_name) # prevent duplicate Vid
                vid_dict.update( {obj_name : int(ev.dv[obj_name])} )
    # end Vids

    # start Events
    f.write('\n[Events]\n')
    for ev in event_list:
        # write Event even EventID doesn't exist
        f.write("Event=[ID={}, Name={}, Enable=True]\n".format(ev.eventid if ev.eventid == '' else int(ev.eventid), ev.alias))
    # end Events

    # start Report
    f.write('\n[Reports]\n')
    i = 1
    event_link_report_list = list()
    for ev in event_list:
        event_vid_list = list()
        for x in ev.dv: # ev.dv = Valid DV list
            event_vid_list.append(vid_dict[x])
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