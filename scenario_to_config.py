# C:Python31python.exe
# -*- coding:utf-8 -*-

import sys
import xlrd

# Settings excel column index
SETTINGS = 2; SETTINGS_VALUE = 3

# Event variable excel column index
class Column:
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

# method
def isDuplicate(list, element):
    for obj in list:
        print(obj)

def main(argv):
    if len(sys.argv) < 2:
        print('No input file')
        input('Press Enter to exit')

    INPUT_EXCEL = sys.argv[1]   # input excel file

    try:
        ad_wb = xlrd.open_workbook(INPUT_EXCEL) # get excel file
        sheet_0 = ad_wb.sheet_by_index(0)       # get sheet
    except Exception as ex:
        print(ex)
        input('Press Enter to exit')


    ##########################################
    # Convert xlsx Event to list(event_list[])
    ##########################################
    event_list = []     # Class Event list

    duplicate_check_set = set()

    for row in range(sheet_0.nrows):
        # scan all row to find which cell is 'Index'
        if (sheet_0.cell_value(row, Column.INDEX) == 'Index' and
            sheet_0.cell_value(row + 1, Column.EVENT_ID) != ''):

            # find cell 'Valid DV' by corresponding position
            if sheet_0.cell_value(row, Column.VALID_DV) == 'Valid DV':
                dv_id_dict = {}     # {Valid DV : DV ID}
                data_type_dict = {} # {Valid DV : Data Type}
                class_dict = {}     # {Valid DV : Class}

                value_row = row + 1

                # put DV ID, Data Type, Class to corresponding dictionary
                while value_row < sheet_0.nrows and sheet_0.cell_value(value_row, Column.VALID_DV) != '':
                    if (sheet_0.cell_value(value_row, Column.DV_ID) != '' and
                            sheet_0.cell_value(value_row, Column.DATA_TYPE) != '' and
                            sheet_0.cell_value(value_row, Column.CLASS) != ''):
                        dv_id_dict.update({sheet_0.cell_value(value_row, Column.VALID_DV) : sheet_0.cell_value(value_row, Column.DV_ID)})
                        data_type_dict.update({sheet_0.cell_value(value_row, Column.VALID_DV) : sheet_0.cell_value(value_row, Column.DATA_TYPE)})
                        class_dict.update({sheet_0.cell_value(value_row, Column.VALID_DV) : sheet_0.cell_value(value_row, Column.CLASS)})
                    value_row += 1 # Load next row if not blank line

            # convert excel Event to class Event
            data_row = row + 1
            event = Event(sheet_0.cell_value(data_row, Column.INDEX), sheet_0.cell_value(data_row, Column.SEMI), # Index, SEMI
                          sheet_0.cell_value(data_row, Column.EVENT), sheet_0.cell_value(data_row, Column.ALIAS), # Event, Alias
                          sheet_0.cell_value(data_row, Column.REPORT), int( sheet_0.cell_value(data_row, Column.EVENT_ID) ), # Report, EventID
                          dv_id_dict, data_type_dict, class_dict) # {Valid ID, DV ID}, {Valid ID, Data Type}, {Valid ID, Class}
            event_list.append(event)
            # check if Event ID is duplicated
            if int(sheet_0.cell_value(data_row, Column.EVENT_ID)) not in duplicate_check_set:
                duplicate_check_set.add(int( sheet_0.cell_value(data_row, Column.EVENT_ID) ))
            else:
                print("WARNING - Duplicated Event:{} in line {}."
                    .format(sheet_0.cell_value(data_row, Column.EVENT), data_row))

    # start GemDCConfig
    vendor = sheet_0.cell_value(1, 0) if sheet_0.cell_value(1, 0) != '' else "Vendor"
    module = sheet_0.cell_value(1, 1) if sheet_0.cell_value(1, 1) != '' else "Module"

    # cell value type, ctyp = 2 stand for float, transform float to string
    if sheet_0.cell(1, 0).ctype == 2:
        vendor = str(int(sheet_0.cell_value(1, 0)))
    if sheet_0.cell(1, 1).ctype == 2:
        module = str(int(sheet_0.cell_value(1, 1)))

    OUTPUT_FILE = vendor + "_" + module + ".cfg"     # output config file

    f = open(OUTPUT_FILE, 'w', encoding = 'UTF-8')
    f.write('[GemDCConfig]\n')
    settings = []
    i = 1
    f.write("Settings=[")
    while True:
        f.write("{}={}".format(sheet_0.cell_value(i, Column.EVENT), sheet_0.cell_value(i, Column.ALIAS)))
        if sheet_0.cell_value(i+1, SETTINGS) == '':
            break;
        else:
            f.write(', '); i += 1
    f.write(']\n')
    # end GemDCConfig

    # start EqConst
    f.write('\n[EqConst]\n');
    # end EqConst

    # start Vids
    f.write('\n[Vids]\n')
    Valid_DV_set = set()
    Valid_dv_id_dict = {} # {Valid DV: DV ID}
    for ev in event_list:
        for obj_name in ev.dv: # ev.dv = Valid DV list
            if obj_name not in Valid_DV_set:
                Valid_DV_set.add(obj_name) # prevent duplicate Vid
                Valid_dv_id_dict.update( {obj_name : int(ev.dv[obj_name])} )
                f.write("Vid=[ID={}, Name={}, Type={}, Class={}]\n".format(Valid_dv_id_dict[obj_name], obj_name, ev.dataType[obj_name], ev.classes[obj_name]))
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
        for x in ev.dv:
            if Valid_dv_id_dict[x] in event_vid_list:
                print("WARNING - Duplicated Vid:{} in Event:{}.".format(x, ev.alias))
            event_vid_list.append(Valid_dv_id_dict[x])
        event_link_report_list.append(ev.eventid)
        f.write("Report=[ID={}, Name={}{}, Vids={}]\n".format(i, ev.SEMI+'_' if ev.SEMI != '' else '', ev.alias, event_vid_list))
        i += 1
    # end Report

    # start ReportLinks
    f.write('\n[ReportLinks]\n')
    for ev in event_link_report_list:
        f.write("ReportLink=[Event={}, Reports=[{}]]\n".format(ev, event_link_report_list.index(ev)+1))
    # end ReportLinks

    # start Alarm
    f.write('\n[Alarms]\n');
    # end Alarm

    f.close

    print("Config file: " + OUTPUT_FILE)


if __name__ == "__main__":
    main(sys.argv)
    input('Press Enter to exit')