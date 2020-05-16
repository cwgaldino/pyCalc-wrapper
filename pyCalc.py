#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""Support function for connecting with libreoffice Calc."""

# standard imports
import numpy as np
from pathlib import Path
import os
import inspect
import psutil
import signal
import subprocess
import time

# import uno
from unotools import Socket, connect
from unotools.component.calc import Calc
from unotools.unohelper import convert_path_to_url

# calcObject (xlsx)
def connect2Calc(file=None, port=8100, counter_max=5000):
    """Open libreoffice and enable conection with Calc.

    Args:
        file (str or pathlib.Path, optional): file to connect. If None, it will
            open a new Calc instance.
        port (int, optional): port for connection.
        counter_max (int, optional): Max number of tentatives to establish a
            connection.

    Returns:
        Calc object.

        The main mathods defined for a Calc object exemplyfied down below:

        >>> # adds one sheet ('Sheet2') at position 1
        >>> calcObject.insert_sheets_new_by_name('Sheet2', 1)
        >>>    
        >>> # adds multiple sheets ('Sheet3' and 'Sheet4) at position 2
        >>> calcObject.insert_multisheets_new_by_name(['Sheet3', 'Sheet4'], 2)
        >>>
        >>> # Get number of sheets
        >>> print(calcObject.get_sheets_count())
        4
        >>> # Remove sheets
        >>> calcObject.remove_sheets_by_name('Sheet3')
        >>> # get sheet data
        >>> sheet1 = calcObject.get_sheet_by_name('Sheet1')
        >>> sheet2 = calcObject.get_sheet_by_index(0)

        Also, use :py:func:`~backpack.figmanip.setFigurePosition`
    """
    # open libreoffice
    libreoffice = subprocess.Popen([f"soffice --accept='socket,host=localhost,port={port};urp;'"], shell=True, close_fds=True)

    # connect to libreoffice
    connected = False
    counter = 0
    while connected == False:
        time.sleep(0.5)
        try:
            context = connect(Socket('localhost', 8100))
            connected = True
        except:
            counter += 1
            if counter == counter_max:
                raise ConnectionError('Cannot establish connection, maybe try increasing counter_max value.')
            pass

    if file is None:
        return Calc(context)
    else:
        file = Path(file)
        return Calc(context, convert_path_to_url(str(file)))


def closeCalc(calcObject):
    """Close Calc.

    Args:
        calcObject (Calc object): Object created by connect2calc().
    """
    calcObject.close(True)
    return


def _findProcessIdByName(string):
    """Get a list of all the PIDs of all the running process whose name contains
    string.

    Args:
        string (str): string.
    """

    listOfProcessObjects = []

    # Iterate over the all the running process
    for proc in psutil.process_iter():
       try:
           pinfo = proc.as_dict(attrs=['pid', 'name', 'create_time'])
           # Check if process name contains the given name string.
           if string.lower() in pinfo['name'].lower() :
               listOfProcessObjects.append(pinfo)
       except (psutil.NoSuchProcess, psutil.AccessDenied , psutil.ZombieProcess):
           pass

    return listOfProcessObjects;


def _libreoffice_processes():
    """Return a list of processes associated with libreoffice.

    Note:
        This function try to match the name of a process with names tipically
        related with libreoffice processes ('soffice.bin' or 'oosplash').
        Therefore, it might return processes that are not related to
        libreoffice if their name mathces with words: 'soffice.bin'
        and 'oosplash'."""
    process_list = []
    for proc in _findProcessIdByName('soffice'):
        process_list.append(proc)
    for proc in _findProcessIdByName('oosplash'):
        process_list.append(proc)
    return process_list


def kill_libreoffice_processes():
    """Kill libreoffice processes.

    Note:
        It will close ALL processes that are related to libreoffice (processes
        that have 'soffice.bin' or 'oosplash' in their name)."""

    process_list = _libreoffice_processes()
    for proc in process_list:
        os.kill(proc['pid'], signal.SIGKILL)


def saveCalc(calcObject, filepath='./untitled.xlsx'):
    """Save xlsx file.

    Note:
        If `filepath` have no suffix, it adds '.xlsx' at the end of filepath.

    Args:
        calcObject (Calc object): Object created by :py:func:`calcmanip.connect2Calc`.
        filepath (string or pathlib.Path, optional): filepath to save file.
    """
    filepath = Path(filepath)

    # fix extension
    if filepath.suffix == '':
        filepath = filepath.parent / (str(filepath.name) + '.xlsx')

    # save
    url = convert_path_to_url(str(filepath))
    calcObject.store_as_url(url, 'FilterName')



# calcObject manipulation
def get_sheets_name(calcObject):
    """Get sheets names in a list.
    """

    return calcObject

def copyCells(calcObject, sheet2CopyFrom, sheet2PasteAt, range2copy=(0, 0, 10, 10), range2paste=None, calcObject2Paste=None):
    """Copy and paste cells.

    Args:
        calcObject (
    """
    if type(sheet2CopyFrom) == str:
        sheetObject = calcObject.get_sheet_by_name(sheet2CopyFrom)
    else:
        sheetObject = sheet2CopyFrom

    if type(sheet2PasteAt) == str:
        sheetObject2 = calcObject.get_sheet_by_name(sheet2PasteAt)
    else:
        sheetObject2 = sheet2PasteAt

    # copy/paste data
    id_list = get_id(sheetObject)
    sheet_data = sheetObject.get_cell_range_by_position(0, 0, 10, len(id_list)+1).getDataArray()

    for row, row_data in enumerate(sheet_data):
        sheetObject2.set_columns_formula(0, row, row_data)
        for column in range(4, 10):  # copy/paste conditional formating
            sheetObject2.get_cell_by_position(column, row+1).ConditionalFormat = sheetObject.get_cell_by_position(column, row+1).ConditionalFormat

    # copy column Size
    cols = sheetObject.getColumns()
    cols2 = sheetObject2.getColumns()
    for col_number, col in enumerate(cols[0:11]):
        cols2[col_number].setPropertyValue('Width', col.Width)

    # fix color formating
    group_color(sheetObject2, calcObject=None)




def copy_sheet(calcObject, sheet2CopyFrom, sheet2PasteAt, calcObject2CopyFrom=None, calcObject2PasteAt=None):

    if type(sheet2CopyFrom) == str:
        sheetObject = calcObject.get_sheet_by_name(sheet2CopyFrom)
    else:
        sheetObject = sheet2CopyFrom

    if type(sheet2PasteAt) == str:
        sheetObject2 = calcObject.get_sheet_by_name(sheet2PasteAt)
    else:
        sheetObject2 = sheet2PasteAt

    # copy/paste data
    id_list = get_id(sheetObject)
    sheet_data = sheetObject.get_cell_range_by_position(0, 0, 10, len(id_list)+1).getDataArray()

    for row, row_data in enumerate(sheet_data):
        sheetObject2.set_columns_formula(0, row, row_data)
        for column in range(4, 10):  # copy/paste conditional formating
            sheetObject2.get_cell_by_position(column, row+1).ConditionalFormat = sheetObject.get_cell_by_position(column, row+1).ConditionalFormat

    # copy column Size
    cols = sheetObject.getColumns()
    cols2 = sheetObject2.getColumns()
    for col_number, col in enumerate(cols[0:11]):
        cols2[col_number].setPropertyValue('Width', col.Width)

    # fix color formating
    group_color(sheetObject2, calcObject=None)



def get_id(sheet, calcObject=None):

    if type(sheet) == str:
        sheetObject = calcObject.get_sheet_by_name(sheet)
    else:
        sheetObject = sheet

    # get id_list
    stop = False
    id_list = []
    row = 1
    while stop == False:
        id = sheetObject.get_cell_range_by_position(1, row, 1, row).getDataArray()

        if id[0][0] == '':
            stop = True
        else:
            id_list.append(id[0][0])
            row += 1
    return id_list




# %% specific
def loadCalc(sheet, calcObject=None):
    """Load xlsx file with fit parameters.

    Args:
        calcObject (Calc object): Object created by connect2calc().
        filename (str or pathlib.Path): path to xlsx file.

    Returns:
        parameter dictionary.
    """

    # connect to sheet
    if type(sheet) == str:
        sheetObject = calcObject.get_sheet_by_name(sheet)
    else:
        sheetObject = sheet

    # get data
    id_list = get_id(sheetObject)
    values = sheetObject.get_cell_range_by_position(0, 0, 10, len(id_list)+1).getDataArray()

    # separate data by group
    group_rows = get_group_rows(sheetObject)
    parameters = dict()
    for group in group_rows:
        for item, row in enumerate([group]):
            parameters[group] = dict()
            parameters[group]['id']          = [values[i][1] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['description'] = [values[i][2] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['dummy']       = [values[i][3] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['min']         = [values[i][4] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['guess']       = [values[i][5] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['max']         = [values[i][6] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['fit']         = [values[i][7] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['error']       = [values[i][8] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['warning']     = [values[i][9] for i in range(1, len(values)) if values[i][0] == group]
            parameters[group]['comments']    = [values[i][10] for i in range(1, len(values)) if values[i][0] == group]

    fixInf(parameters)
    fixNone(parameters)

    return sheetObject, parameters


def fixInf(parameters):
    """Substitute 'inf' by np.inf in parameter variable.

    Args:
        parameters (dict): parameter dict.
    """
    for group in parameters:
        for t in ('min', 'max'):
            matching = [j for j in range(len(parameters[group][t])) if 'inf' == parameters[group][t][j]]
            for k in matching:
                parameters[group][t][k] = np.inf
            matching = [j for j in range(len(parameters[group][t])) if '-inf' == parameters[group][t][j]]
            for k in matching:
                parameters[group][t][k] = -np.inf
    return


def fixNone(parameters):
    """Substitute None in max by +np.inf, min by -np.inf, and guess by 0 .

    Args:
        parameters (dict): parameter dict.
    """
    for group in parameters:
        t = 'min'
        matching = [j for j in range(len(parameters[group][t])) if parameters[group][t][j] == '']
        for k in matching:
            parameters[group][t][k] = -np.inf

        t = 'max'
        matching = [j for j in range(len(parameters[group][t])) if parameters[group][t][j] == '']
        for k in matching:
            parameters[group][t][k] = np.inf

        t = 'guess'
        matching = [j for j in range(len(parameters[group][t])) if parameters[group][t][j] == '']
        for k in matching:
            parameters[group][t][k] = 0
    return


def get_group(sheet, calcObject=None):

    if type(sheet) == str:
        sheetObject = calcObject.get_sheet_by_name(sheet)
    else:
        sheetObject = sheet

    # get id_list
    stop = False
    group_list = []
    row = 1
    while stop == False:
        group = sheetObject.get_cell_range_by_position(0, row, 0, row).getDataArray()

        if group[0][0] == '':
            stop = True
        else:
            group_list.append(group[0][0])
            row += 1
    return group_list


def get_group_rows(sheet, calcObject=None):

    if type(sheet) == str:
        sheetObject = calcObject.get_sheet_by_name(sheet)
    else:
        sheetObject = sheet

    group_list = get_group(sheetObject)
    group_rows = dict()
    for group in set(group_list):
        group_rows[group] = [i+1 for i,x in enumerate(group_list) if x==group]
    return group_rows


def update_xlsx(parameters, sheet, calcObject=None):

    # connect to sheet
    if type(sheet) == str:
        sheetObject = calcObject.get_sheet_by_name(sheet)
    else:
        sheetObject = sheet

    group_rows = get_group_rows(sheetObject)

    for group in group_rows:
        for item, row in enumerate(group_rows[group]):
            sheetObject.set_rows_formula(1, row, [parameters[group]['id'][item], ])
            sheetObject.set_rows_formula(2, row, [parameters[group]['description'][item], ])
            sheetObject.set_rows_formula(3, row, [parameters[group]['dummy'][item], ])

            if parameters[group]['min'][item] == -np.inf:
                sheetObject.set_rows_formula(4, row, ['-inf', ])
            else:
                sheetObject.set_rows_formula(4, row, [parameters[group]['min'][item], ])

            sheetObject.set_rows_formula(5, row, [parameters[group]['guess'][item], ])

            if parameters[group]['max'][item] == np.inf:
                sheetObject.set_rows_formula(6, row, ['inf', ])
            else:
                sheetObject.set_rows_formula(6, row, [parameters[group]['max'][item], ])

            sheetObject.set_rows_formula(7, row, [parameters[group]['fit'][item], ])
            sheetObject.set_rows_formula(8, row, [parameters[group]['error'][item], ])
            sheetObject.set_rows_formula(9, row, [parameters[group]['warning'][item], ])
            sheetObject.set_rows_formula(10, row, [parameters[group]['comments'][item], ])


def group_color(sheet, calcObject=None):

    if type(sheet) == str:
        sheetObject = calcObject.get_sheet_by_name(sheet)
    else:
        sheetObject = sheet

    # header old_color_max
    sheetObject.get_cell_range_by_position(0, 0, 10, 0).setPropertyValue('CellBackColor', 11711154)
    sheetObject.get_cell_range_by_position(0, 0, 10, 0).setPropertyValue('CharWeight', 150)  # bold


    group_rows = get_group(sheetObject)
    old_group = group_rows[0]
    color = -1
    for row, group in enumerate(group_rows):
        if group == old_group:
            sheetObject.get_cell_range_by_position(0, row+1, 10, row+1).setPropertyValue('CellBackColor', color)
        else:
            old_group = group_rows[row]
            if color == -1:
                color = 12771502
            else:
                color = -1
            sheetObject.get_cell_range_by_position(0, row+1, 10, row+1).setPropertyValue('CellBackColor', color)
