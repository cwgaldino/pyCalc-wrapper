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
import warnings

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

        The main mathods defined for a Calc object are exemplyfied below:

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
    libreoffice = subprocess.Popen([f"soffice --nodefault --accept='socket,host=localhost,port={port};urp;'"], shell=True, close_fds=True)

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


def saveCalc(calcObject, filepath=None):
    """Save xlsx file.

    Note:
        If `filepath` have no suffix, it adds '.ods' at the end of filepath.

    Args:
        calcObject (Calc object): Object created by :py:func:`calcmanip.connect2Calc`.
        filepath (string or pathlib.Path, optional): filepath to save file.
    """
    if filepath is None:
        if calcObject.Location == '':
            filepath = Path('./Untitled.ods')
            warnings.warn('Saving at ./Untitled.ods')
        else:
            filepath = Path(calcObject.Location)
    else:
        filepath = Path(filepath)

    # fix extension
    if filepath.suffix == '':
        filepath = filepath.parent / (str(filepath.name) + '.ods')

    # save
    url = convert_path_to_url(str(filepath))
    calcObject.store_as_url(url, 'FilterName')



# calcObject manipulation
def get_sheets_name(calcObject):
    """Get sheets names in a tuple."""
    return calcObject.Sheets.ElementNames




def set_col_width(sheetObject, col, width):

    colsObject = sheetObject.getColumns()
    colsObject[col].setPropertyValue('Width', width)


def get_col_width(sheetObject, col):

    colsObject = sheetObject.getColumns()
    return colsObject[col].Width


def set_row_height(sheetObject, row, height):

    rowsObject = sheetObject.getRows()
    rowsObject[row].setPropertyValue('Height', height)


def get_row_height(sheetObject, row):

    colsObject = sheetObject.getRows()
    return colsObject[row].Height




def get_cell_value(sheetObject, row, col, type='formula'):
    """
    type='data', 'formula'
    """
    if type == 'formula':
        return sheetObject.get_cell_by_position(col, row).getFormula()
    elif type == 'data':
        return sheetObject.get_cell_by_position(col, row).getString()
    else:
        warnings.warn(f"type = {type} is not a valid option. Using type = 'data'.")
        return sheetObject.get_cell_by_position(col, row).getString()


def set_cell_value(sheetObject, row, col, value, type='formula'):
    """
    type='data', 'formula'
    """
    if type == 'formula':
        sheetObject.get_cell_by_position(col, row).setFormula(value)
    elif type == 'data':
        sheetObject.get_cell_by_position(col, row).setString(value)
    else:
        warnings.warn(f"type = {type} is not a valid option. Using type = 'data'.")
        sheetObject.get_cell_by_position(col, row).setString(value)


def copy_cell(sheet2copyFrom, sheet2pasteAt, row, col, type='formula',
              Font=1, ConditionalFormat=False, Border=False, resize=None,
              row2pasteAt=None, col2pasteAt=None, additional=None):
    """
    type='string', 'formula', None

    resize = None, 'r', 'c', 'rc' or 'cr'

    This function do not copy ALL the properties of a cell, because it is very
    time consuming. Instead, it copys only the most used properties. If you
    need to include additional properties, have a look at
    ``sheetObject.get_cell_by_position(0, 0)._show_attributes()`` and find the
    desired propertie. Then, include it in ``additional``.
    """
    Font = int(Font)
    if Font > 4:
        Font = 4
    elif Font <0:
        Font = 0

    if row2pasteAt is None:
        row2pasteAt = row
    if col2pasteAt is None:
        col2pasteAt = col

    # cell value
    if type is not None:
        set_cell_value(sheet2pasteAt, row=row2pasteAt, col=col2pasteAt, value=get_cell_value(sheet2copyFrom, row, col, type=type), type=type)

    # font name
    font_property_list_parsed = [['FormatID', 'CharWeight', 'CharHeight', 'CharColor', 'CellBackColor'],
                                 [ 'CharFontName',  'CharFont', 'CellStyle'],
                                 ['CharUnderline', 'CharCrossedOut', 'CharEmphasis', 'CharEscapement', 'CharContoured'],
                                 ['CharPosture',  'CharPostureComplex',  'CharRelief',  'CharShadowed',  'CharStrikeout',   'CharUnderlineColor',  'CharUnderlineHasColor',]
                                ]

    font_property_list = [item for sublist in font_property_list_parsed[0:Font] for item in sublist]
    for property in font_property_list:
        sheet2pasteAt.get_cell_by_position(col2pasteAt, row2pasteAt).setPropertyValue(property, getattr(sheet2copyFrom.get_cell_by_position(col, row), property))

    # conditional formating
    if ConditionalFormat:
        font_property_list = ['ConditionalFormat']
        for property in font_property_list:
            sheet2pasteAt.get_cell_by_position(col2pasteAt, row2pasteAt).setPropertyValue(property, getattr(sheet2copyFrom.get_cell_by_position(col, row), property))

    # border
    if Border:
        border_property_list = ['TableBorder', 'TableBorder2']#, 'LeftBorder', 'LeftBorder2', 'RightBorder', 'RightBorder2', 'TopBorder', 'TopBorder2', 'BottomBorder', 'BottomBorder2']
        for property in border_property_list:
            sheet2pasteAt.get_cell_by_position(col2pasteAt, row2pasteAt).setPropertyValue(property, getattr(sheet2copyFrom.get_cell_by_position(col, row), property))

    # additional
    if additional is not None:
        for property in additional:
            sheet2pasteAt.get_cell_by_position(col2pasteAt, row2pasteAt).setPropertyValue(property, getattr(sheet2copyFrom.get_cell_by_position(col, row), property))

    # col and row width
    if resize is not None:
        if resize == 'r':
            set_row_height(sheet2pasteAt, row2pasteAt, get_row_height(sheet2copyFrom, row))
        elif resize == 'c':
            set_col_width(sheet2pasteAt, col2pasteAt, get_col_width(sheet2copyFrom, col))
        elif resize == 'cr' or resize == 'rc':
            set_row_height(sheet2pasteAt, row2pasteAt, get_row_height(sheet2copyFrom, row))
            set_col_width(sheet2pasteAt, col2pasteAt, get_col_width(sheet2copyFrom, col))
        else:
            warnings.warn(f"resize = {resize} is not a valid option. Using resize = None.")



def get_cells_value(sheetObject, row_init, col_init, row_final, col_final, type='data'):
    """
    type= formula or data.
    """
    sheet_data = sheetObject.get_cell_range_by_position(row_init, col_init, row_final, col_final)
    if type == 'formula':
        return sheet_data.getFormulaArray()
    elif type == 'data':
        return sheet_data.getDataArray()
    else:
        warnings.warn(f"type = {type} is not a valid option. Using type = 'data'.")
        return sheet_data.getDataArray()


def set_cells_value(sheetObject, row_init, col_init, data, type='formula'):
    """
    type=formula or data.

    another option would be value: sheet444.set_columns_value(x, y, data)
    """

    if type == 'formula':
        for row, row_data in enumerate(data):
            sheetObject.set_columns_formula(row_init, row+col_init, row_data)
    elif type == 'data':
        for row, row_data in enumerate(data):
            sheetObject.set_columns_str(row_init, row+col_init, row_data)
    else:
        warnings.warn(f"type = {type} is not a valid option. Using type = 'data'.")
        for row, row_data in enumerate(data):
            sheetObject.set_columns_str(row_init, row+col_init, row_data)


def copy_cells(sheet2copyFrom, sheet2pasteAt, row_init, col_init, row_final, col_final, type='formula',
              Font=0, ConditionalFormat=False, Border=False, resize=None,
              row2pasteAt=None, col2pasteAt=None, additional=None):
    """
        type='data', 'formula', 'none'
    """

    if row2pasteAt is None:
        row2pasteAt = row_init
    if col2pasteAt is None:
        col2pasteAt = col_init

    if Font>0 or ConditionalFormat is not False or Border is not False or additional is not False:
        for row_relative, row in enumerate(range(row_init, row_final)):
            for col_relative, col in enumerate(range(col_init, col_final)):
                copy_cell(sheet2copyFrom, sheet2pasteAt, row, col, type=type,
                          Font=Font, ConditionalFormat=ConditionalFormat, Border=Border, resize=None,
                          row2pasteAt=row2pasteAt+row_relative, col2pasteAt=col2pasteAt+col_relative, additional=additional)
    else:
        data = get_cells_value(sheet2copyFrom, row_init, col_init, row_final, col_final, type=type)
        set_cells_value(row2pasteAt, row2pasteAt, col2pasteAt, data, type=type)

    # col and row width
    if resize is not None:
        if resize == 'r' or resize == 'c' or resize == 'cr' or resize == 'rc':
            if 'r' in resize:
                for row_relative, row in enumerate(range(row_init, row_final)):
                    set_row_height(sheet2pasteAt, row2pasteAt+row_relative, get_row_height(sheet2copyFrom, row))
            if 'c' in resize:
                for col_relative, col in enumerate(range(col_init, col_final)):
                    set_col_width(sheet2pasteAt, col2pasteAt+col_relative, get_col_width(sheet2copyFrom, col))
        else:
            warnings.warn(f"resize = {resize} is not a valid option. Using resize = None.")


def copy_sheet(sheet2copy, sheet2paste, type='formula',
              Font=0, ConditionalFormat=False, Border=False, resize=None, additional=None):
    """"
    """"
    last_col = len(sheet2copy.getColumnDescriptions())
    last_row = len(sheet2copy.getRowDescriptions())

    copy_cells(sheet2copy, sheet2paste, 0, 0, last_row, last_col, type=type, Font=Font, ConditionalFormat=ConditionalFormat, Border=Border, resize=resize, additional=None)


def get_cell_value_from_sheets(sheetObject_list, row, col, type='data'):
    """
    """
    values = []
    for sheetObject in sheetObject_list:
        values.append(get_cell_value(sheetObject, row, col, type))
    return values


# %% specific
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
