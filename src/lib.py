#---------------------------------------------------------------------------------------
# Name:        lib.py
# Purpose:     Collection of custom-functions, classes and methods for operating ILP over
#              the SAPGUI. Part of the 'versatile' application.
#
# Author:      I758972
#
# Created:     11/03/2026
# Copyright:   (c) I758972 2026
# Licence:     <conm>
#---------------------------------------------------------------------------------------

import configparser
import os
import threading

# Central registry for transactions
TRANS_ACTIONS = {
    "abap_reporting": "SA38 - ABAP REPORTING",
    "derive_person_nr": "ZCXPER_DERIVE_PERNR",
    "open_erp": "ZLSO_VAP1"
}

ACTION_MAP = {
  # --- First window elements
  # --- Regular elements (resolved via findById) ---
  "ERP_INPUT": {
      "type": "element",
      "match": "ctxtRF02D-KUNNR",        # stable suffix of the element ID
      "wnd_idx": 0,
  },
  "MAX_RECORDS": {
      "type": "element",
      "match": "txtBURS_JOEL_SELECTION-MAX_RECORD",
      "wnd_idx": 0,
  },
  "FIRST_NAME_INPUT": {
      "type": "element",
      "match": "BUT000-NAME_FIRST",
      "wnd_idx": 0,
  },
  "LAST_NAME_INPUT": {
      "type": "element",
      "match": "BUT000-NAME_LAST",
      "wnd_idx": 0,
  },
  "CORR_LANGUAGE":{
      "type": "element",
      "match": "BUS000FLDS-LANGUCORR",
      "wnd_idx": 0,
  },
  "RELATIONS_CAT": {
      "type": "element",
      "match": "cmbBURS_JOEL_MAIN-DIRECTED_TYPE_C",
      "wnd_idx": 0,
  },
  # --- Shell toolbar buttons (resolved via shell lookup + pressToolbarButton) ---
  "EXPORT_RELATIONS": {
      "type": "shell_btn",
      "shell_match": "RIGHT_AREA",        # substring to identify the correct shell
      "btn_id": "&MB_EXPORT",             # stable button ID string
      "sub_btn_id": "&PC",
      "wnd_idx": 0,
  },
  "EXPORT_LOCAL_FILE": {
      "type": "shell_btn",
      "shell_match": "RIGHT_AREA",
      "btn_id": "&PC",
      "wnd_idx": 0,
  },
  # --- Second window selection elements ---
  # --- Regular elements (resolved via findById) ---
  "DIR_INPUT": {
      "type": "element",
      "match": "ctxtDY_PATH",
      "wnd_idx": 1
  },
  "FILE_NAME_INPUT": {
      "type": "element",
      "match": "ctxtDY_FILENAME",
      "wnd_idx": 1,
  },
  # --- Radio Buttons ---
  "EXPORT_UNCONVERTED": {
      "type": "radio_btn",
      "btn_label": "Unconverted",
      "wnd_idx": 1,
  },
  "EXPORT_WITH_TABS": {
      "type": "radio_btn",
      "btn_label": "Text with Tabs",
      "wnd_idx": 1,
  },
}

SAP_MAP = {
    # SA38 Screen Technical IDs
    "SA38_INPUT_FIELD": "wnd[0]/usr/ctxtRS38M-PROGRAMM",
    "SA38_EXECUTE_BTN": "wnd[0]/tbar[1]/btn[8]",

    "ZCXPER_DERIVE_PERNR": "wnd[0]/usr/txtUSERID",

    # Navigating ILP
    # Within ZLSO_VAP1
    "ZLSO_ERP_INPUT_FIELD": "wnd[0]/usr/ctxtRF02D-KUNNR",
    "NO_OF_PRS_DISPLAYED": "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1200/tabsGS_SCREEN_1200_TABSTRIP/tabpSCREEN_1200_TAB_01/ssubSCREEN_1200_TABSTRIP_AREA:SAPLBUPA_DIALOG_JOEL:1218/ssubSCREEN_1210_SELECTION_AREA:SAPLBUPA_DIALOG_JOEL:1260/txtBURS_JOEL_SELECTION-MAX_RECORD",
    "EXPORT_RELATIONS_CTXT": "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1200/tabsGS_SCREEN_1200_TABSTRIP/tabpSCREEN_1200_TAB_01/ssubSCREEN_1200_TABSTRIP_AREA:SAPLBUPA_DIALOG_JOEL:1218/subSCREEN_1210_OVERVIEW_AREA:SAPLBUPA_DIALOG_JOEL:1220/cntlSCREEN_1220_CUSTOM_CONTROL/shellcont/shell",
    "TRIAL_VALUE": "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1200/tabsGS_SCREEN_1200_TABSTRIP/tabpSCREEN_1200_TAB_01/ssubSCREEN_1200_TABSTRIP_AREA:SAPLBUPA_DIALOG_JOEL:1210/subSCREEN_1210_OVERVIEW_AREA:SAPLBUPA_DIALOG_JOEL:1220/cntlSCREEN_1220_CUSTOM_CONTROL/shellcont/shell",
    "EXPORT_RELATIONS_BTN": "&MB_EXPORT",
    "SELECT_EXPORT_CONTEXT": "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2036/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1200/tabsGS_SCREEN_1200_TABSTRIP/tabpSCREEN_1200_TAB_01/ssubSCREEN_1200_TABSTRIP_AREA:SAPLBUPA_DIALOG_JOEL:1218/subSCREEN_1210_OVERVIEW_AREA:SAPLBUPA_DIALOG_JOEL:1220/cntlSCREEN_1220_CUSTOM_CONTROL/shellcont/shell",
    "SELECT_LOCAL_FILE": "&PC",
    "SELECT_UNCONVERTED_FILE": "wnd[1]/usr/ctxtDY_FILENAME",

    "FILENAME_INPUT_FIELD": "wnd[1]/usr/ctxtDY_FILENAME",
    "FILEPATH_INPUT_FIELD": "wnd[1]/usr/ctxtDY_PATH",


    # Common Toolbar IDs (Shared across many screens)
    "BTN_BACK": "wnd[0]/tbar[0]/btn[3]",
    "BTN_EXE": "wnd[0]/tbar[1]/btn[8]",
    "BTN_SAVE": "wnd[0]/tbar[0]/btn[11]",

    "1ST_BTN_2ND_WNDW": "wnd[1]/tbar[0]/btn[0]"
}

#class SAP_MAP:
#
#    # SA38 Screen Technical IDs
#    SA38_INPUT_FIELD = "wnd[0]/usr/ctxtRS38M-PROGRAMM"
#    SA38_EXECUTE_BTN = "wnd[0]/tbar[1]/btn[8]"
#
#    ZCXPER_DERIVE_PERNR = "wnd[0]/usr/txtUSERID"
#
#    # Navigating ILP
#    # Within ZLSO_VAP1
#    ZLSO_ERP_INPUT_FIELD = "wnd[0]/usr/ctxtRF02D-KUNNR"
#
#    # Common Toolbar IDs (Shared across many screens)
#    BTN_BACK = "wnd[0]/tbar[0]/btn[3]"
#    BTN_EXE  = "wnd[0]/tbar[1]/btn[8]"
#    BTN_SAVE = "wnd[0]/tbar[0]/btn[11]"

def start_tesses(conn, path):
    """A function for quickly enabling a test-session in the py-interpreter."""
    pass

def start_worker(app, config, data, target):
    """Initializes a worker-thread to handle SAPGUI sessions."""
    worker = threading.Thread(
        target=target,
        args=(app, config, data),
        daemon=True # Ensures thread dies if GUI is closed
    )
    worker.start()

def load_config(file_path="data/config.ini"):
    """Loads the config.ini file for this application."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Config file not found at {file_path}")

    config = configparser.ConfigParser()
    config.read(file_path)
    return config

def debug_children(sap_session):
    usr = sap_session.session.findById("wnd[0]/usr")
    print(f"Found {usr.Children.Count} children in the UserArea.")
    for i in range(usr.Children.Count):
        child = usr.Children.Item(i)
        # Print the Type and the Text of every object found
        text = child.Text if hasattr(child, "Text") else "NO TEXT"
        print(f"Child[{i}] Type: {child.Type} | Text: {text}")

