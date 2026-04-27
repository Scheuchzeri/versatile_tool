#---------------------------------------------------------------------------------------
# Name:        main.py
# Purpose:     Main operating logic for the 'versatile' application.
#
# Author:      I758972
#
# Created:     11/03/2026
# Copyright:   (c) I758972 2026
# Licence:     <conm>
#---------------------------------------------------------------------------------------


import src.engine as engine
import src.gui as gui
import src.lib as lib
import pythoncom
import os
import pathlib as pal
import time

def create_participants(app, config, data):
    """Triggers the creation of PT-Users in ILP over the ZLSO_VAP1 transaction based on
    an input dictionary.
    """

    lib.start_worker(app, config, data, _create_participant)

def _create_participant(app, config, data):
    """A helper function to create_participants."""

    pythoncom.CoInitialize()

    erp = data["erp"]
    first_name = data["first_name"]
    last_name = data["last_name"]
    corr_lang = data["corr_lang"]

    try:
        curr = handle_start(config, "connection_name_ilp")
        if curr is None:
            raise Exception("SAP Session could not be started.")

        curr.go_to("open_erp")
        target_erp_input = curr.resolve_action('ERP_INPUT')
        target_erp_input.text = erp
        curr.press_enter()

        curr.press_f5_key()

        target_first_name = curr.resolve_action('FIRST_NAME_INPUT')
        target_first_name.text = first_name

        target_last_name = curr.resolve_action('LAST_NAME_INPUT')
        target_last_name.text = last_name

        target_corr_lang = curr.resolve_action('CORR_LANGUAGE')
        target_corr_lang.text = corr_lang

        curr.press_shift_f1()

        app.after(0, lambda: app.output_manager.log(f"Relationships created sucessfully: {data}."))

    except Exception as e:
        app.after(0, lambda err=e: print(f"Thread Error: {err}"))
    finally:
        app.after(0, lambda: app.execute_btn.configure(text="Execute", state="normal"))


def check_for_participants(app, config, data):
    """This function checks the versatile-data folder and retrieves the 'relations_raw.txt'
    fill if available. It then checks the file against all input e-mail addresses and prints
    those that are not available to the Output-Manager.
    """

    dialog = gui.ctk.CTkInputDialog(text="Please provide an ERP number:", title="Test")
    erp = dialog.get_input()
    data = {"erp": erp, "emails": data}

    lib.start_worker(app, config, data, _export_rels)

def _export_rels(app, config, data):
    """This functions exports all relationships of an ERP as txt."""

    pythoncom.CoInitialize()
    erp = data["erp"]
    emails = data["emails"]

    try:
        curr = handle_start(config, "connection_name_ilp")
        if curr is None:
            raise Exception("SAP Session could not be started.")

        curr.go_to("open_erp")
        target_erp_input = curr.resolve_action('ERP_INPUT')
        target_erp_input.text = erp
        curr.press_enter()

        target_max_records = curr.resolve_action('MAX_RECORDS')
        if target_max_records:
            target_max_records.text = '5000'
            curr.press_enter()

        curr.resolve_action('EXPORT_RELATIONS')
        curr.select_radio_option('EXPORT_UNCONVERTED')
        curr.press_enter()

        script_path = os.path.dirname(os.path.abspath(__file__))
        target_path = f"{script_path}/data"

        target_dir_input = curr.resolve_action('DIR_INPUT')
        target_dir_input.text = target_path

        target_file_name = curr.resolve_action('FILE_NAME_INPUT')
        target_file_name.text = "relations_raw.txt"
        target_file_name.setFocus()
        curr.save(1)

        time.sleep(2)

        _compare_against_file(app, emails)

        app.after(0, lambda: app.output_manager.log(r"Relationships exported sucessfully to {script_path}\data"))

    except Exception as e:
        app.after(0, lambda err=e: print(f"Thread Error: {err}"))
    finally:
        app.after(0, lambda: app.execute_btn.configure(text="Execute", state="normal"))

def _compare_against_file(app, data):
    """A helper function that compares a list of e-mails against a relations-file and
    prints those entries to the OutputManager which have not yet been created under the
    respective ERP-Account.
    """

    script_path = pal.Path(__file__).parent
    file_path = script_path/"data"/"relations_raw.txt"
    rels_as_list = []

    if not file_path.exists():
        raise Exception("The file is not available. Please run 'Export Relationships' first.")

    with file_path.open("r", encoding="utf-8", errors="replace") as f:
        rels_as_list = [line.rstrip("\n") for line in f]

    tbc = []

    for addr in data:
        if any(addr.lower() in pt.lower() for pt in rels_as_list):
            continue
        tbc.append(addr)


    if len(tbc) == 0:
        app.output_manager.log("All of the participants already exist.")
    else:
        for pt in tbc:
            app.output_manager.log(f"{pt} has not yet been created.")

def list_person_ids(app, config, data):
    """This function accesses the SAPGUI ISP and retrieves the person-id of an internal
    participant.
    """
    lib.start_worker(app, config, data, _get_person_nrs)

def _get_person_nrs(app, config, data):
    """This is a sub-function to list_person_nrs . It performs the process-logic on the
    SAPGUI and returns the retrieved values for display.
    """
    pythoncom.CoInitialize()

    try:
        curr = handle_start(config, "connection_name_isp")
        if curr is None:
            raise Exception("SAP Session could not be started.")
        curr.go_to("abap_reporting")
        curr.run_abap_report("derive_person_nr")

        for addr in data:
            person_id = curr.derive_person_number(addr) or "[NO VALUE FOUND!]"
            curr.go_back()
            app.after(0, lambda r=person_id: app.output_manager.log(r))

        curr.clear()

    except Exception as e:
        app.after(0, lambda err=e: print(f"Thread Error: {err}"))
    finally:
        app.after(0, lambda: app.execute_btn.configure(text="Execute", state="normal"))

def handle_start(config, system):
    print("Main: Instructing Engine to start...")

    # 3. Extract the values
    # We use .get() to provide a fallback if the key is missing
    target_system = config.get("SAP", system, fallback="ILP [PUBLIC]")
    exe_path = config.get("SAP", "sap_logon_path", fallback=r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

    # 4. Initialize the SAP Engine with these specific values
    try:
        # We pass the strings directly into the Session constructor
        sap_session = engine.Session(connection_name=target_system, sap_path_config=exe_path)

        # Test if it worked
        if sap_session.session:
            print("Successfully connected using config.ini settings!")

    except Exception as e:
        print(f"Failed to initialize SAP: {e}")

    return sap_session

def main():

    config = lib.load_config()

    # Initialize the UI and pass the 'nerves' in


    tasks = {
    "ISP": {
        "Find Person-ID": {"runner": list_person_ids, "desc": "Retrieves PERNR from ISP"},
    },
    "ILP": {
        "Check for Participants": {"runner": check_for_participants, "desc": "If available this function will check the versatile-data folder for a relationships file and compare whether any of the input e-mail addresses match entries in the file."},
        "Create a Participant": {"runner": create_participant, "desc": "Creates PT-Users dynamically over the ZLSO_VAP1 transaction."}
    },
    }


    app = gui.MainUI(tasks, config)

    app.mainloop()

if __name__ == '__main__':
    main()
