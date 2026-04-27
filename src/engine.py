#---------------------------------------------------------------------------------------
# This Python file uses the following encoding: utf-8
# Name:        engine.py
# Purpose:     Engine for calling and accessing the SAPGUI from windows. Part of the
#              'versatile' application.
#
# Author:      I758972
#
# Created:     11/03/2026
# Copyright:   (c) I758972 2026
# Licence:     <conm>
#---------------------------------------------------------------------------------------

import win32com.client
import pythoncom
import sys
import subprocess
import time
import psutil
import src.lib as lib

class Session:

    def __init__(self, connection_name="ILP [PUBLIC]", sap_path_config="C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"):
        self.connection_name = connection_name
        self.session = self._get_or_create_session(sap_path_config)

        if self.session:
            print(f"Connected to: {self.session.Info.SystemName} (Client: {self.session.Info.Client})")
        else:
            raise Exception("Could not establish SAP Session.")

    def _get_or_create_session(self, sap_path_config):
        try:
            # 1. Ensure SAP Logon is running
            self._ensure_sap_is_running(sap_path_config)

            # 2. Access the GUI Engine
            sap_gui_auto = win32com.client.GetObject("SAPGUI") # Setup in config file in case app is called differently on different computers.
            application = sap_gui_auto.GetScriptingEngine

            # 3. Look for an existing session first
            for conn in application.Children:
                # Check if this connection matches your target
                if self.connection_name in conn.Description:
                    if conn.Children.Count > 0:
                        # Grab the first available session in that connection
                        return conn.Children(0)

            # 4. If not found, open a new one
            print(f"No active session found for {self.connection_name}. Opening new...")
            connection = application.OpenConnection(self.connection_name, True)
            time.sleep(2) # Allow UI to initialize
            return connection.Children(0)

        except Exception as e:
            print(f"Critical Connection Error: {e}")
            return None

    def _ensure_sap_is_running(self, sap_path_config):
        """Checks if saplogon.exe is in the process list, starts it if not."""
        import psutil # 'pip install psutil' is highly recommended for robustness

        sap_running = any("saplogon.exe" in p.name().lower() for p in psutil.process_iter())

        if not sap_running:
            print('SAP is not running.')
            sap_path = sap_path_config # / How to generalize path?
            subprocess.Popen(sap_path)
            time.sleep(5) # Give it time to boot

    def go_to(self, transaction_alias):
        """Takes an alias to a transaction-code and navigates by consulting lib.py."""
        t_code = lib.TRANS_ACTIONS.get(transaction_alias)
        if not t_code:
            raise ValueError(f"Transaction alias '{transaction_alias}' not found.")

        self.session.StartTransaction(t_code)

    def save(self, wnd_idx):
        """Save file or process."""
        self.session.findById(f"wnd[{wnd_idx}]").sendVKey(11)

    def press_enter(self):
        """Sends the enter-key command to current window."""
        self.session.findById("wnd[0]").sendVKey(0)

    def clear(self):
        # Using .findById is safer than hardcoded strings if the UI structure changes
        self.session.findById("wnd[0]/tbar[0]/okcd").text = '/n'                # Is it possible to abstract these identification parameters?
        self.session.findById("wnd[0]").sendVKey(0)

    def cancel(self):
        """General purpose cancel command."""
        self.session.findById("wnd[0]").sendVKey(12)

    def go_back(self):
        """Goes back once instance in the GUI."""
        self.session.findById("wnd[0]").sendVKey(3)

    def press_f5_key(self):
        """Sends the F5 signal to the current sesssion. In the ZLSO_VAP1 transaction
        this will create a new contact person."""
        self.session.findById("wnd[0]").sendVKey(5)

    def press_shift_f1(self):
        """Sends the Shift + F1 signal to the current session."""
        self.session.findById("wnd[0]").sendVKey(13)

    def _identify_target(self, key):
        """This helper function scans traverses the active session's view-structure and
        returns the ID of the element to be target by subsequent transactions based on
        a configured dictionary.
        """

        while self.session.Busy:
            time.sleep(0.2)

        available = self.return_view_structure()
        for el in available:
            if key in el:
                target = el
            else:
                continue

        print(target)

        if not target:
            return
        else:
            return target

    def select_radio_option(self, key):
        """Resolves a radio_btn entry from ACTION_MAP, scans the target window for a
        matching label, and selects it. Raises ValueError if no match found.
        """
        entry = lib.ACTION_MAP.get(key)
        if not entry or entry["type"] != "radio_btn":
            raise ValueError(f"Key '{key}' is not a valid radio_btn entry.")

        wnd_idx = entry.get("wnd_idx")
        target_label = entry["btn_label"]

        while self.session.Busy:
            time.sleep(0.2)

        user_area = self.session.findById(f"wnd[{wnd_idx}]/usr")
        view = self._return_sub_elements_extended(user_area)

        for el_id in view["elements"]:
            el = self.session.findById(el_id)
            if el.Type == "GuiRadioButton" and el.Text.strip() == target_label:
                el.select()
                return

        raise ValueError(f"No radio button with label '{target_label}' found in wnd[{wnd_idx}].")

    def resolve_action(self, key):
        """Looks up key in ACTION_MAP, crawls the live view, and returns either
        a findById target (for 'element' type) or a shell object (for 'shell_btn' type).
        Raises ValueError with a clear message if the key is unknown or no match is found.
        """
        entry = lib.ACTION_MAP.get(key)
        if not entry:
            raise ValueError(f"Unknown action key: '{key}'")

        while self.session.Busy:
            time.sleep(0.2)

        window_idx = entry["wnd_idx"]
        view = self.return_view_structure_extended(window_idx)

        if entry["type"] == "element":
            match_token = entry["match"]
            matches = [el for el in view["elements"] if match_token in el]
            if not matches:
                print(f"No element found matching token: '{match_token}'")
                return
            if len(matches) > 1:
                raise ValueError(f"Ambiguous match for token '{match_token}': {matches}")
            return self.session.findById(matches[0])

        elif entry["type"] == "shell_btn":
            shell_token = entry["shell_match"]
            matches = [s for s in view["shells"] if shell_token in s]
        if not matches:
            raise ValueError(f"No shell found matching token: '{shell_token}'")
        if len(matches) > 1:
            raise ValueError(f"Ambiguous shell match for token '{shell_token}': {matches}")
        shell = self.session.findById(matches[0])
        shell.pressToolbarButton(entry["btn_id"])
        shell.selectContextMenuItem(entry["sub_btn_id"])
        return

    def return_view_structure_extended(self, window_idx=0):
        """Like return_view_structure, but also includes GuiShell objects.
        Returns a dict with two keys:
        'elements' -> list of ID strings for standard interactive elements
        'shells'   -> list of live GuiShell COM object references
        """

        idx = window_idx

        user_area = self.session.findById(f"wnd[{idx}]/usr")
        return self._return_sub_elements_extended(user_area)

    def _return_sub_elements_extended(self, container):
        interactive_types = {
          "GuiTextField", "GuiCTextField", "GuiButton",
          "GuiCheckBox", "GuiRadioButton", "GuiComboBox"
        }

        parsed_ids = []
        shell_ids = []

        for i in range(container.Children.Count):
            child = container.Children(i)

            if child.Type in interactive_types:
                parsed_ids.append(child.Id)

            elif child.Type == "GuiShell":
                shell_ids.append(child.Id)

            if hasattr(child, "Children"):
                sub = self._return_sub_elements_extended(child)
                parsed_ids.extend(sub["elements"])
                shell_ids.extend(sub["shells"])

        return {"elements": parsed_ids, "shells": shell_ids}

    def select_context_menu(self, ctxt, btn):

        while self.session.Busy:
            time.sleep(0.2)

        context = lib.SAP_MAP.get(ctxt)
        self.session.findById(context).pressToolbarContextButton(btn)

    def export_unconverted_local_file(self, path_input_field, file_input_field, script_path):
        """This method exports an unconverted, local file from the ERP-account view in
        ILP, transporting the relationships pertaining to that account as an unconverted
        txt file to the specified directory.
        """

        path_input_field = lib.SAP_MAP.get(path_input_field)
        file_input_field = lib.SAP_MAP.get(file_input_field)
        target_path = f"{script_path}\data"
        file_name = "relations_raw.txt"

        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById(path_input_field).text = target_path
        self.session.findById(file_input_field).text = file_name
        self.session.findById(file_input_field).setFocus

        while self.session.Busy:
            time.sleep(1)

        self.session.findById("wnd[1]").sendVKey(11)


    def select_local_file_for_export(self, ctxt, btn):

        while self.session.Busy:
            time.sleep(0.2)

        context = lib.SAP_MAP.get(ctxt)
        btn = lib.SAP_MAP.get(btn)

        self.session.findById(context).selectContextMenuItem(btn)

    def press_toolbar_btn(self, key):

        while self.session.Busy:
            time.sleep(0.2)

        target_id = self._identify_target(key)

        if not target_id:
            raise ValueError(f"The target element is not available: {key}")

        target = self.session.findById(target_id)
        target.setFocus()

        self.press_enter()

    def select_and_trigger_input(self, key, data):

       while self.session.Busy:
        time.sleep(0.2)

       target_id = self._identify_target(key)
       if not target_id:
        raise ValueError(f"The target element is not available: {key}")


       try:
        # 2. Re-verify the object exists right now
        target = self.session.findById(target_id)

        # 3. Explicitly set focus before changing text
        target.setFocus()
        target.text = data[0] # Ensure it's a string

        self.press_enter()

       except Exception as e:
        print(f"SAP Error: {e}")
        return
        # If it fails here, the session might need to be refreshed

    def get_user_name(self):
        """Example: Reading data from the status bar."""
        return self.session.Info.User

    def get_person_id_from_children(self):
        """Identifies the 'is:' label and retrieves the value from the
        subsequent child object.
        """
        try:
            usr = self.session.findById("wnd[0]/usr")
            count = usr.Children.Count

            for i in range(count):
                child = usr.Children.Item(i)

                # Find the anchor label
                if hasattr(child, "Text") and "is:" in child.Text:
                    # Based on your debug, the number is 2 indexes ahead
                    # Child[6] is "is:", Child[7] is empty, Child[8] is the ID
                    target_index = i + 2

                    if target_index < count:
                        target_child = usr.Children.Item(target_index)
                        person_id = target_child.Text.strip()
                        print(f"Engine: Successfully extracted ID: {person_id}")
                        return person_id

            print("Engine: Anchor 'is:' not found or list is empty.")
            return None
        except Exception as e:
            print(f"Positional extraction failed: {e}")
            return None

    def get_person_id_via_clipboard(self):
        try:
            # 1. Select the entire list
            # %pc is the command for "Save to local file", but
            # Ctrl+A / Ctrl+C is easier if the Grid/List allows it.
            # Alternatively, use the 'System' menu:
            self.session.findById("wnd[0]").sendVKey(5) # F5 is often 'Select All'

            # 2. Use the 'Copy' Command
            self.session.findById("wnd[0]/tbar[1]/btn[16]").press() # Typical 'Copy' btn

            # 3. Access clipboard via Python
            root = tk.Tk()
            root.withdraw() # Hide the tiny window
            clipboard_text = root.clipboard_get()
            root.destroy()

        except Exception as e:
            print(f"Clipboard method failed: {e}")
        return None

    def find_grid_dynamically(self):
        """Crawls the 'user' area of the current window to find a GridView
        regardless of its technical ID.
        """
        try:
            # Start looking in the main user area
            user_area = self.session.findById("wnd[0]/usr")
            return self._search_children(user_area)
        except Exception as e:
            print(f"Search failed: {e}")
            return None

    def _search_children(self, parent):
        """Recursive helper to find the first GridView"""
        for child in parent.Children:
            # If it's the grid, return it
            if child.Type == "GuiShell" and child.SubType == "GridView":
                return child

            # If it has children, dive deeper (like into containers/splitters)
            if hasattr(child, "Children"):
                result = self._search_children(child)
                if result:
                    return result
        return None

    def get_grid_value(self, row=0, column="PERNR"):
        grid = self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        value = grid.getCellValue(row, column)
        return value

    def return_view_structure(self):
        """Similar to log_view_structure only that it returns the elements' ids as a list."""
        user_area = self.session.findById("wnd[0]/usr")
        ids = self._return_sub_elements(user_area)
        return ids

    def _return_sub_elements(self, container):
        interactive_types = {"GuiTextField","GuiCTextField","GuiButton","GuiCheckBox","GuiRadioButton","GuiComboBox"}

        parsed_ids = []

        for i in range(container.Children.Count):
            child = container.Children(i)
            if child.Type in interactive_types:
                parsed_ids.append(child.Id)
            if hasattr(child, "Children"):
                parsed_ids.extend(self._return_sub_elements(child))
        return parsed_ids

    def log_view_structure(self):
        """
        Interrogates the current screen and prints all interactive
        elements and their SAP IDs.
        """
        print(f"\n--- Interrogating View: {self.session.Info.Transaction} ---")
        # Start at the top-level window user area
        user_area = self.session.findById("wnd[0]")
        self._parse_sub_elements(user_area)
        print("--- End of View Structure ---\n")

    def _parse_sub_elements(self, container):
        """
        Recursive helper to crawl the SAP GUI tree.
        """
        for i in range(container.Children.Count):
            child = container.Children(i)

            # Filter for types that are usually 'Input' or 'Interactive'
            # SAP Types: https://help.sap.com/viewer/6522ef2730ca45719365eab4050945a4/7.60.4/en-US
            interactive_types = ["GuiTextField", "GuiCTextField", "GuiButton",
                                 "GuiCheckBox", "GuiRadioButton", "GuiComboBox"]
            if child.Type in interactive_types:
                # Some elements have labels, others have 'Text' (value)
                label = getattr(child, 'Text', 'N/A')
                print(f"Type: {child.Type:15} | ID: {child.Id:40} | Value/Label: {label}")


            # If this child is a container (like a Tab or GroupBox), dive deeper
            if hasattr(child, "Children"):
                self._parse_sub_elements(child)

    def create_pt_user(self, data):
        """This method accepts a dictionary as input which must containt at least the following:
        first-name, last-name, email and based on this input creates a pt-user in ILP
        over the ZLSO_VAP1 transaction.
        """
        pass

    def run_abap_report(self, report_name):
        """Runs a specific report within the 'SA38 - ABAP Reporting' transaction."""
        # The ID is pulled from lib.py
        # The value (specific report) is passed by main.py from the config-file.

        report = lib.TRANS_ACTIONS.get(report_name)

        input_field = lib.SAP_MAP.get("SA38_INPUT_FIELD")
        execute_btn = lib.SAP_MAP.get("SA38_EXECUTE_BTN")

        self.session.findById(input_field).text = report
        self.session.findById(execute_btn).press()

    def derive_person_number(self, employee_id):
        """Runs the ZCXPER_DERIVE_PERNR report in ISP and returns the employee's internal
        user-number."""

        input_field = lib.SAP_MAP.get("ZCXPER_DERIVE_PERNR")
        exe_btn = lib.SAP_MAP.get("BTN_EXE")

        self.session.findById(input_field).text = employee_id
        self.session.findById(exe_btn).press()

        return self.get_person_id_from_children()

def call_ilp_s(conn="ILP [PUBLIC]", path="C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"):
    """Starts a default session for testing and debugging in the py-interpreter."""
    curr = Session(conn, path)
    return curr


# path = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
# isp = "ISP [PUBLIC] (001)"