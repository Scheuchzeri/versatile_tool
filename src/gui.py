#---------------------------------------------------------------------------------------
# Name:        gui.py
# Purpose:     Setup for the graphical user-interface of the 'versatile' application.
#
# Author:      I758972
#
# Created:     11/03/2026
# Copyright:   (c) I758972 2026
# Licence:     <conm>
#---------------------------------------------------------------------------------------

import customtkinter as ctk
import re

class MainUI(ctk.CTk):
    def __init__(self, tasks, config):
        super().__init__()
        self.tasks = tasks
        self.active_system = None
        self.active_task = None
        self.config = config

        # Window Configuration
        self.title("Versatile")
        self.geometry("1200x1050")

        # Grid Configuration (3 columns, multiple rows)
        self.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)

        # Header
        self.header_label = ctk.CTkLabel(self, text="SAP ILP", font=("Arial", 20, "bold"))
        self.header_label.grid(row=0, column=2, padx=20, pady=20)

        # --- Sidebar: System Selection ---
        self.sidebar = ctk.CTkFrame(self, width=200)
        self.sidebar.grid(row=0, column=0, rowspan=7, sticky="nsew")

        for i, system_name in enumerate(self.tasks.keys()):
            btn = ctk.CTkButton(
                self.sidebar, text=system_name,
                command=lambda s=system_name: self.update_task_menu(s)
            )
            btn.pack(pady=10, padx=10)

        # --- Center: Task Selection ---

        self.task_selector = ctk.CTkOptionMenu(
            self, values=["---"], command=self.set_active_task
        )
        self.task_selector.grid(row=1, column=2, pady=10)

        # Create the Input-Manager on the left side of the window
        self.input_manager = InputManager(master=self, width=400, height=500)
        self.input_manager.grid(row=0, column=1, rowspan=6, padx=20, pady=20, sticky="nsew")
        # Create Input Cells

        # Output-Manager
        self.output_manager = OutputManager(master=self, width=400, height=500)
        self.output_manager.grid(row=0, column=3, rowspan=6, padx=20, pady=20, sticky="nsew")

        # The Floating Copy Button
        # We place it in the same grid cell (row 0, col 3)
        # and use sticky="ne" to float it in the top right corner.
        self.copy_btn = ctk.CTkButton(
            self,
            text="Copy All",
            width=70,
            height=22,
            font=("Arial", 11),
            command=self.output_manager.copy_all_to_clipboard
        )
        # pady=28 aligns it roughly with the text of the "Automation Log" header
        self.copy_btn.grid(row=0, column=3, sticky="ne", padx=30, pady=28)

        # The Universal Execute Button
        self.execute_btn = ctk.CTkButton(
            self, text="Execute", state="disabled",
            fg_color="green", command=self.run_current_automation
        )
        self.execute_btn.grid(row=5, column=1, pady=20)

        # Exit Button (Red for clarity)
        self.exit_btn = ctk.CTkButton(self, text="Exit", fg_color="firebrick", hover_color="#8B0000",
                                      command=self.quit)
        self.exit_btn.grid(row=6, column=3, padx=20, pady=20)

    def update_task_menu(self, system_name):
        """Triggered when a System (ISP/ILP) is clicked."""
        self.active_system = system_name
        available_tasks = list(self.tasks[system_name].keys())

        # Update the dropdown menu with tasks specific to that system
        self.task_selector.configure(values=available_tasks)
        self.task_selector.set(available_tasks[0]) # Default to first task
        self.set_active_task(available_tasks[0])

        self.header_label.configure(text=f"{self.active_system}")

        self.output_manager.log(f"System changed to {system_name}", "info")

    def set_active_task(self, task_name):
        """Triggered when a specific Transaction is chosen from the dropdown."""
        self.active_task = task_name
        task_info = self.tasks[self.active_system][task_name]

        # Store the runner function
        self.active_runner = task_info["runner"]

        # Update UI feedback

        self.execute_btn.configure(state="normal", text=f"Run {task_name}")

    def switch_context(self, mode_name):
        """Prepares the UI for a specific SAP transaction."""
        mode_data = self.tasks.get(mode_name)

        # 1. Update Header and internal state
        self.header_label.configure(text=f"SAP: {mode_name}")
        self.active_runner = mode_data["runner"]

        # 2. Reset/Prepare Input/Output
        # You might want to clear previous data when switching modes
        # self.input_manager.clear_all()

        # 3. Enable the Execute Button
        self.execute_btn.configure(state="normal", text=f"Run {mode_name}")

    def run_current_automation(self):
        """Gathers data and fires the backend function."""
        if not self.active_runner:
            return

        self.output_manager.clear_all()

        data_list = self.input_manager.get_all_values()
        if not data_list:
            self.output_manager.log("Error: No input data provided.", "error")
            return

        # Disable UI during run to prevent double-clicks
        self.execute_btn.configure(state="disabled", text="Processing...")

        # Fire the runner (main.py handles the threading)
        self.active_runner(self, self.config, data_list)


    def populate_output(self, data):
        """Accepts data returned from the SAPGUI via main.py and popultes the output-frame
        of the GUI with it."""
        for el in data:
            self.output_manager.log(el)

class InputRow(ctk.CTkFrame):
    def __init__(self, master, delete_callback, paste_func, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        # 2. Create the Entry (Notice: NO paste_func inside these parentheses)
        self.entry = ctk.CTkEntry(
            self,
            width=300,
            placeholder_text="Enter value..."
        )
        self.entry.pack(side="left", padx=(0, 5), fill="x", expand=True)

        # 3. BIND the function AFTER the entry is created
        self.entry.bind("<Control-v>", paste_func)

        # Individual delete button for precision control
        self.delete_btn = ctk.CTkButton(
            self, text="×", width=30, fg_color="transparent",
            text_color="gray", hover_color="#333333",
            command=lambda: delete_callback(self)
        )
        self.delete_btn.pack(side="right")

    def get_value(self):
        return self.entry.get().strip()

    def set_value(self, text):
        self.entry.delete(0, "end")
        self.entry.insert(0, text.strip())

class InputManager(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, label_text="Input List", **kwargs)
        self.rows = []

        # Add initial empty row
        self.add_row()

    def add_row(self, content=""):
        # We pass self.handle_paste as the 'paste_func' argument
        row = InputRow(
            self,
            delete_callback=self.remove_row,
            paste_func=self.handle_paste
        )
        row.pack(fill="x", pady=2, padx=5)
        row.set_value(content)
        self.rows.append(row)
        return row

    def remove_row(self, row_instance):
        if len(self.rows) > 1:  # Keep at least one row
            row_instance.destroy()
            self.rows.remove(row_instance)

    def handle_paste(self, event):
        # 1. Get clipboard data immediately
        try:
            clipboard = self.clipboard_get()
        except:
            return

        # 2. Split by any combination of newline, carriage return, or tab
        # This is critical for Excel columns (\r\n)
        entries = [e.strip() for e in re.split(r'[\r\n\t]+', clipboard) if e.strip()]

        if len(entries) > 1:
            # Find which row the user is currently in
            focused_entry = event.widget
            active_row = next((r for r in self.rows if r.entry == focused_entry), None)

            # If the current row is empty, use it for the first entry
            start_index = 0
            if active_row and not active_row.get_value():
                active_row.set_value(entries[0])
                start_index = 1

            # Create new rows for the remaining entries
            for i in range(start_index, len(entries)):
                self.add_row(entries[i])

            # Ensure the UI updates and scrolls
            self.update_idletasks()
            self._parent_canvas.yview_moveto(1.0)

            # CRITICAL: Return "break" to prevent the default single-line paste
            return "break"

        # If it's just one entry, let the default paste handle it
        return None

    def get_all_values(self):
        """Returns the final Python list for your automation engine."""
        return [row.get_value() for row in self.rows if row.get_value()]

class OutputRow(ctk.CTkFrame):
    def __init__(self, master, message, status="info", **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        # Color coding based on SAP result
        colors = {"success": "#2e7d32", "error": "#c62828", "info": "gray"}
        color = colors.get(status, "gray")

        self.dot = ctk.CTkLabel(self, text="●", text_color=color, width=20)
        self.dot.pack(side="left", padx=5)

        self.text_area = ctk.CTkTextbox(
            self,
            height=25,
            activate_scrollbars=False,
            fg_color="transparent",
            border_width=0
        )
        self.text_area.pack(side="left", fill="x", expand=True)

        # Insert the message
        self.text_area.insert("0.0", message)

        # LOCK IT: This allows highlighting/Ctrl+C but blocks typing/Backspaces
        self.text_area.configure(state="disabled")

class OutputManager(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, label_text="Automation Log", **kwargs)
        self.log_entries = []

    def log(self, message, status="info"):
        """Method called by the SAP engine to update the UI."""
        row = OutputRow(self, message, status)
        row.pack(fill="x", pady=1, padx=5)
        self.log_entries.append(row)

        # Auto-scroll to bottom so user sees latest update
        self._parent_canvas.yview_moveto(1.0)

    def remove_row(self, row_instance):
        """Deletes a specific log entry."""
        if row_instance in self.log_entries:
            row_instance.destroy()
            self.log_entries.remove(row_instance)

    def clear_all(self):
        """Wipes the entire log."""
        for row in self.log_entries:
            row.destroy()
        self.log_entries.clear()

    def copy_all_to_clipboard(self):
        """Extracts text from every OutputRow and moves it to the clipboard."""
        if not self.log_entries:
            return

        # 1. Collect text from every row
        # '0.0' to 'end-1c' gets all text minus the trailing newline character
        full_text = "\n".join(
            [row.text_area.get("0.0", "end-1c") for row in self.log_entries]
        )

        # 2. Update Windows Clipboard
        self.master.clipboard_clear()
        self.master.clipboard_append(full_text)

        # 3. Visual Feedback (Optional: temporarily change button text)
        old_text = self.copy_btn.cget("text")
        self.copy_btn.configure(text="Copied!", fg_color="#2e7d32")
        self.after(2000, lambda: self.copy_btn.configure(text=old_text, fg_color=("#3a7ebf", "#1f538d")))
