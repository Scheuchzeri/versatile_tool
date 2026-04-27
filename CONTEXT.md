# Versatile Shop — Claude Code Context

## Agent Role

Claude acts as a **consulting agent** on this project. Responsibilities:
- Research SAP GUI scripting patterns and relevant Python techniques
- Propose and discuss solution approaches
- Ask critical questions to guide design decisions
- Log progress and decisions here in CONTEXT.md

Claude does **not** write or modify source code directly. All implementation is done by the developer.

---

## Project Overview

This repository implements a Windows-based Python automation tool for SAP ILP and ISP environments. It is built around a GUI-driven workflow where the user provides ERP or person identifiers, and the program scripts SAP GUI actions to execute SAP transactions automatically.

---

## Key Files

**main.py**
Entry point and task dispatcher. Defines SAP automation actions such as `export_relationships`, `open_erp`, and `list_person_ids`. Uses `pythoncom` and thread-based workers via `src.lib.start_worker`.

**src/gui.py**
CustomTkinter GUI implementation. Provides input rows, output logging, task selection, and an execute button. Collects user data and calls backend runner functions.

**src/engine.py**
SAP automation engine using `win32com.client`. Manages SAP session lifecycle via `Session`. Implements navigation, field entry, file export, and basic SAP GUI operations.

**src/lib.py**
Contains SAP transaction mappings (`TRANS_ACTIONS`) and UI element IDs. Loads `config.ini`, spawns worker threads, and provides helper utilities. `SAP_MAP` is being replaced by `ACTION_MAP` (see below).

**data/config.ini**
Stores SAP connection names and paths for ILP/ISP. Expected SAP GUI path is `C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe`.

---

## Behavior

1. User selects ILP or ISP mode in the GUI.
2. User chooses a task and enters one or more values.
3. `main.py` launches a background worker that connects to SAP, navigates to the transaction, and inputs values.
4. Results and log messages are displayed in the GUI output area.

---

## Environment

- Windows only
- Requires SAP GUI installed with scripting support
- Requires Python packages: `customtkinter`, `pywin32` / `pythoncom`, `psutil`
- The app expects `saplogon.exe` available at the configured path

---

## Research: GuiShell Toolbar API

### The Two-Tier Object Model

The SAP GUI Scripting object model has two distinct tiers:

**Tier 1 — The standard window tree** (what `return_view_structure` currently crawls)
```
GuiSession → GuiFrameWindow → GuiContainer (wnd[0]/usr) → ... → GuiShell
```
Everything here is traversable via `.Children` and addressable via `findById()`. `_return_sub_elements` operates entirely within this tier.

**Tier 2 — Inside the GuiShell itself**
Once a `GuiShell` element is reached (subtype `GridView`, `ToolbarControl`, etc.), the object model changes. Toolbar buttons inside the shell are **not** exposed as child nodes in the standard tree. They are only reachable via dedicated methods:
- `shell.pressToolbarButton(buttonId)` — presses a toolbar button by its string ID
- `shell.pressToolbarContextButton(buttonId)` — presses a context menu toolbar button

### Key Constraints

- There is **no `.Toolbars` collection**, no `.GetToolbarButtonCount()`, and no enumeration API on GuiShell.
- The toolbar button interface is **write-only from the scripting perspective**: you can press buttons but cannot ask the shell what buttons exist.
- Pressing `F1` on a toolbar button in the live SAP GUI does **not** open the technical information dialog — confirming that these elements are outside the standard element inspection mechanism.
- Button ID strings such as `&MB_EXPORT` are **stable** across different screen renderings. The volatility is in the **path to the shell**, not in the button IDs themselves.

### Implication for Dynamic Discovery

The volatile ID problem therefore splits into two separable sub-problems:

1. **Finding the shell dynamically** — The current `_return_sub_elements` recursion skips `GuiShell` because it is not in `interactive_types`. This is solvable by extending the crawl to detect and return `GuiShell` elements by `child.Type == "GuiShell"`.

2. **Calling toolbar buttons on the found shell** — Once the shell object is located, `pressToolbarButton(buttonId)` can be called using stable, known button ID strings. No enumeration is needed.

This means the solution architecture is:
- **Dynamic:** locate the shell at runtime via tree crawl
- **Static:** button IDs remain as known constants (e.g. `&MB_EXPORT`, `&PC`)

---

## ZLSO_VAP1 View Structure (Observed)

Running `return_view_structure_extended` on the ZLSO_VAP1 transaction confirmed **two GuiShell objects** are present simultaneously, both of subtype `GridView`, both named `shell`:

| | Shell 0 | Shell 1 |
|---|---|---|
| Location | Left panel (search results) | Right panel (relationships table) |
| Stable anchor | `subSCREEN_1010_LEFT_AREA` | `subSCREEN_1010_RIGHT_AREA` |
| Container suffix | `cntlSCREEN_1080_CONTAINER/shellcont/shell` | `cntlSCREEN_1220_CUSTOM_CONTROL/shellcont/shell` |
| Target for export | No | **Yes** |

`Name` and `SubType` are identical for both shells — the only reliable discriminator is a substring of the ID path. `RIGHT_AREA` has been chosen as the match token for the relationships table shell, on the basis that the table always renders on the right side of the screen. This assumption will be validated in practice.

### Export Format Dialog (wnd[1])

After selecting "Export as local file", SAP opens a standard modal dialog (`wnd[1]`). This window is invisible to `return_view_structure_extended` when anchored at `wnd[0]/usr`. The dialog uses SAP standard function pool `SAPLSPO5` and renders radio buttons as a step-loop with the field name `SPOPLI-SELFLAG`, indexed `[0,0]` through `[4,0]`.

Observed structure:
```
/app/con[0]/ses[0]/wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]
/app/con[0]/ses[0]/wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]
... (through [4,0])
/app/con[0]/ses[0]/wnd[1]/usr/subSUBSCREEN_CONTROL:SAPLSPO5:0170/cntlHTML_CONTROL_CON/shellcont/shell  ← preview pane, not interactive
```

Because `SAPLSPO5` is a SAP standard pool, this same dialog structure will appear in other transactions. The solution is designed for reuse.

---

## Design: ACTION_MAP and Dynamic Lookup

### Unified Lookup Dictionary (current state)

A single `ACTION_MAP` in `lib.py` replaces `SAP_MAP`. Each entry declares a `type` field so the engine routes the call correctly. A `wnd_idx` field specifies the target window, defaulting to `0` when absent.

```python
ACTION_MAP = {
    # --- Regular elements (resolved via findById) ---
    "ERP_INPUT": {
        "type": "element",
        "match": "ctxtRF02D-KUNNR",
    },
    "MAX_RECORDS": {
        "type": "element",
        "match": "txtBURS_JOEL_SELECTION-MAX_RECORD",
    },

    # --- Shell toolbar buttons (resolved via shell lookup + pressToolbarButton) ---
    "EXPORT_RELATIONS": {
        "type": "shell_btn",
        "shell_match": "RIGHT_AREA",
        "btn_id": "&MB_EXPORT",
        "sub_btn_id": "&PC",            # chained context menu action, called in same step
    },
    "EXPORT_LOCAL_FILE": {
        "type": "shell_btn",
        "shell_match": "RIGHT_AREA",
        "btn_id": "&PC",
    },

    # --- Radio buttons (resolved via label match in target window) ---
    "EXPORT_UNCONVERTED": {
        "type": "radio_btn",
        "btn_label": "Unconverted",
        "wnd_idx": 1,
    },
}
```

**Known design tension:** `EXPORT_RELATIONS` encodes a two-step action (`pressToolbarContextButton` followed by `selectContextMenuItem`) as a single entry using `sub_btn_id`. This was necessary because the shell did not recognize the commands when called in separate steps. This may cause ambiguity as more entries are added — flagged for future review.

### Engine Methods

**`return_view_structure_extended(wnd_idx=0)`**
Crawls `wnd[{wnd_idx}]/usr` and returns `{"elements": [...id strings...], "shells": [...id strings...]}`. Shell objects returned as ID strings, not raw COM references.

**`_return_sub_elements_extended(container)`**
Recursive helper. Collects standard interactive element IDs and `GuiShell` IDs separately. `GuiShell` detected by `child.Type == "GuiShell"`.

**`resolve_action(key)`**
Consults `ACTION_MAP`, crawls the live view, and returns a ready-to-use `findById` object for `element` type, or finds and acts on the shell directly for `shell_btn` type. Raises `ValueError` with a descriptive message on unknown key, no match, or ambiguous match.

**`select_radio_option(key)`**
Resolves a `radio_btn` entry from `ACTION_MAP`. Crawls `wnd[{wnd_idx}]/usr`, iterates all `GuiRadioButton` elements, matches on `.Text.strip()` against `btn_label`, and calls `.select()`. Raises `ValueError` if label not found.

```python
def select_radio_option(self, key):
    entry = lib.ACTION_MAP.get(key)
    if not entry or entry["type"] != "radio_btn":
        raise ValueError(f"Key '{key}' is not a valid radio_btn entry.")

    wnd_idx = entry.get("wnd_idx", 0)
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
```

---

## Known Issues / Flagged for Review

- **`EXPORT_RELATIONS` chained action:** `sub_btn_id` encodes a two-step context menu interaction in a single `ACTION_MAP` entry. Works currently but may cause design conflicts as more entries are added.
- **`_identify_target` unbound variable:** The original method leaves `target` unbound if no element matches the key, causing `NameError` instead of a clean failure. Now superseded by `resolve_action` but should be removed or fixed to avoid confusion.

---

## Open Problems / Next Session

- **ZLSO_VAP1 "Relationships" tab navigation:** `sendVkey(21)` (Shift+F1) is rejected by SAP as "virtual key not enabled" in the ZLSO_VAP1 main view. Recorder output confirmed the correct invocation is `session.findById("wnd[0]/tbar[1]/btn[13]").press()` — a standard `tbar[1]` button, no shell involved. Next step: add an `ACTION_MAP` entry for this button and confirm whether `resolve_action` returning the raw element is sufficient to call `.press()` on it.

---

## Decision Log

- **2026-04-22:** Chose `RIGHT_AREA` as the shell discriminator substring for the ZLSO_VAP1 relationships table. Rationale: table always renders on the right panel; `subSCREEN_1010_RIGHT_AREA` is structurally stable. To be validated in practice.
- **2026-04-22:** Shell elements returned as ID strings from `return_view_structure_extended`, not as raw COM objects. Rationale: keeps return value as pure data; `findById` called fresh at action time.
- **2026-04-22:** Adopted unified `ACTION_MAP` with explicit `type` field to replace `SAP_MAP`. Rationale: single lookup point for both regular elements and shell toolbar buttons.
- **2026-04-22:** Added `wnd_idx` field to `ACTION_MAP` entries. Rationale: allows targeting modal windows (`wnd[1]`, etc.) without duplicating crawl logic. Defaults to `0` when absent.
- **2026-04-22:** `select_radio_option` matches by label string rather than index. Rationale: self-documenting, robust to reordering of options in the dialog.
- **2026-04-23:** Resolved threading race in `check_for_participants`. Lookup logic moved into `_export_rels` as a continuation after `curr.save(1)`, running on the same worker thread. `erp` and `emails` bundled as a dict and passed as the `data` argument to `start_worker`.
- **2026-04-23:** Extracted `_compare_against_file(app, emails)` as a standalone helper. Performs case-insensitive substring match (`addr.lower() in pt.lower()`) against lines in `relations_raw.txt`. Logs participants not yet created to the OutputManager.
- **2026-04-23:** Removed `_export_relationships` wrapper — was dead code after `check_for_participants` was refactored to call `start_worker` directly.
- **2026-04-23:** Collapsed `handle_start_ilp` / `handle_start_isp` into a single `handle_start(config, system)` function parameterised by config key.
