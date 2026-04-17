# InsertSelectedComponent — SolidWorks Macro

## Purpose

Inserts one or more copies of a selected component (part or subassembly) into the
active SolidWorks assembly using a simple, repeatable workflow driven by a
lightweight user input dialog.

---

## Problem

In assembly workflows, inserting multiple instances of the same component often
requires repeating the Insert Component command, navigating dialogs, and manually
confirming placement each time.

When working with patterned hardware, repeated subassemblies, or copied layout
components, this process becomes unnecessarily repetitive and interrupts design
flow.

---

## Solution

This macro streamlines the insertion process by allowing the user to **select a
component first**, then specify the desired number of copies using a simple input
form.

The macro uses the active selection to determine what component to insert and
handles multiple insertions automatically, eliminating repeated manual steps.

The macro:

- Requires a selected component (part or subassembly)
- Prompts the user for the number of copies to insert
- Inserts all requested instances into the active assembly
- Gracefully handles cancellation
- Uses a clean, non‑blocking user input form

---

## Demo

![Insert Selected Component Demo](InsertSelectedComponent.gif)

---

## How It Works (High‑Level)

1. Verifies that the active document is an assembly
2. Reads the currently selected component
3. Displays a form prompting for the number of copies to insert
4. If cancelled, exits safely without modifying the assembly
5. Inserts the selected component the specified number of times
6. Leaves placement and mating to the user as needed

---

## User Input Form (`frmCopies`)

The macro uses a simple user form to capture copy count input:

- SpinButton with defined minimum and maximum values
- TextBox synchronized with the SpinButton
- OK button to confirm insertion
- Cancel button and window close handling to exit safely

This ensures valid input while keeping user interaction minimal.

---

## Why This Matters

- Speeds up repetitive component insertion tasks
- Reduces UI friction during assembly creation
- Improves focus when working with repeated components
- Provides a predictable, repeatable insertion workflow
- Demonstrates clean macro‑UI integration using VBA forms

---

## Compatibility

- SolidWorks assemblies only
- Requires an actively selected component
- Tested with SolidWorks 2025
- Written in VBA

---

## Files

- `InsertSelectedComponent.swp` — Executable SolidWorks macro  
- `InsertSelectedComponent.bas` — Readable source code  
- `frmCopies.frm` — User input form for copy count  
- `InsertSelectedComponent.gif` — Visual demonstration

---

*Demonstration performed using non‑proprietary sample assemblies.*
