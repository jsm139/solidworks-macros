# SelectFloatingPart

## Overview

**SelectFloatingPart** is a SOLIDWORKS VBA macro that analyzes the **top-level components** of the active assembly and automatically **selects all under-constrained (floating) components**.

The macro is designed to help engineers quickly identify components that are free to move at the top level of an assembly, without traversing into subassemblies or relying on manual inspection.

It provides **non-blocking UI feedback** using SOLIDWORKS’ built-in progress indicator and displays a final summary of the results.

---

## What the Macro Does

- Verifies the active document is an **assembly**
- Evaluates **top-level components only**
- Uses `IComponent2::GetConstrainedStatus`
- Treats **only under-constrained components** as actionable
- Selects all floating components in the graphics area
- Displays progress in the SOLIDWORKS status bar
- Shows a final summary message with the count of selected components
- Supports **ESC key cancellation**

---

## What the Macro Does *Not* Do

- Does **not** traverse into subassemblies
- Does **not** modify mates or fix components
- Does **not** treat fixed, grounded, or toolbox components as errors
- Does **not** rely on UserForms or modal dialogs

---

## Constraint Status Logic (Important)

This macro intentionally treats **only one constraint state as actionable**:

| Status Value | Meaning (Assembly Context) | Behavior |
|-------------|----------------------------|----------|
| `2` | Under-constrained (floating) | Selected |
| Any other value (`3`, `4`, etc.) | Fully constrained, fixed, grounded, toolbox, or equivalent | Ignored |

**Why this matters**  
In assemblies, `GetConstrainedStatus()` can return values beyond just “under-” and “fully-constrained”.  
Toolbox fasteners and fixed components commonly return values such as `4`, which are valid and not errors.

---

## User Interface Behavior

- Uses **SOLIDWORKS UserProgressBar** (official API)
- Displays messages such as:
  - `Analyzing components (23 of 187)...`
  - `Selecting components (94 of 187)...`
- Allows the user to press **ESC** to cancel execution
- Cleans up the progress indicator on exit or cancellation

---

## Requirements

- SOLIDWORKS **2025 or newer**
- Assembly document must be active
- VBA macros enabled

---

## How to Use

1. Open a SOLIDWORKS **assembly**
2. Run the `SelectFloatingPart` macro
3. Wait for the progress indicator to complete
4. Review the selected floating components
5. Use the final summary message as confirmation

---

## Output

- Floating components are **selected in the FeatureManager and graphics area**
- A message box displays:
XX under-constrained components selected.

---

## References / Technical Basis

This macro is based on official SOLIDWORKS API documentation and established best practices:

1. **User Progress Indicator (VBA)**  
   SOLIDWORKS API Help – Start, Update, and Stop User Progress Indicator  
   https://help.solidworks.com/2025/english/api/sldworksapi/Start,_Update,_and_Stop_User_Progress_Indicator_Example_VB.htm

2. **IComponent2 Interface**  
   SOLIDWORKS API Help – `GetConstrainedStatus`, `GetChildren`, `Name2`  
   https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IComponent2.html

3. **Assembly Traversal Pattern**  
   SOLIDWORKS API Help – Root component and child traversal  
   https://help.solidworks.com/2025/english/api/sldworksapi/Get_Component_IDs_Example_VB.htm

4. **UserProgressBar Best Practices**  
   CodeStack – Handling long operations and ESC cancel behavior  
   https://www.codestack.net/solidworks-api/application/frame/user-progress-bar/

5. **Component Traversal in VBA**  
   The CAD Coder – Practical use of `Component2::GetChildren`  
   https://thecadcoder.com/solidworks-vba-macros/assembly-traverse-sequentially/

---

## Author & Credits

- **Author:** Jonathan Mendoza  
- **Date:** April 16, 2026  
- **Credits:** SolidWorks Macro Expert (Copilot)  
  - Macro structure
  - SOLIDWORKS API guidance
  - Progress indicator implementation
  - Constraint-status interpretation

---

## Revision History

- **REV A** – Initial release (Apr 16, 2026)

---

## Future Enhancements (Optional)

- Ignore Toolbox components
- Automatically fix selected floating components
- Export floating component list to CSV
- Add configuration-specific filtering
- Traverse subassemblies (optional mode)

---

## License / Usage

Internal engineering automation tool.  
Modify and reuse freely within your organization.
