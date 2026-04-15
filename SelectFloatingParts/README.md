# SelectFloatingParts — SolidWorks Macro

## Purpose

Identifies and selects under‑defined (floating) top‑level components in a
SolidWorks assembly to help diagnose constraint issues quickly and visually.

---

## Problem

When working with assemblies, SolidWorks may indicate that an assembly is
under‑defined, even when it appears to be fully constrained—or, conversely,
may indicate that a large assembly is fully defined when some components are
still floating.

Manually locating under‑defined components in these situations can be
time‑consuming, especially in large assemblies where floating parts are not
immediately obvious in the graphics area or FeatureManager tree.

---

## Solution

This macro traverses the FeatureManager tree at the **top level** of the active
assembly, evaluates each component’s constraint status, and automatically
selects all under‑defined components.

Once selected, the user can immediately isolate the floating parts in the
graphics view to inspect, constrain, or troubleshoot them.

The macro:

- Evaluates top‑level components only
- Includes suppressed components in the analysis
- Identifies under‑constrained components using SolidWorks API status codes
- Selects all floating parts in a single operation
- Provides clear feedback if the assembly is fully defined or in an invalid state

---

## Scope

- Only **top‑level components** (direct children of the root component) are
  evaluated
- Subassemblies are not traversed internally
- Intended for diagnostic use during assembly constraint validation

---

## Demo

SelectFloatingParts.gif

---

## How It Works (High‑Level)

1. Verifies that the active document is an assembly
2. Retrieves the active configuration and root component
3. Collects all top‑level child components
4. Evaluates each component’s constrained status
5. If all components are fully constrained, reports success and exits
6. Otherwise, selects all under‑defined components
7. User can then isolate selected components for clear visualization

---

## Why This Matters

- Quickly identifies floating components in complex assemblies
- Eliminates guesswork when assemblies report inconsistent constraint states
- Improves confidence in assembly definition before release
- Speeds up troubleshooting and constraint cleanup
- Complements isolation and visual inspection workflows

---

## Compatibility

- SolidWorks assemblies only
- Active configuration only
- Suppressed components included
- Uses `GetConstrainedStatus` API enumeration

---

## Files

- `SelectFloatingParts.swp` — Executable SolidWorks macro  
- `SelectFloatingParts.bas` — Readable source code  
- `SelectFloatingParts.gif` — Visual demonstration

---

*Demonstration uses non‑proprietary sample assemblies.*
