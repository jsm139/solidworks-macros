# SelectFloatingPart — SolidWorks Macro

## Purpose
Automatically identifies and selects under‑constrained (floating) **top‑level**
components in a SolidWorks assembly to improve assembly validation and mate
completeness checks.

## Problem
In complex or fastener‑heavy assemblies, unintentionally floating components can
be difficult to identify. While SolidWorks allows users to drag components to
check for movement, it does not provide a native way to programmatically identify
all under‑constrained components at once.

Manual inspection becomes increasingly inefficient as assembly size grows,
especially when fixed, grounded, and Toolbox components coexist alongside
components that are unintentionally free to move.

## Solution
This macro evaluates the constraint status of **top‑level components only**
(direct children of the root assembly) and automatically selects those that are
under‑constrained.

By intentionally ignoring fully constrained, fixed, grounded, and Toolbox
components, the macro focuses exclusively on components that require engineering
attention.

The macro:
- Identifies under‑constrained top‑level components
- Selects floating components in the graphics area
- Ignores fixed, grounded, and Toolbox components
- Avoids traversing into subassemblies
- Provides non‑blocking progress feedback
- Supports ESC key cancellation
- Executes efficiently, even in large assemblies

## Demo
![Select Floating Part Demo](SelectFloatingPart.gif)

## How It Works (High‑Level)
1. Macro verifies that an assembly document is active
2. Retrieves the root component of the active configuration
3. Iterates through top‑level child components only
4. Evaluates each component using `GetConstrainedStatus`
5. Selects components that are under‑constrained
6. Displays progress and a final selection summary

## Why This Matters
- Quickly exposes unintentionally floating components
- Eliminates repetitive manual drag‑testing
- Improves assembly validation consistency
- Scales effectively for large assemblies
- Provides deterministic and repeatable results

## Files
- `SelectFloatingPart.swp` — Executable SolidWorks macro
- `SelectFloatingPart.bas` — Readable source code

*(Development and testing performed using non‑proprietary assemblies.)*
