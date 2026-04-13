# CenterDrawingView — SolidWorks Macro

## Purpose
Automatically centers a selected drawing view on the active drawing sheet
by aligning the view’s X and Y position to the sheet center.

## Problem
When working with long or narrow components (such as ball screws), drawing
views are often repositioned multiple times to achieve a clean, visually
balanced layout. Manually centering views is repetitive and can interrupt
the dimensioning workflow—especially when precision and symmetry are important.

## Solution
This macro instantly centers a selected drawing view on the active sheet,
ensuring consistent placement and improving visual clarity for downstream
dimensioning tasks.

The macro:
- Centers the view in both **X and Y directions**
- Works on the **active drawing sheet**
- Requires only a single view selection
- Executes instantly within SolidWorks

## Demo
CenterDrawingView.gif

## How It Works (High-Level)
1. User selects a drawing view
2. Macro determines the active sheet’s center location
3. Macro updates the view’s position so it is centered in X and Y
4. Dimensions can then be added cleanly and consistently

## Why This Matters
- Improves drawing readability and visual balance
- Speeds up layout preparation before dimensioning
- Especially useful for **long components** (e.g., ball screws, shafts, rails)
- Eliminates repetitive manual repositioning

## Files
- `CenterDrawingView.swp` — Executable SolidWorks macro
- `CenterDrawingView.bas` — Readable source code
- `CenterDrawingView.gif` — Visual demonstration

*(Demonstration uses SOLIDWORKS sample drawings.)*
``
