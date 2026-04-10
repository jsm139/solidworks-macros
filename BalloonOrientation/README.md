# BalloonOrientation — SolidWorks Macro

## Purpose
Automatically orient drawing balloons **vertically or horizontally** based on the
leader’s X/Y direction to improve drawing consistency and readability.

## Problem
In large or frequently edited drawings, balloons often become misaligned.
Manually rotating multiple balloons is repetitive and time-consuming.

## Solution
This macro:
- Uses the balloon leader’s X/Y location
- Applies consistent orientation logic
- Supports **single or multiple balloon selections**
- Executes instantly within SolidWorks drawings

## Demo
demo.gif

## How It Works (High-Level)
1. User selects one or more drawing balloons
2. Macro prompts for desired orientation
3. For each balloon:
   - Evaluates leader position
   - Applies orientation logic
   - Preserves leader attachment

## Files
- `BalloonOrientation.swp` — executable macro
- `BalloonOrientation.bas` — readable source code

## Usage
1. Open a SolidWorks drawing
2. Select one or more balloons
3. Run `BalloonOrientation.swp`
4. Choose orientation when prompted
