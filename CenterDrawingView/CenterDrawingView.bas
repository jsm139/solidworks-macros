Attribute VB_Name = "CenterDrawingView1"

'------------------------------------------------------------
' Macro Name: CenterDrawingView
' Description:
'   This macro centers the currently selected drawing view
'   on the active sheet in a SolidWorks drawing document.
'
' Usage:
'   1. Open a drawing in SolidWorks.
'   2. Select a drawing view.
'   3. Run this macro to move the view to the sheet center.
'
' Credits:
'   Author: Jonathan Mendoza
'   Assistant: SolidWorks Macro Expert (AI)
'
' References:
'   - SldWorks.Sheet.GetProperties2:
'       https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISheet~GetProperties2.html
'   - SldWorks.View.Position:
'       https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IView~Position.html
'   - SldWorks.SelectionMgr.GetSelectedObject6:
'       https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISelectionMgr~GetSelectedObject6.html
'------------------------------------------------------------

Option Explicit

Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim swSheet As SldWorks.Sheet
    Dim swView As SldWorks.View
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim vProps As Variant
    Dim width_m As Double, height_m As Double
    Dim centerX As Double, centerY As Double
    Dim viewPos As Variant
    Dim selType As Long

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Check if there is an active document
    If swModel Is Nothing Then
        MsgBox "Please open a drawing document.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the active document is a drawing
    If swModel.GetType <> swDocDRAWING Then
        MsgBox "This macro only works on drawing files.", vbExclamation
        Exit Sub
    End If

    Set swDraw = swModel
    Set swSheet = swDraw.GetCurrentSheet
    Set swSelMgr = swModel.SelectionManager

    If swSelMgr.GetSelectedObjectCount2(-1) <> 1 Then
        MsgBox "Please select exactly one drawing view.", vbExclamation
        Exit Sub
    End If

    ' Validate selection type
    selType = swSelMgr.GetSelectedObjectType3(1, -1)
    If selType <> swSelDRAWINGVIEWS Then
        MsgBox "Please select a drawing view before running the macro.", vbExclamation
        Exit Sub
    End If

    ' Safe to get the selected view
    Set swView = swSelMgr.GetSelectedObject6(1, -1)

    ' Get sheet size
    vProps = swSheet.GetProperties2
    width_m = vProps(5)
    height_m = vProps(6)

    centerX = width_m / 2#
    centerY = height_m / 2#

    ' Move view to sheet center
    viewPos = swView.Position
    viewPos(0) = centerX
    viewPos(1) = centerY
    swView.Position = viewPos

    swModel.GraphicsRedraw2
End Sub


