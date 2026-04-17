VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCopies 
   Caption         =   "Insert Copies"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmCopies.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCopies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Title: InsertSelectedComponent
'
' Description:
'   Inserts the currently selected component (part or subassembly) into the
'   active SolidWorks assembly document. The macro assumes that a valid
'   component is selected prior to execution and automates the insertion
'   process using the SolidWorks API.
'
' Author:
'   Jonathan Mendoza
'
' Credits:
'   Developed with assistance from SolidWorks Macro Expert (Microsoft Copilot)
'
' References:
'   - SolidWorks API Help:
'     https://help.solidworks.com
'
'   - SldWorks.SldWorks Object Model
'   - SldWorks.AssemblyDoc Interface
'   - ModelDoc2::SelectionManager
'
' Notes:
'   - Requires an active Assembly document.
'   - A component must be selected before running the macro.
'   - Tested with SolidWorks 2025.
'
'-----------------------------------------------------------------------------
Option Explicit

Public CopyCount As Long  ' Stores number of copies or 0 if cancelled

Private Sub UserForm_Initialize()
    ' Initialize SpinButton
    spinCopies.Min = 1
    spinCopies.Max = 100
    spinCopies.Value = 1

    ' Always start TextBox at 1
    txtCopies.Text = "1"
End Sub

' When SpinButton changes, update TextBox only
Private Sub spinCopies_Change()
    txtCopies.Text = spinCopies.Value
End Sub

' When user types in TextBox, update SpinButton only
Private Sub txtCopies_Change()
    Dim val As Long
    If IsNumeric(txtCopies.Text) Then
        val = CLng(txtCopies.Text)
        If val < spinCopies.Min Then val = spinCopies.Min
        If val > spinCopies.Max Then val = spinCopies.Max
        spinCopies.Value = val
    End If
End Sub

' OK button: save value and close form
Private Sub btnOK_Click()
    CopyCount = CLng(txtCopies.Text)
    Me.Hide
End Sub

' Cancel button: set CopyCount = 0 and close
Private Sub btnCancel_Click()
    CopyCount = 0
    Me.Hide
End Sub

' Handle X button in upper right corner
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then  ' User clicked X
        CopyCount = 0
        Me.Hide
        Cancel = True       ' Prevent default unload
    End If
End Sub

