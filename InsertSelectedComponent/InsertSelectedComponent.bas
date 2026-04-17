Attribute VB_Name = "InsertSelectedComponent1"
Option Explicit

Sub Main()
    Dim swApp      As SldWorks.SldWorks
    Dim swModel    As SldWorks.ModelDoc2
    Dim swAssy     As SldWorks.AssemblyDoc
    Dim swSelMgr   As SldWorks.SelectionMgr
    Dim swComp     As SldWorks.Component2
    Dim newComp    As SldWorks.Component2
    Dim filePath   As String
    Dim selCount   As Long
    Dim boolStat   As Boolean
    Dim CopyCount  As Long
    Dim i          As Long
    Dim offsetX    As Double
    Dim posX      As Double
    Dim stepIndex  As Long
    
    ' Default X-axis offset in meters (100 mm)
    offsetX = 0.1
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    '-----------------------------------------------------------
    ' Validate active document
    '-----------------------------------------------------------
    If swModel Is Nothing Then
        swApp.SendMsgToUser2 "No active document.", swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
        Exit Sub
    End If
    
    If swModel.GetType <> swDocASSEMBLY Then
        swApp.SendMsgToUser2 "Active document is not an assembly.", swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
        Exit Sub
    End If
    
    Set swAssy = swModel
    Set swSelMgr = swModel.SelectionManager
    
    '-----------------------------------------------------------
    ' Check selection count
    '-----------------------------------------------------------
    selCount = swSelMgr.GetSelectedObjectCount2(-1)
    If selCount <> 1 Then
        swApp.SendMsgToUser2 "Select exactly one component in the graphics view or tree.", _
                             swMessageBoxIcon_e.swMbWarning, swMessageBoxBtn_e.swMbOk
        Exit Sub
    End If
    
    '-----------------------------------------------------------
    ' Get the component from graphics view or tree
    '-----------------------------------------------------------
    Set swComp = swSelMgr.GetSelectedObjectsComponent2(1)
    If swComp Is Nothing Then
        swApp.SendMsgToUser2 "Selected object is not a component.", swMessageBoxIcon_e.swMbWarning, swMessageBoxBtn_e.swMbOk
        Exit Sub
    End If
    
    filePath = swComp.GetPathName
    If filePath = "" Then
        swApp.SendMsgToUser2 "Could not get the selected component's file path.", swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
        Exit Sub
    End If
    
    '-----------------------------------------------------------
    ' Show UserForm to choose number of copies
    '-----------------------------------------------------------
    frmCopies.Show vbModal
    CopyCount = frmCopies.CopyCount
    
    ' User pressed Cancel or closed the form
    If CopyCount = 0 Then Exit Sub
    
    '-----------------------------------------------------------
    ' Insert copies with cumulative zig-zag X-axis placement
    '-----------------------------------------------------------
    For i = 0 To CopyCount - 1
        If i = 0 Then
            posX = 0
        Else
            stepIndex = Int((i + 1) / 2)
            If i Mod 2 = 1 Then
                posX = stepIndex * offsetX
            Else
                posX = -stepIndex * offsetX
            End If
        End If
        
        Set newComp = swAssy.AddComponent5( _
                        filePath, _
                        swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, _
                        "", _
                        False, _
                        "", _
                        posX, 0#, 0#)
                        
        If newComp Is Nothing Then
            swApp.SendMsgToUser2 "Failed to insert copy #" & (i + 1), swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
            Exit Sub
        End If
    Next i
    
    '-----------------------------------------------------------
    ' Rebuild and select the last inserted component
    '-----------------------------------------------------------
    swModel.EditRebuild3
    boolStat = newComp.Select4(True, Nothing, False)
    
    'swApp.SendMsgToUser2 copyCount & " copies of '" & swComp.Name2 & "' inserted in cumulative zig-zag X-axis. Last copy is selected.", _
                         swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk
End Sub


