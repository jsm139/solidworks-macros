Attribute VB_Name = "SelectFloatingParts1"
' ============================================================
' Script Title : SelectFloatingPart
'
' Description  :
'   Analyzes the active SOLIDWORKS assembly at the TOP LEVEL ONLY.
'   Identifies under-constrained (floating) components using
'   IComponent2::GetConstrainedStatus and selects them.
'
'   - Evaluates only direct children of the root component
'   - Subassemblies are not traversed
'   - Uses SOLIDWORKS UserProgressBar for non-blocking UI feedback
'   - Displays dynamic progress text and a final summary message
'
' Author       : Jonathan Mendoza
' Date         : April 16, 2026
'
' Credits      :
'   SolidWorks Macro Expert (Copilot)
'   Macro structure, API usage guidance, and UI best practices
'
' References   :
'   1) SOLIDWORKS API Help – UserProgressBar (VBA)
'      Start, update, and stop progress indicator example
'      https://help.solidworks.com/2025/english/api/sldworksapi/
'      Start,_Update,_and_Stop_User_Progress_Indicator_Example_VB.htm
'
'   2) SOLIDWORKS API Help – IComponent2 Interface
'      GetConstrainedStatus(), GetChildren(), Name2
'      https://help.solidworks.com/2025/english/api/sldworksapi/
'      SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IComponent2.html
'
'   3) SOLIDWORKS API Help – Traversing Assembly Components (VBA)
'      Use of GetRootComponent3(True) and GetChildren
'      https://help.solidworks.com/2025/english/api/sldworksapi/
'      Get_Component_IDs_Example_VB.htm
'
'   4) CodeStack – SOLIDWORKS UserProgressBar Best Practices
'      Performance considerations and ESC-cancel behavior
'      https://www.codestack.net/solidworks-api/application/frame/user-progress-bar/
'
'   5) The CAD Coder – Assembly Component Traversal in VBA
'      Practical examples using Component2::GetChildren
'      https://thecadcoder.com/solidworks-vba-macros/assembly-traverse-sequentially/
'
' SOLIDWORKS   : 2025+
' Language     : VBA
' ============================================================
Option Explicit

Const DEV_MODE As Boolean = False

Sub main()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swConfMgr As SldWorks.ConfigurationManager
    Dim swConf As SldWorks.Configuration
    Dim swRootComp As SldWorks.Component2
    Dim swAssy As SldWorks.AssemblyDoc
    Dim vChildren As Variant
    Dim swChild As SldWorks.Component2
    Dim statusVal As Integer
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim selData As SldWorks.SelectData
    Dim i As Long
    Dim underCount As Long
    Dim fullyCount As Long
    Dim totalCount As Long

    ' Progress bar
    Dim pb As SldWorks.UserProgressBar
    Dim retVal As Boolean
    Dim pbResult As Long

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No active document open.", vbExclamation, "Error"
        Exit Sub
    End If

    If swModel.GetType <> swDocASSEMBLY Then
        MsgBox "This macro only works on assemblies.", vbExclamation, "Invalid Document"
        Exit Sub
    End If

    Set swAssy = swModel
    On Error GoTo Cleanup

    ' ------------------------------------------------------------
    ' Get root component and children
    ' ------------------------------------------------------------
    Set swConfMgr = swModel.ConfigurationManager
    Set swConf = swConfMgr.ActiveConfiguration
    Set swRootComp = swConf.GetRootComponent3(True)

    If swRootComp Is Nothing Then GoTo Cleanup

    vChildren = swRootComp.GetChildren
    If IsEmpty(vChildren) Then GoTo Cleanup

    totalCount = UBound(vChildren) + 1

    ' ------------------------------------------------------------
    ' Initialize progress bar
    ' ------------------------------------------------------------
    retVal = swApp.GetUserProgressBar(pb)
    pb.Start 0, 100, "SelectFloatingParts: analyzing components..."

    ' ------------------------------------------------------------
    ' First pass: count constraint statuses
    ' ------------------------------------------------------------
    For i = 0 To UBound(vChildren)

        Set swChild = vChildren(i)
        statusVal = swChild.GetConstrainedStatus()

        If statusVal = 2 Then
            underCount = underCount + 1
        ElseIf statusVal = 3 Then
            fullyCount = fullyCount + 1
        End If

        pb.UpdateTitle "Analyzing components (" & (i + 1) & _
                       " of " & totalCount & ")..."

        pbResult = pb.UpdateProgress((i + 1) * 50 \ totalCount)
        If pbResult = 2 Then GoTo Cleanup ' ESC pressed

    Next i

    If underCount = 0 Then
        MsgBox "Assembly is fully defined.", vbInformation, "SelectFloatingParts"
        GoTo Cleanup
    End If

    ' ------------------------------------------------------------
    ' Second pass: select under-constrained components
    ' ------------------------------------------------------------
    Set swSelMgr = swModel.SelectionManager
    Set selData = swSelMgr.CreateSelectData

    swModel.ClearSelection2 True
    underCount = 0

    For i = 0 To UBound(vChildren)

        Set swChild = vChildren(i)

        If swChild.GetConstrainedStatus() = 2 Then
            swChild.Select4 True, selData, False
            underCount = underCount + 1
        End If

        pb.UpdateTitle "Selecting components (" & (i + 1) & _
                       " of " & totalCount & ")..."

        pbResult = pb.UpdateProgress(50 + (i + 1) * 50 \ totalCount)
        If pbResult = 2 Then GoTo Cleanup ' ESC pressed

    Next i

    ' ------------------------------------------------------------
    ' Final summary
    ' ------------------------------------------------------------
    MsgBox underCount & " under-constrained components selected.", _
           vbInformation, "SelectFloatingParts"

Cleanup:
    ' ------------------------------------------------------------
    ' Always stop progress bar
    ' ------------------------------------------------------------
    If Not pb Is Nothing Then
        pb.End
    End If

End Sub

