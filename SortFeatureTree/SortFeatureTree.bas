Attribute VB_Name = "SortFeatureTree1"
' ============================================================================
' Macro Title : SortFeatureTree
'
' Description :
'   Reorders top-level components in a SolidWorks assembly FeatureManager tree
'   into a consistent, human-readable order.
'
'   The macro performs the following operations:
'     1) Retrieves all top-level components (including suppressed components)
'     2) Groups components into:
'        - Numeric-starting names (e.g. "24009-3-1")
'        - Letter-starting names (e.g. "M5-0.8X12 (27B)-6")
'     3) Sorts components within each group using a unified comparison logic:
'        - Primary key: Base component name (instance suffix removed)
'        - Secondary key: Numeric instance number (e.g. -1, -2, -10)
'     4) Reorders the FeatureManager tree safely using a temporary folder
'
' Sorting Logic :
'   - Components are first grouped by whether their name begins with a digit
'     or a letter (numeric-first, letter-second).
'   - Within each group, components are sorted by:
'       a) Base name (case-insensitive string compare)
'       b) Instance number (numeric comparison, not string-based)
'
'   Example:
'     24009-3-1
'     24009-3-2
'     ...
'     24009-3-9
'     24009-3-10
'
' Compatibility :
'   - SolidWorks Assemblies only
'   - All configurations supported
'   - Suppressed components included
'
' Authors :
'   Jonathan Mendoza
'
' Credits :
'   SolidWorks Macro Expert (ChatGPT)
'
' References :
'   - SolidWorks API Help:
'       AssemblyDoc.ReorderComponents
'       FeatureManager.InsertFeatureTreeFolder2
'       Component2.Name2
'   - Common SolidWorks macro best practices for FeatureManager reordering
'
' ============================================================================
Option Explicit

' ============================================================
' Helper: Split SolidWorks component name into base + instance
' Example: "24009-3-10" -> baseName = "24009-3", instanceNum = 10
' ============================================================
Private Sub SplitNameInstance(ByVal fullName As String, _
                              ByRef baseName As String, _
                              ByRef instanceNum As Long)

    Dim pos As Long
    pos = InStrRev(fullName, "-")

    If pos > 0 Then
        baseName = Left$(fullName, pos - 1)

        On Error Resume Next
        instanceNum = CLng(Mid$(fullName, pos + 1))
        If Err.Number <> 0 Then instanceNum = -1
        On Error GoTo 0
    Else
        baseName = fullName
        instanceNum = -1
    End If

End Sub

' ============================================================
' Helper: Case-insensitive string comparison
' ============================================================
Private Function StrCompare(a As String, b As String) As Long
    a = LCase$(a)
    b = LCase$(b)

    If a < b Then
        StrCompare = -1
    ElseIf a > b Then
        StrCompare = 1
    Else
        StrCompare = 0
    End If
End Function

' ============================================================
' Unified component comparison function
'
' Returns:
'   -1 if compA < compB
'    0 if equal
'   +1 if compA > compB
' ============================================================
Private Function CompareComponents(compA As SldWorks.Component2, _
                                   compB As SldWorks.Component2) As Long

    Dim baseA As String, baseB As String
    Dim instA As Long, instB As Long
    Dim cmp As Long

    SplitNameInstance compA.Name2, baseA, instA
    SplitNameInstance compB.Name2, baseB, instB

    ' 1) Compare base names
    cmp = StrCompare(baseA, baseB)
    If cmp <> 0 Then
        CompareComponents = cmp
        Exit Function
    End If

    ' 2) Same base name -> compare numeric instance
    If instA >= 0 And instB >= 0 Then
        If instA < instB Then
            CompareComponents = -1
        ElseIf instA > instB Then
            CompareComponents = 1
        Else
            CompareComponents = 0
        End If
    Else
        CompareComponents = 0
    End If

End Function

' ============================================================
' Main Macro
' ============================================================
Sub Main()

    Dim swApp As SldWorks.SldWorks
    Dim swDoc As SldWorks.ModelDoc2
    Dim swAsm As SldWorks.AssemblyDoc
    Dim swConf As SldWorks.Configuration
    Dim swRootComp As SldWorks.Component2
    Dim vChildren As Variant

    Dim i As Long, j As Long
    Dim tmpComp As SldWorks.Component2

    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc

    If swDoc Is Nothing Then
        MsgBox "No SolidWorks document open."
        Exit Sub
    End If

    If swDoc.GetType <> swDocASSEMBLY Then
        MsgBox "Active document is not an assembly."
        Exit Sub
    End If

    Set swAsm = swDoc
    Set swConf = swDoc.GetActiveConfiguration
    Set swRootComp = swConf.GetRootComponent3(True)

    vChildren = swRootComp.GetChildren
    If IsEmpty(vChildren) Then
        MsgBox "No top-level components found."
        Exit Sub
    End If

    ' ------------------------------------------------------------
    ' Separate components into numeric-starting and letter-starting
    ' ------------------------------------------------------------
    Dim numList() As SldWorks.Component2
    Dim letList() As SldWorks.Component2
    Dim nCount As Long: nCount = 0
    Dim lCount As Long: lCount = 0

    For i = 0 To UBound(vChildren)
        Dim name As String
        name = vChildren(i).Name2

        If Len(name) > 0 And Mid$(name, 1, 1) Like "[0-9]" Then
            ReDim Preserve numList(0 To nCount)
            Set numList(nCount) = vChildren(i)
            nCount = nCount + 1
        Else
            ReDim Preserve letList(0 To lCount)
            Set letList(lCount) = vChildren(i)
            lCount = lCount + 1
        End If
    Next i

    ' ------------------------------------------------------------
    ' Unified sorting using CompareComponents
    ' ------------------------------------------------------------
    If nCount > 1 Then
        For i = 0 To nCount - 2
            For j = i + 1 To nCount - 1
                If CompareComponents(numList(i), numList(j)) > 0 Then
                    Set tmpComp = numList(i)
                    Set numList(i) = numList(j)
                    Set numList(j) = tmpComp
                End If
            Next j
        Next i
    End If

    If lCount > 1 Then
        For i = 0 To lCount - 2
            For j = i + 1 To lCount - 1
                If CompareComponents(letList(i), letList(j)) > 0 Then
                    Set tmpComp = letList(i)
                    Set letList(i) = letList(j)
                    Set letList(j) = tmpComp
                End If
            Next j
        Next i
    End If

    ' ------------------------------------------------------------
    ' Combine numeric + letter lists
    ' ------------------------------------------------------------
    Dim total As Long
    total = nCount + lCount

    Dim combined() As SldWorks.Component2
    ReDim combined(0 To total - 1)

    Dim idx As Long: idx = 0

    For i = 0 To nCount - 1
        Set combined(idx) = numList(i)
        idx = idx + 1
    Next i

    For i = 0 To lCount - 1
        Set combined(idx) = letList(i)
        idx = idx + 1
    Next i

    ' ------------------------------------------------------------
    ' Reorder using FeatureManager folder (safe method)
    ' ------------------------------------------------------------
    On Error GoTo CleanupAndFail
    swDoc.FeatureManager.EnableFeatureTree = False

    combined(0).Select4 False, Nothing, False
    Dim swFeatFolder As SldWorks.Feature
    Set swFeatFolder = swDoc.FeatureManager.InsertFeatureTreeFolder2( _
        swFeatureTreeFolderType_e.swFeatureTreeFolder_Containing)

    For i = 0 To total - 1
        swAsm.ReorderComponents combined(i), swFeatFolder, _
            swReorderComponentsWhere_e.swReorderComponents_LastInFolder
    Next i

    swFeatFolder.Select2 False, -1
    swDoc.Extension.DeleteSelection2 swDeleteSelectionOptions_e.swDelete_Absorbed

    swDoc.FeatureManager.EnableFeatureTree = True
    swDoc.ForceRebuild3 False

    'MsgBox "Top-level components reordered successfully."
    Exit Sub

CleanupAndFail:
    On Error Resume Next
    swDoc.FeatureManager.EnableFeatureTree = True
    swDoc.Extension.DeleteSelection2 swDeleteSelectionOptions_e.swDelete_Absorbed
    MsgBox "Reorder failed. Check assembly and try again."

End Sub

