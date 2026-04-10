Attribute VB_Name = "BalloonOrientation1"
'------------------------------------------------------------
' Macro Name: BalloonOrientation
'
' Description:
'   Aligns selected balloons in a SolidWorks drawing either
'   based on their leader arrowhead position:
'       - "Vertical" option aligns balloons left/right (X-axis)
'       - "Horizontal" option aligns balloons up/down (Y-axis)
'
' Usage Instructions:
'   1. Open a SolidWorks drawing document.
'   2. Select one or more balloons (notes) in the drawing.
'   3. Run the macro.
'   4. A UserForm will appear:
'       - Choose alignment type (Vertical or Horizontal).
'       - Click OK to apply or Cancel/X to exit.
'
' Requirements:
'   - SolidWorks drawing must be active.
'   - At least one note (swSelNOTES) must be selected.
'   - Only notes with valid leader data will actually be moved.
'
' Credits:
'   - Jonathan Mendoza (Implementation & Testing)
'   - SolidWorks Macro Expert (GPT-5 Assistant for API guidance)
'
' References:
'   - SolidWorks API Help: IAnnotation::SetPosition2
'   - SolidWorks API Help: IAnnotation::GetLeaderPointsAtIndex
'   - SolidWorks API Help: Note::GetHeight
'------------------------------------------------------------

Option Explicit

' Turn debug logging on/off
Private Const DebugMode As Boolean = False

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swSelMgr As SldWorks.SelectionMgr
Dim swNote As SldWorks.Note
Dim swAnno As SldWorks.Annotation
Dim leaderArray As Variant
Dim balloonX As Double, balloonY As Double, balloonZ As Double
Dim arrowheadX As Double, arrowheadY As Double
Dim noteHeight As Double
Dim centerY As Double, offset As Double, newY As Double
Dim selCount As Long, i As Long
Dim userChoice As String

'---------------------------------------------
' Simple annotation type name helper
' (You can expand this using swAnnotationType_e values)
'---------------------------------------------
Private Function AnnotationTypeName(ByVal annType As Long) As String
    Select Case annType
        Case 6          ' swAnnotationType_e.swNote according to SW help
            AnnotationTypeName = "Note"
        Case Else
            AnnotationTypeName = "Other (" & annType & ")"
    End Select
End Function

Sub main()
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
    
    Set swSelMgr = swModel.SelectionManager
    selCount = swSelMgr.GetSelectedObjectCount2(-1)
    
    ' Validate selection: count only notes (swSelNOTES)
    Dim validCount As Long
    validCount = 0
    
    For i = 1 To selCount
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelNOTES Then
            validCount = validCount + 1
        End If
    Next i
    
    If validCount = 0 Then
        MsgBox "Please select one or more balloons (notes).", vbExclamation
        Exit Sub
    End If
    
    ' Show UserForm for alignment choice
    Dim frm As frmBalloonAlign
    Set frm = New frmBalloonAlign
    
    frm.Cancelled = False
    frm.Show vbModal
    
    If frm.Cancelled = True Then
        Exit Sub
    End If
    
    userChoice = frm.AlignmentChoice   ' "V" or "H"
    
    ' Process selected notes
    For i = 1 To selCount
        
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelNOTES Then
            Set swNote = swSelMgr.GetSelectedObject6(i, -1)
            If swNote Is Nothing Then GoTo NextSelection
            
            Set swAnno = swNote.GetAnnotation
            If swAnno Is Nothing Then GoTo NextSelection
            
            ' --- Debug: annotation info for this selection ---
            If DebugMode Then
                Dim annType As Long
                annType = swAnno.GetType        ' swAnnotationType_e
                Debug.Print "Selection " & i & _
                            " | AnnotationType value = " & annType & _
                            " | " & AnnotationTypeName(annType)
            End If
            ' ------------------------------------------------
            
            ' Get leader points for first leader safely
            leaderArray = swAnno.GetLeaderPointsAtIndex(0)
            
            If Not IsArray(leaderArray) Then
                If DebugMode Then
                    Debug.Print "Selection " & i & " | No leader array, skipping."
                End If
                GoTo NextSelection
            End If
            
            ' Protect against too-short leader array
            If UBound(leaderArray) < 2 Then
                If DebugMode Then
                    Debug.Print "Selection " & i & " | Leader array too short, skipping."
                End If
                GoTo NextSelection
            End If
            
            ' Current balloon position
            balloonX = swAnno.GetPosition(0)
            balloonY = swAnno.GetPosition(1)
            balloonZ = swAnno.GetPosition(2)
            
            ' Get leader arrowhead coordinates
            arrowheadX = leaderArray(UBound(leaderArray) - 2)
            arrowheadY = leaderArray(UBound(leaderArray) - 1)
            
            If userChoice = "V" Then
                ' Vertical alignment: align X to arrowhead (left/right)
                swAnno.SetPosition2 arrowheadX, balloonY, balloonZ
            ElseIf userChoice = "H" Then
                ' Horizontal alignment: align Y to arrowhead center (up/down)
                noteHeight = swNote.GetHeight()
                centerY = balloonY - (noteHeight / 2)
                offset = arrowheadY - centerY
                newY = balloonY + offset
                swAnno.SetPosition2 balloonX, newY, balloonZ
            End If
            
            ' Optional: Make balloon text horizontal
            swNote.Angle = 0
        End If
        
NextSelection:
    Next i
    
    ' Rebuild the drawing so stacked balloons and annotations update
    swModel.ForceRebuild3 False   ' or swModel.EditRebuild3
    
    ' Optional summary:
    'MsgBox validCount & " note(s) processed for alignment (" & _
    '       IIf(userChoice = "H", "Horizontal", "Vertical") & ").", vbInformation
End Sub

