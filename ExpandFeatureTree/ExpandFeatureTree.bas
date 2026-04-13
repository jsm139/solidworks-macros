Attribute VB_Name = "ExpandFeatureTree1"
' ============================================================
' Filename   : ExpandFeatureTree
' Description:
'   Rebuilds the active SOLIDWORKS document, then traverses and
'   expands all nodes in the main FeatureManager design tree.
'
'   - Works only on PART and ASSEMBLY documents.
'   - Performs a rebuild BEFORE and AFTER expanding the tree to
'     ensure the model and tree are fully up to date.
'   - Includes an optional DEBUG_MODE flag to print detailed
'     information about features and components to the Immediate
'     window during traversal.
'
' References:
'   - SOLIDWORKS API Help example:
'       "Traverse FeatureManager design tree and expand/collapse nodes"
'   - FeatureManager tree traversal and ITreeControlItem examples:
'       CodeStack – Traverse feature manager nodes using SOLIDWORKS API
'       CodeStack – Expand body folder node in the FeatureManager Tree
'
' Authors    : Jonathan Mendoza
'              (with assistance from M365 Copilot – SolidWorks Macro Expert)
' Created    : 2026-03-23
' ============================================================

Option Explicit

' ==========================
' USER SETTINGS
' ==========================
' Set this to True when you want to see Debug.Print output
' in the Immediate window during traversal.
Private Const DEBUG_MODE As Boolean = False

Dim traverseLevel As Integer
Dim expandThis As Boolean

Sub main()
    Dim swApp      As SldWorks.SldWorks
    Dim myModel    As SldWorks.ModelDoc2
    Dim featureMgr As SldWorks.FeatureManager
    Dim rootNode   As SldWorks.TreeControlItem
    Dim docType    As Long

    Set swApp = Application.SldWorks
    Set myModel = swApp.ActiveDoc
    
    If myModel Is Nothing Then
        MsgBox "No active document.", vbExclamation
        Exit Sub
    End If

    ' ==========================
    ' 1) ALLOW ONLY PART OR ASSEMBLY
    ' ==========================
    docType = myModel.GetType
    If Not (docType = swDocumentTypes_e.swDocPART Or _
            docType = swDocumentTypes_e.swDocASSEMBLY) Then
        
        MsgBox "This macro only works on Part or Assembly documents.", _
               vbExclamation, "Invalid Document Type"
        Exit Sub
    End If

    ' ==========================
    ' 2) REBUILD FIRST
    ' ==========================
    myModel.ForceRebuild3 False

    ' ==========================
    ' 3) Get FeatureManager tree (TOP pane)
    ' ==========================
    Set featureMgr = myModel.FeatureManager
    Set rootNode = featureMgr.GetFeatureTreeRootItem2(swFeatMgrPane_e.swFeatMgrPaneTop)
    
    If rootNode Is Nothing Then
        MsgBox "Could not get FeatureManager tree root.", vbCritical
        Exit Sub
    End If
    
    ' ==========================
    ' 4) Expand all nodes
    ' ==========================
    expandThis = True
    
    If DEBUG_MODE Then Debug.Print
    traverseLevel = 0
    traverse_node rootNode

    ' ==========================
    ' 5) REBUILD AGAIN AT THE END
    ' ==========================
    myModel.ForceRebuild3 False

End Sub

Private Sub traverse_node(node As SldWorks.TreeControlItem)

    Dim childNode      As SldWorks.TreeControlItem
    Dim featureNode    As SldWorks.Feature
    Dim componentNode  As SldWorks.Component2
    Dim nodeObjectType As Long
    Dim nodeObject     As Object
    Dim restOfString   As String
    Dim indent         As String
    Dim i              As Integer
    Dim displayNodeInfo As Boolean
    Dim compName       As String
    Dim suppr          As Long, supprString As String
    Dim vis            As Long, visString As String
    Dim fixed          As Boolean, fixedString As String
    Dim componentDoc   As Object, docString As String
    Dim refConfigName  As String

    displayNodeInfo = False
    nodeObjectType = node.ObjectType
    Set nodeObject = node.Object

    Select Case nodeObjectType
    
        Case SwConst.swTreeControlItemType_e.swFeatureManagerItem_Feature
            displayNodeInfo = True
            If Not nodeObject Is Nothing Then
                Set featureNode = nodeObject
                restOfString = "[FEATURE: " & featureNode.Name & "]"
            Else
                restOfString = "[FEATURE: object Null?!]"
            End If
        
        Case SwConst.swTreeControlItemType_e.swFeatureManagerItem_Component
            displayNodeInfo = True
            
            If Not nodeObject Is Nothing Then
                Set componentNode = nodeObject
                compName = componentNode.Name2
                If compName = "" Then compName = "???"
                
                ' Suppression state
                suppr = componentNode.GetSuppression
                Select Case suppr
                    Case SwConst.swComponentSuppressionState_e.swComponentFullyResolved
                        supprString = "Resolved"
                    Case SwConst.swComponentSuppressionState_e.swComponentLightweight
                        supprString = "Lightweight"
                    Case SwConst.swComponentSuppressionState_e.swComponentSuppressed
                        supprString = "Suppressed"
                End Select
                
                ' Visibility
                vis = componentNode.Visible
                Select Case vis
                    Case SwConst.swComponentVisibilityState_e.swComponentHidden
                        visString = "Hidden"
                    Case SwConst.swComponentVisibilityState_e.swComponentVisible
                        visString = "Visible"
                End Select
                
                ' Fixed / Floating
                fixed = componentNode.IsFixed
                If fixed = 0 Then
                    fixedString = "Floating"
                Else
                    fixedString = "Fixed"
                End If
                
                ' Loaded / Not Loaded
                Set componentDoc = componentNode.GetModelDoc2
                If componentDoc Is Nothing Then
                    docString = "NotLoaded"
                Else
                    docString = "Loaded"
                End If
                
                ' Referenced configuration
                refConfigName = componentNode.ReferencedConfiguration
                If refConfigName = "" Then
                    refConfigName = "???"
                End If
                
                restOfString = "[COMPONENT: " & compName & " " & docString & " " & _
                               supprString & " " & visString & " " & fixedString & " " & _
                               refConfigName & "]"
            Else
                restOfString = "[COMPONENT: object Null?!]"
            End If
    
        Case Else
            displayNodeInfo = True
            If Not nodeObject Is Nothing Then
                restOfString = "[object type not handled]"
            Else
                restOfString = "[object Null?!]"
            End If
    End Select

    ' Build indentation for hierarchy printing
    For i = 1 To traverseLevel
        indent = indent & "  "
    Next i
    
    ' Only print if DEBUG_MODE is enabled
    If DEBUG_MODE And displayNodeInfo Then
        Debug.Print indent & node.Text & " : " & restOfString
    End If

    ' Expand this node
    On Error Resume Next
    node.Expanded = expandThis
    On Error GoTo 0
    
    ' Recurse into children
    traverseLevel = traverseLevel + 1
    Set childNode = node.GetFirstChild
    
    While Not childNode Is Nothing
        traverse_node childNode
        Set childNode = childNode.GetNext
    Wend

    traverseLevel = traverseLevel - 1

End Sub

