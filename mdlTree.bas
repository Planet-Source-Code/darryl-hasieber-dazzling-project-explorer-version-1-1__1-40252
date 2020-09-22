Attribute VB_Name = "mdlTree"
Option Explicit
'
Private VBInstance As VBIDE.VBE
Private objFSO As Scripting.FileSystemObject
'
Private pCurrentParentNode As Node
Private pLastNodeAdded As Node
Private pstrProjectFile As String
Private blFileOpen As Boolean
Private objTextStream As Scripting.TextStream
'
Private Enum enNodeType
   ReferenceNode = &H1
   ObjectNode = &H2
   FormNode = &H4
   ModuleNode = &H8
   ClassModuleNode = &H10
   UserControlNode = &H20
   PropertyPageNode = &H40
   UserDocumentNode = &H80
   WebClassNode = &H100
   DataReportNode = &H200
   DHTMLPageNode = &H400
   DesignerNode = &H800
   DataEnvironmentNode = &H1000
   ResFileNode = &H2000
   FileNode = &H4000
   ControlsNode = &H8000
   SubNode = &H10000
   PublicSubNode = &H10001
   PrivateSubNode = &H10002
   FunctionNode = &H10003
   PublicFunctionNode = &H10004
   PrivateFunctionNode = &H10005
   ProcedureNodes = SubNode And FunctionNode And PublicSubNode And PrivateSubNode And PublicFunctionNode And PrivateFunctionNode
   PropertyNode = &H20000
   PublicPropertyNode = &H20001
   PrivatePropertyNode = &H20002
   PropertyNodes = PropertyNode And PublicPropertyNode And PrivatePropertyNode
   EventNode = &H40000
   PublicEventNode = &H40001
   PrivateEventNode = &H40002
   EventNodes = EventNode And PublicEventNode And PrivateEventNode
   ControlEventNode = &H80000
   PublicControlEventNode = &H80001
   PrivateControlEventNode = &H80002
   ControlEventNodes = ControlEventNode And PublicControlEventNode And PrivateControlEventNode
End Enum
'

Private Function GetFileName(NodeType As enNodeType) As String
   Dim strFileName As String
   Dim strTemp As String
   Dim strFile As String
   Dim strPath As String
   '
      Select Case NodeType
      Case ControlEventNode, ControlEventNodes, PublicControlEventNode, PrivateControlEventNode
         If InStr(1, pCurrentParentNode.Key, "~", vbTextCompare) > 0 Then
            strPath = StripFileFromPath(pCurrentParentNode.Key)
            strFile = StripPathFromFile(pCurrentParentNode.Key)
            strFile = Left(strFile, InStr(1, strFile, "~", vbTextCompare) - 1)
            strFileName = strPath & strFile
         Else
            strFileName = pCurrentParentNode.Key
         End If
      Case Else
         If InStr(1, pCurrentParentNode.Key, "~", vbTextCompare) > 0 Then
            strTemp = pCurrentParentNode.Parent.Key
            If InStr(1, strTemp, "~", vbTextCompare) > 0 Then
               If Right(strTemp, 3) = "vbp" Then
                  strPath = StripFileFromPath(pCurrentParentNode.Key)
                  strFile = StripPathFromFile(pCurrentParentNode.Key)
                  'strFile = Right(strFile, Len(strFile) - InStr(1, strFile, "~", vbTextCompare))
                  strFile = Left(strFile, InStr(1, strFile, "~", vbTextCompare) - 1)
                  strFileName = strPath & strFile
               Else
                  strPath = StripFileFromPath(pCurrentParentNode.Key)
                  strFile = StripPathFromFile(pCurrentParentNode.Key)
                  strFile = Left(strFile, InStr(1, strFile, "~", vbTextCompare) - 1)
                  strFileName = strPath & strFile
               End If
            Else
               strFileName = pCurrentParentNode.Parent.Key
            End If
         Else
            strFileName = pCurrentParentNode.Key
         End If
'      Case Else
'         strFileName = pCurrentParentNode.Parent.Key
      End Select
      GetFileName = strFileName
End Function

Private Function GetLastWord(Text As String) As String
   Dim strNewText As String
   '
      strNewText = Text
      Do While InStr(1, Trim(strNewText), " ", vbTextCompare) <> 0
         strNewText = Right(Trim(strNewText), Len(Trim(strNewText)) - InStr(1, Trim(strNewText), " ", vbTextCompare))
      Loop
      GetLastWord = strNewText
End Function

Private Function GetName(Text As String) As String
   Dim strNewText As String
   '
      strNewText = Text
      If InStr(1, Trim(strNewText), "Public", vbTextCompare) <> 0 Then
         strNewText = Right(strNewText, Len(strNewText) - 7)
      End If
      If InStr(1, Trim(strNewText), "Private", vbTextCompare) <> 0 Then
         strNewText = Right(strNewText, Len(strNewText) - 8)
      End If
      If InStr(1, Trim(strNewText), "Sub", vbTextCompare) <> 0 Then
         strNewText = Right(strNewText, Len(strNewText) - 4)
      End If
      If InStr(1, Trim(strNewText), "Function", vbTextCompare) <> 0 Then
         strNewText = Right(strNewText, Len(strNewText) - 9)
      End If
      If InStr(1, strNewText, "(", vbTextCompare) <> 0 Then
         strNewText = Left(strNewText, InStr(1, strNewText, "(", vbTextCompare) - 1)
      End If
      GetName = Trim(strNewText)
End Function

Private Function NodeTypeSearchString(NodeType As enNodeType) As String
Dim strConrolType As String
   Select Case NodeType
   Case ReferenceNode
      NodeTypeSearchString = "Reference="
   Case ObjectNode
      NodeTypeSearchString = "Object="
   Case FormNode
      NodeTypeSearchString = "Form="
   Case ModuleNode
      NodeTypeSearchString = "Module="
   Case ClassModuleNode
      NodeTypeSearchString = "Class="
   Case UserControlNode
      NodeTypeSearchString = "UserControl="
   Case PropertyPageNode
      NodeTypeSearchString = "PropertyPage="
   Case UserDocumentNode
      NodeTypeSearchString = "UserDocument="
   Case WebClassNode
      NodeTypeSearchString = "WebClass="
   Case DataReportNode
      NodeTypeSearchString = "DataReport="
   Case DHTMLPageNode
      NodeTypeSearchString = "DHTMLPage="
   Case DesignerNode
      NodeTypeSearchString = "Designer="
   Case DataEnvironmentNode
      NodeTypeSearchString = "DataEnvironment="
   Case ResFileNode
      NodeTypeSearchString = "ResFile32="
   Case FileNode
      NodeTypeSearchString = "File="
   Case ControlsNode
      NodeTypeSearchString = "Begin "
   Case SubNode
      NodeTypeSearchString = "Sub "
   Case PublicSubNode
      NodeTypeSearchString = "Public Sub "
   Case PrivateSubNode
      NodeTypeSearchString = "Private Sub "
   Case FunctionNode
      NodeTypeSearchString = "Function "
   Case PublicFunctionNode
      NodeTypeSearchString = "Public Function "
   Case PrivateFunctionNode
      NodeTypeSearchString = "Private Function "
'   Case ProcedureNodes
'      NodeTypeSearchString = "Procedures"
   Case PropertyNode
      NodeTypeSearchString = "Property "
   Case PublicPropertyNode
      NodeTypeSearchString = "Public Property "
   Case PrivatePropertyNode
      NodeTypeSearchString = "Private Property "
'   Case PropertyNodes
'      NodeTypeSearchString = "Properties"
   Case EventNode
      NodeTypeSearchString = "Event "
   Case PublicEventNode
      NodeTypeSearchString = "Public Event "
   Case PrivateEventNode
      NodeTypeSearchString = "Private Event "
'   Case EventNodes
'      NodeTypeSearchString = "Events"
   Case ControlEventNode
      NodeTypeSearchString = "Sub " & GetName(pCurrentParentNode.Text)
   Case PublicControlEventNode
      NodeTypeSearchString = "Public Sub " & GetName(pCurrentParentNode.Text)
   Case PrivateControlEventNode
'      If GetName(pCurrentParentNode.Text) = Left(pCurrentParentNode.Parent.Parent, Len(pCurrentParentNode.Parent.Parent) - 4) Then
      If GetName(pCurrentParentNode.Text) = pCurrentParentNode.Parent.Parent Then
         strConrolType = Right(pCurrentParentNode, Len(pCurrentParentNode) - InStr(1, pCurrentParentNode, ".", vbTextCompare))
         If InStr(1, strConrolType, "(", vbTextCompare) <> 0 Then
            strConrolType = Left(strConrolType, InStr(1, strConrolType, "(", vbTextCompare) - 1)
         End If
         strConrolType = Left(Trim(strConrolType), Len(Trim(strConrolType)) - 1)
         Select Case strConrolType
         Case "Form", "MDIForm"
            NodeTypeSearchString = "Private Sub Form"
         Case "Module"
            NodeTypeSearchString = "Private Sub Module"
         Case "Class"
            NodeTypeSearchString = "Private Sub Class"
         Case "UserControl"
            NodeTypeSearchString = "Private Sub UserControl"
         Case "PropertyPage"
            NodeTypeSearchString = "Private Sub PropertyPage"
         Case "UserDocument"
            NodeTypeSearchString = "Private Sub UserDocument"
         Case "WebClass"
            NodeTypeSearchString = "Private Sub WebClass"
         Case "DataReport"
            NodeTypeSearchString = "Private Sub DataReport"
         Case "ActiveReport"
            NodeTypeSearchString = "Private Sub ActiveReport"
         Case "DHTMLPage"
            NodeTypeSearchString = "Private Sub DHTMLPage"
         Case "DTCDesigner"
            NodeTypeSearchString = "Private Sub DTCDesigner"
         Case "DataEnvironment"
            NodeTypeSearchString = "Private Sub DataEnvironment"
'         Case ResFileNode
'            NodeTypeSearchString = "Private Sub Form"
'         Case FileNode
'            NodeTypeSearchString = "Private Sub Form"
'         Case ControlsNode
'            NodeTypeSearchString = "Private Sub Form"
         End Select
      Else
         NodeTypeSearchString = "Private Sub " & GetName(pCurrentParentNode.Text)
      End If
'   Case ControlEventNodes
'      NodeTypeSearchString = "Sub " & GetName(pCurrentParentNode.Text)
   End Select
End Function

Private Function NodeTypeDisplayName(NodeType As enNodeType) As String
   Select Case NodeType
   Case ReferenceNode
      NodeTypeDisplayName = "References"
   Case ObjectNode
      NodeTypeDisplayName = "Objects"
   Case FormNode
      NodeTypeDisplayName = "Forms"
   Case ModuleNode
      NodeTypeDisplayName = "Modules"
   Case ClassModuleNode
      NodeTypeDisplayName = "Class Modules"
   Case UserControlNode
      NodeTypeDisplayName = "User Controls"
   Case PropertyPageNode
      NodeTypeDisplayName = "Property Pages"
   Case UserDocumentNode
      NodeTypeDisplayName = "User Documents"
   Case WebClassNode
      NodeTypeDisplayName = "Web Classes"
   Case DataReportNode
      NodeTypeDisplayName = "Data Reports"
   Case DHTMLPageNode
      NodeTypeDisplayName = "DHTML Pages"
   Case DesignerNode
      NodeTypeDisplayName = "Designers"
   Case DataEnvironmentNode
      NodeTypeDisplayName = "Data Environments"
   Case ResFileNode
      NodeTypeDisplayName = "Resource Files"
   Case FileNode
      NodeTypeDisplayName = "Files"
   Case ControlsNode
      NodeTypeDisplayName = "Controls"
   Case SubNode
      NodeTypeDisplayName = "Procedures"
   Case PublicSubNode
      NodeTypeDisplayName = "Procedures"
   Case PrivateSubNode
      NodeTypeDisplayName = "Procedures"
   Case FunctionNode
      NodeTypeDisplayName = "Procedures"
   Case PublicFunctionNode
      NodeTypeDisplayName = "Procedures"
   Case PrivateFunctionNode
      NodeTypeDisplayName = "Procedures"
'   Case ProcedureNodes
'      NodeTypeDisplayName = "Procedures"
   Case PropertyNode
      NodeTypeDisplayName = "Properties"
   Case PublicPropertyNode
      NodeTypeDisplayName = "Properties"
   Case PrivatePropertyNode
      NodeTypeDisplayName = "Properties"
'   Case PropertyNodes
'      NodeTypeDisplayName = "Properties"
   Case EventNode
      NodeTypeDisplayName = "Events"
   Case PublicEventNode
      NodeTypeDisplayName = "Events"
   Case PrivateEventNode
      NodeTypeDisplayName = "Events"
'   Case EventNodes
'      NodeTypeDisplayName = "Events"
   Case ControlEventNode
      NodeTypeDisplayName = "Control Events"
   Case PublicControlEventNode
      NodeTypeDisplayName = "Control Events"
   Case PrivateControlEventNode
      NodeTypeDisplayName = "Control Events"
'   Case ControlEventNodes
'      NodeTypeDisplayName = "Control Events"
   End Select
End Function

Private Function CreateNodeKey(ItemText As String, NodeType As enNodeType) As String
Dim strInitText As String
Dim strNewText As String
'
   strInitText = ItemText
   Select Case NodeType
   Case ReferenceNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
      strInitText = Right(strInitText, Len(strInitText) - 3)
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 10) & "~" & strInitText
   Case ObjectNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
      strInitText = Right(strInitText, Len(strInitText) - 3)
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 7) & "~" & strInitText
   Case FormNode
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & strInitText
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 5) & "~" & strInitText
   Case ModuleNode
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & GetLastWord(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 5) & "~" & strInitText
   Case ClassModuleNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & GetLastWord(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 6) & "~" & GetLastWord(strInitText)
   Case UserControlNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & GetLastWord(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 12) & "~" & strInitText
   Case PropertyPageNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & GetLastWord(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 13) & "~" & strInitText
   Case UserDocumentNode
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & strInitText
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 5) & "~" & strInitText
   Case WebClassNode
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & strInitText
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 5) & "~" & strInitText
   Case DataReportNode
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & strInitText
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 5) & "~" & strInitText
   Case DHTMLPageNode
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & strInitText
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 5) & "~" & strInitText
   Case DesignerNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & GetLastWord(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 9) & "~" & strInitText
   Case DataEnvironmentNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & GetLastWord(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 0) & "~" & strInitText
   Case ResFileNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
'      strNewText = StripPathFromFile(strInitText)
      strNewText = strInitText
      If InStr(1, strNewText, """", vbTextCompare) = 1 Then
         strNewText = Right(strNewText, Len(strNewText) - 1)
      End If
      If InStr(1, strNewText, """", vbTextCompare) <> 0 Then
         strNewText = Left(strNewText, Len(strNewText) - 1)
      End If
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & GetLastWord(strNewText)
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 10) & "~" & strNewText
   Case FileNode
'      strNewText = pCurrentParentNode.Key & "~" & strInitText
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & GetLastWord(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 0) & "~" & strInitText
   Case ControlsNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 6) & "~" & GetLastWord(strInitText)
   Case SubNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
   Case PublicSubNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
   Case PrivateSubNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 4) & "~" & GetName(strInitText)
   Case FunctionNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
   Case PublicFunctionNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
   Case PrivateFunctionNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
'   Case ProcedureNodes
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
   Case PropertyNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
   Case PublicPropertyNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
   Case PrivatePropertyNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
'   Case PropertyNodes
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 11) & "~" & GetName(strInitText)
   Case EventNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 7) & "~" & GetName(strInitText)
   Case PublicEventNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 7) & "~" & GetName(strInitText)
   Case PrivateEventNode
      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 7) & "~" & GetName(strInitText)
'   Case EventNodes
'      strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 7) & "~" & GetName(strInitText)
   Case ControlEventNode
      'strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 0) & "~" & GetName(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, InStr(1, pCurrentParentNode.Key, "~", vbTextCompare) - 1) & "~" & GetName(strInitText)
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & Left(StripPathFromFile(pCurrentParentNode.Key), InStr(1, StripPathFromFile(pCurrentParentNode.Key), "~", vbTextCompare) - 1) & "~" & GetName(strInitText)
   Case PublicControlEventNode
      'strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 0) & "~" & GetName(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, InStr(1, pCurrentParentNode.Key, "~", vbTextCompare) - 1) & "~" & GetName(strInitText)
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & Left(StripPathFromFile(pCurrentParentNode.Key), InStr(1, StripPathFromFile(pCurrentParentNode.Key), "~", vbTextCompare) - 1) & "~" & GetName(strInitText)
   Case PrivateControlEventNode
      'strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 0) & "~" & GetName(strInitText)
      'strNewText = Left(pCurrentParentNode.Key, InStr(1, pCurrentParentNode.Key, "~", vbTextCompare) - 1) & "~" & GetName(strInitText)
      strNewText = StripFileFromPath(pCurrentParentNode.Key) & Left(StripPathFromFile(pCurrentParentNode.Key), InStr(1, StripPathFromFile(pCurrentParentNode.Key), "~", vbTextCompare) - 1) & "~" & GetName(strInitText)
'   Case ControlEventNodes
'      'strNewText = Left(pCurrentParentNode.Key, Len(pCurrentParentNode.Key) - 0) & "~" & GetName(strInitText)
'      strNewText = Left(pCurrentParentNode.Key, InStr(1, pCurrentParentNode.Key, "~", vbTextCompare) - 1) & "~" & GetName(strInitText)
   End Select
   GetNodeKey = strNewText
End Function

Private Function CreateNodeText(ItemText As String, NodeType As enNodeType, ProjectFile As String) As String
Dim strInitText As String
Dim strNewText As String
'
   strInitText = ItemText
   Select Case NodeType
   Case ReferenceNode
      strNewText = Right(strInitText, Len(strInitText) - 3)
      strNewText = StripPathFromFile(strNewText)
      strNewText = Right(strNewText, Len(strNewText) - InStr(1, strNewText, "#", vbTextCompare))
      If strNewText = "" Then
         strNewText = StripPathFromFile(strInitText)
         strNewText = Left((strNewText), Len(strNewText) - 1)
      End If
   Case ObjectNode
      strInitText = Right(strInitText, Len(strInitText) - 3)
      strNewText = StripPathFromFile(strInitText)
      strNewText = Right(strNewText, Len(strNewText) - InStr(1, strNewText, ";", vbTextCompare))
      strNewText = Trim(strNewText)
      If strNewText = "" Then
         strNewText = StripPathFromFile(strInitText)
         strNewText = Left((strNewText), Len(strNewText) - 1)
      End If
   Case FormNode
'      strNewText = Right(strInitText, Len(strInitText) - 3)
      strNewText = StripPathFromFile(strInitText)
      strNewText = GetObjectName(strInitText, ProjectFile)
   Case ModuleNode
      strNewText = StripPathFromFile(strInitText)
      If InStr(1, strNewText, ";", vbTextCompare) <> 0 Then
         strNewText = Left(strNewText, InStr(1, strNewText, ";", vbTextCompare) - 1)
      End If
   Case ClassModuleNode
      strNewText = StripPathFromFile(strInitText)
      If InStr(1, strNewText, ";", vbTextCompare) <> 0 Then
         strNewText = Left(strNewText, InStr(1, strNewText, ";", vbTextCompare) - 1)
      End If
   Case UserControlNode
      strNewText = StripPathFromFile(strInitText)
   Case PropertyPageNode
      strNewText = StripPathFromFile(strInitText)
   Case UserDocumentNode
      strNewText = StripPathFromFile(strInitText)
   Case WebClassNode
      strNewText = StripPathFromFile(strInitText)
   Case DataReportNode
      strNewText = StripPathFromFile(strInitText)
   Case DHTMLPageNode
      strNewText = StripPathFromFile(strInitText)
   Case DesignerNode
      strNewText = StripPathFromFile(strInitText)
   Case DataEnvironmentNode
      strNewText = StripPathFromFile(strInitText)
   Case ResFileNode
      strNewText = StripPathFromFile(strInitText)
      If InStr(1, strNewText, """", vbTextCompare) = 1 Then
         strNewText = Right(strNewText, Len(strNewText) - 1)
      End If
      If InStr(1, strNewText, """", vbTextCompare) <> 0 Then
         strNewText = Left(strNewText, Len(strNewText) - 1)
      End If
   Case FileNode
      strNewText = StripPathFromFile(strInitText)
      If InStr(1, strNewText, """", vbTextCompare) = 1 Then
         strNewText = Right(strNewText, Len(strNewText) - 1)
      End If
      If InStr(1, strNewText, """", vbTextCompare) <> 0 Then
         strNewText = Left(strNewText, Len(strNewText) - 1)
      End If
   Case ControlsNode
      strNewText = Right(strInitText, Len(strInitText) - InStr(1, strInitText, NodeTypeSearchString(ControlsNode), vbTextCompare) - 5)
'      strNewText = Left(strNewText, InStr(1, strNewText, " ", vbTextCompare) - 1) & " : " & Right(strNewText, Len(strNewText) - InStr(1, strNewText, " ", vbTextCompare))
      strNewText = Right(strNewText, Len(strNewText) - InStr(1, strNewText, " ", vbTextCompare)) & " (" & Left(strNewText, InStr(1, strNewText, " ", vbTextCompare) - 1) & ")"
   Case SubNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PublicSubNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PrivateSubNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case FunctionNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PublicFunctionNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PrivateFunctionNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case ProcedureNodes
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PropertyNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PublicPropertyNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PrivatePropertyNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PropertyNodes
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case EventNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PublicEventNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PrivateEventNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case EventNodes
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case ControlEventNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PublicControlEventNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case PrivateControlEventNode
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   Case ControlEventNodes
'      strNewText = GetName(strInitText)
      strNewText = Left(strInitText, InStr(1, strInitText, "(", vbTextCompare) - 1)
   End Select
   GetNodeText = strNewText
End Function

Private Sub AddTreeRoot()
   Dim strFileName As String
   '
'      strFileName = StripPathFromFile(pstrProjectFile)
      strFileName = GetObjectName(pstrProjectFile, pstrProjectFile)
      Set pCurrentParentNode = gdocDazzExplorer.trvProjExplorer.Nodes.Add(, , pstrProjectFile, strFileName)
      pCurrentParentNode.Tag = pstrProjectFile
      pCurrentParentNode.Expanded = True
      Set pLastNodeAdded = pCurrentParentNode
      '
      If StrComp(Right(pstrProjectFile, 3), "vbg", vbTextCompare) = 0 Then
         Call AddProjectNodes
      Else
         Call SingleProjectAddSubNodes
      End If
End Sub

Private Function NodesToAdd(NodeType As String, outName As String) As Boolean
Dim strLine As String
Dim blFound As Boolean
Dim lngNameStartPos As Long
Dim lngFileStartPos As Long
'
   Do While Not blFound
      If objTextStream.AtEndOfStream Then
         NodesToAdd = False
         Exit Do
      Else
         strLine = objTextStream.ReadLine
         If InStr(1, Trim(strLine), NodeType, vbTextCompare) = 1 Then
            lngNameStartPos = InStr(1, strLine, "=", vbTextCompare)
            outName = Mid(strLine, lngNameStartPos + 1, Len(strLine) - lngFileStartPos - 1)
            blFound = True
            NodesToAdd = True
         End If
      End If
   Loop
End Function

Private Sub AddCategoryNode(NodeType As enNodeType)
   Set pLastNodeAdded = UserDocument.trvProjExplorer.Nodes.Add(pCurrentParentNode.Key, tvwChild, pCurrentParentNode.Key & "~" & Left(NodeTypeSearchString(NodeType), Len(NodeTypeSearchString(NodeType)) - 1), NodeTypeDisplayName(NodeType))
   Set pCurrentParentNode = pLastNodeAdded
   pLastNodeAdded.Expanded = True
End Sub

Private Sub AddNodes(NodeType As enNodeType)
On Error GoTo ErrorHandler
Dim lngChildNode As Long
Dim strNodeItem As String
Dim strNodeKey As String
Dim strNodeText As String
Dim strFileName As String
'
   strFileName = GetFileName(NodeType)
   Call OpenFile(strFileName)
   Do While NodesToAdd(NodeTypeSearchString(NodeType), strNodeItem)
      strNodeKey = GetNodeKey(strNodeItem, NodeType)
      strNodeText = GetNodeText(strNodeItem, NodeType, strFileName)
      Set pLastNodeAdded = UserDocument.trvProjExplorer.Nodes.Add(pCurrentParentNode.Key, tvwChild, strNodeKey, strNodeText)
      pLastNodeAdded.Expanded = True
   Loop
   Call CloseFile
   '
ExitRoutine:
   Exit Sub
ErrorHandler:
   Select Case Err.Number
   Case 35602
      If Not (pCurrentParentNode.Text = NodeTypeDisplayName(ProcedureNodes) And UserDocument.trvProjExplorer.Nodes(strNodeKey).Parent.Parent = NodeTypeDisplayName(ControlsNode)) Then
         Dim strNewText As String
         strNewText = GetLastWord(UserDocument.trvProjExplorer.Nodes(strNodeKey).Text)
         strNewText = Trim(strNewText)
         If Left(Trim(strNewText), 1) = "x" Then
            strNewText = Right(strNewText, Len(strNewText) - 1)
            strNewText = Val(strNewText) + 1
            strNewText = " x" & strNewText
            UserDocument.trvProjExplorer.Nodes(strNodeKey).Text = Left(UserDocument.trvProjExplorer.Nodes(strNodeKey).Text, Len(UserDocument.trvProjExplorer.Nodes(strNodeKey).Text) - Len(GetLastWord(UserDocument.trvProjExplorer.Nodes(strNodeKey).Text))) & strNewText
         Else
            strNewText = " x2"
            UserDocument.trvProjExplorer.Nodes(strNodeKey).Text = UserDocument.trvProjExplorer.Nodes(strNodeKey).Text & strNewText
         End If
         strNewText = ""
      End If
      Resume Next
   Case Else
      Resume ExitRoutine
   End Select
End Sub

Private Sub RemoveEmptyCategoryNode(NodeType As enNodeType)
   If pLastNodeAdded.Text = NodeTypeDisplayName(NodeType) Then
      Call UserDocument.trvProjExplorer.Nodes.Remove(pLastNodeAdded.Index)
   End If
End Sub

Private Sub AddTreeNode(NodeText As String, NodeKey As String, ParentNode As String)
   Set pLastNodeAdded = UserDocument.trvProjExplorer.Nodes.Add(ParentNode, tvwChild, NodeKey, NodeText)
   pLastNodeAdded.Expanded = True
End Sub

Private Sub AddReferenceNodes()
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(ReferenceNode)
   Call AddNodes(ReferenceNode)
   Call RemoveEmptyCategoryNode(ReferenceNode)
   '
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddObjectNodes()
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(ObjectNode)
   Call AddNodes(ObjectNode)
   Call RemoveEmptyCategoryNode(ObjectNode)
   '
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddFormNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(FormNode)
   Call AddNodes(FormNode)
   Call RemoveEmptyCategoryNode(FormNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(FormNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   '
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddModuleNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(ModuleNode)
   Call AddNodes(ModuleNode)
   Call RemoveEmptyCategoryNode(ModuleNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(ModuleNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddClassModuleNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(ClassModuleNode)
   Call AddNodes(ClassModuleNode)
   Call RemoveEmptyCategoryNode(ClassModuleNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(ClassModuleNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddUserControlNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(UserControlNode)
   Call AddNodes(UserControlNode)
   Call RemoveEmptyCategoryNode(UserControlNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(UserControlNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddPropertyPageNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(PropertyPageNode)
   Call AddNodes(PropertyPageNode)
   Call RemoveEmptyCategoryNode(PropertyPageNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(PropertyPageNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddUserDocumentNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(UserDocumentNode)
   Call AddNodes(UserDocumentNode)
   Call RemoveEmptyCategoryNode(UserDocumentNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(UserDocumentNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddWebClassNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(WebClassNode)
   Call AddNodes(WebClassNode)
   Call RemoveEmptyCategoryNode(WebClassNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(WebClassNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddDataReportNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(DataReportNode)
   Call AddNodes(DataReportNode)
   Call RemoveEmptyCategoryNode(DataReportNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(DataReportNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddDHTMLPageNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(DHTMLPageNode)
   Call AddNodes(DHTMLPageNode)
   Call RemoveEmptyCategoryNode(DHTMLPageNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(DHTMLPageNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddDesignerNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(DesignerNode)
   Call AddNodes(DesignerNode)
   Call RemoveEmptyCategoryNode(DesignerNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(DesignerNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddDataEnvironmentNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(DataEnvironmentNode)
   Call AddNodes(DataEnvironmentNode)
   Call RemoveEmptyCategoryNode(DataEnvironmentNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(DataEnvironmentNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddResFileNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(ResFileNode)
   Call AddNodes(ResFileNode)
   Call RemoveEmptyCategoryNode(ResFileNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(ResFileNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddFileNodes()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(FileNode)
   Call AddNodes(FileNode)
   Call RemoveEmptyCategoryNode(FileNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(FileNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControls
         Call AddProcedures
         Call AddProperties
         Call AddEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddControls()
Dim lngChildNode As Long
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(ControlsNode)
   Call AddNodes(ControlsNode)
   Call RemoveEmptyCategoryNode(ControlsNode)
   '
   If pLastNodeAdded.Text <> NodeTypeDisplayName(ControlsNode) Then
      Set pCurrentParentNode = pCurrentParentNode.Child
      For lngChildNode = 1 To pCurrentParentNode.Parent.Children
         If lngChildNode > 1 Then
            Set pCurrentParentNode = pCurrentParentNode.Next
         End If
         If pCurrentParentNode Is Nothing Then
            Exit For
         End If
         Call AddControlEvents
         pCurrentParentNode.Expanded = False
      Next
      pCurrentParentNode.Parent.Expanded = False
   End If
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddProcedures()
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(ProcedureNodes)
   Call AddNodes(SubNode)
   Call AddNodes(PublicSubNode)
   Call AddNodes(PrivateSubNode)
   Call AddNodes(FunctionNode)
   Call AddNodes(PublicFunctionNode)
   Call AddNodes(PrivateFunctionNode)
   Call RemoveEmptyCategoryNode(ProcedureNodes)
   '
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddProperties()
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(PropertyNodes)
   Call AddNodes(PropertyNode)
   Call AddNodes(PublicPropertyNode)
   Call AddNodes(PrivatePropertyNode)
   Call RemoveEmptyCategoryNode(PropertyNodes)
   '
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddEvents()
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddCategoryNode(EventNodes)
   Call AddNodes(EventNode)
   Call AddNodes(PublicEventNode)
   Call AddNodes(PrivateEventNode)
   Call RemoveEmptyCategoryNode(EventNodes)
   '
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub

Private Sub AddControlEvents()
Dim nParent As Node
'
   Set nParent = pCurrentParentNode
   '
   Call AddNodes(ControlEventNode)
   Call AddNodes(PublicControlEventNode)
   Call AddNodes(PrivateControlEventNode)
   '
   pCurrentParentNode.Expanded = False
   Set pCurrentParentNode = nParent
End Sub


Private Function GetObjectName(ObjFile As String, ProjFile As String) As String
   Dim VBProj As VBProject
   Dim VBComp As VBComponent
   Dim CompFile As String
   Dim strName As String
   Dim blFound As Boolean
      If Right(ObjFile, 3) = "vbg" Then
         strName = StripPathFromFile(ProjFile)
         strName = Left(strName, Len(strName) - 4)
      Else
         CompFile = StripFileFromPath(ProjFile) & StripPathFromFile(ObjFile)
         For Each VBProj In VBInstance.VBProjects
            If Right(ObjFile, 3) = "vbp" Then
               If VBProj.FileName = CompFile Then
                  strName = VBProj.Name
                  blFound = True
               End If
            Else
               For Each VBComp In VBProj.VBComponents
                  If VBComp.FileNames(1) = CompFile Then
                     strName = VBComp.Name
                     blFound = True
                  End If
                  If blFound Then Exit For
               Next
            End If
            If blFound Then Exit For
         Next
      End If
      GetObjectName = strName
End Function

