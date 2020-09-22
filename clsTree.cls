VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'This code was written for the public domain by Darryl Hasieber (www.dazzlingsoftware.com)
'You are free to use this project in whole or in part for any purpose you see fit (except toilet paper)
'I hope that this code is not just useful in what it can do but also as an educational tool to show
'what I consider to be good naming of procedures and variables.
'Although potentially complex I hope that people will make sense of it and be able to use this as an
'example to produce their own AddIns.
'If you chose to use this class module as whole I request that you please keep a refernce to me in it
'and not claim this as your own work.
'Apart from being dishonost (legally called plagarism) it could land you in trouble claiming to be
'capable of things you are not.
'One last thing if anyone enhances this in anyway I would greatly appreciate a copy of the updated source code.
'Any communication can be sent via my website www.dazzlingsoftware.com
'
Private VBInstance As VBIDE.VBE
Private WithEvents Tree As TreeView
Attribute Tree.VB_VarHelpID = -1
Private frmImages As frmSupport
Private CurrentParentNode As Node
Private LastNodeAdded As Node
Private CurrentSelectedNode As Node
'
'Enumeration of Object Types / Tree node categories
Private Enum enObjectType
   otProjectGroup
   otProject
   otReference
   otComponent
   otControl
   otProcedure
   otNone
End Enum
'
'I use these constants instead of strings in the code to avoid syntax errors and make maintenance easier
Private Const catProjectGroup As String = "ProjectGroup"
Private Const catProject As String = "Project"
Private Const catReference As String = "Reference"
Private Const catComponent As String = "Component"
Private Const catControl As String = "Control"
Private Const catProcedure As String = "Procedure"
Private Const catTitle As String = "Title"
Private Const cstrComponents As String = "Components"
Private Const cstrReferences As String = "References"
'

Public Property Set VBIDEInstance(VBIDEAppObject As Object)
'Set a reference to the VB IDE instance
   Set VBInstance = VBIDEAppObject
End Property

Public Property Set TreeCtrl(TreeControl As Object)
'Set a reference to the tree control that will be displayed
   Set Tree = TreeControl
   Set Tree.ImageList = frmImages.ImageList1
End Property

Private Sub Class_Initialize()
'Initialize the imagelist that will be used for the tree images
   Set frmImages = New frmSupport
   Load frmImages
End Sub

Private Sub Class_Terminate()
'Release references
   Set VBInstance = Nothing
   Set Tree = Nothing
   Unload frmImages
   Set frmImages = Nothing
   Set CurrentParentNode = Nothing
   Set LastNodeAdded = Nothing
   Set CurrentSelectedNode = Nothing
End Sub

Private Function ClassObjectsInitialized(outErrorString As String) As Boolean
'Check that all the properties and objects required to run this add-in have been set
'This was used initially in development to ensure That I was setting everything that would be needed by this module.
'It has been left here incase this Class is used again in another project.
On Error Resume Next
   Dim strErrors As String
   Dim blMissingItems As Boolean
   '
      strErrors = vbNullString
      blMissingItems = False
      '
      If VBInstance Is Nothing Then
         strErrors = "VBIDEInstance is Not Set" & vbCrLf
         blMissingItems = True
      End If
      '
      If Tree Is Nothing Then
         strErrors = "TreeControl is Not Set" & vbCrLf
         blMissingItems = True
      End If
      '
      If blMissingItems Then
         strErrors = "The Following are missing:" & vbCrLf & strErrors
      End If
      '
      outErrorString = strErrors
      ClassObjectsInitialized = Not blMissingItems
End Function

Public Sub BuildTree()
'This is the entry point into the Build Tree class of the AddIn
On Error GoTo ErrorHandler
   Dim strErrorString As String
   '
      Screen.MousePointer = MousePointerConstants.vbHourglass
      If ClassObjectsInitialized(strErrorString) Then
         Tree.Scroll = True
         Tree.Nodes.Clear
         Call AddTreeRoot
      Else
         MsgBox strErrorString
      End If
      '
ExitRoutine:
   Screen.MousePointer = MousePointerConstants.vbDefault
   Exit Sub
ErrorHandler:
   MsgBox "The Following Error occurred:" & vbCrLf & Err.Number & ": " & Err.Description
   Resume ExitRoutine
End Sub

Private Sub AddTreeRoot()
'Add the Tree's Root Node
   Dim strFileName As String
   Dim strSolutionFile As String
   '
      strSolutionFile = VBInstance.VBProjects.FileName
      strFileName = StripPathFromFile(strSolutionFile)
      If AddNode(strFileName, strSolutionFile, True) Then
         Set CurrentParentNode = LastNodeAdded
         LastNodeAdded.Expanded = True
         Call SetNodeTag(Nothing, otProjectGroup)
         Call SetNodeImage
         Call SetNodeColorsFont(otProjectGroup)
         Call AddProjectNodes
      End If
End Sub

Private Sub AddProjectNodes()
'Add a node for each Project
On Error GoTo ErrorHandler
   Dim strNewNode As String
   Dim strNewNodeKey As String
   Dim nPrevCurrentParentNode As Node
   Dim Proj As VBProject
   '
      Set nPrevCurrentParentNode = CurrentParentNode
      'For Each Project Add A Node
      For Each Proj In VBInstance.VBProjects
         strNewNode = Proj.Name
         '
         'This prevents and error if the project has not yet been saved
         If Proj.FileName <> "" Then
            strNewNodeKey = Proj.FileName
         Else
            strNewNodeKey = Proj.Name
         End If
         '
         If AddNode(strNewNode, strNewNodeKey) Then
            Call SetNodeTag(Proj, otProject)
            Call SetNodeImage
            Call SetNodeColorsFont(otProject)
            Set CurrentParentNode = LastNodeAdded
         End If
         '
         Call AddDummyNode
         Set CurrentParentNode = nPrevCurrentParentNode
      Next
      '
      'This if statement added to prevent error on NO project
      If VBInstance.VBProjects.Count <> 0 Then
         'This if statement added to prevent error on an unsaved project
         If VBInstance.VBProjects.StartProject.FileName = "" Then
            Tree.Nodes(VBInstance.VBProjects.StartProject.Name).ForeColor = RGB(0, 0, 255)
            Tree.Nodes(VBInstance.VBProjects.StartProject.Name).Bold = True
         Else
            Tree.Nodes(VBInstance.VBProjects.StartProject.FileName).ForeColor = RGB(0, 0, 255)
            Tree.Nodes(VBInstance.VBProjects.StartProject.FileName).Bold = True
         End If
      End If
      '
ExitRoutine:
   Set CurrentParentNode = nPrevCurrentParentNode
   Set nPrevCurrentParentNode = Nothing
   Set Proj = Nothing
   Exit Sub
ErrorHandler:
   MsgBox "The Following Error occurred in clsTree.AddProjectNodes:" & vbCrLf & Err.Number & ": " & Err.Description
   Resume ExitRoutine
End Sub

Private Sub AddReferenceNode(Proj As VBProject, Ref As Reference)
'Add a Reference Node
   Dim strNewNode As String
   Dim strNewNodeKey As String
   Dim nPrevCurrentParentNode As Node
   Dim lngItem As Long
   '
      Set nPrevCurrentParentNode = CurrentParentNode
      strNewNode = Ref.Name & " (" & Ref.Description & ")"
      '
      On Error Resume Next
      strNewNodeKey = Ref.FullPath
      On Error GoTo ErrorHandler
      If strNewNodeKey = vbNullString Then
         For lngItem = 1 To Proj.Collection.Count
            If Proj.Collection(lngItem).Name = Ref.Name Then
               strNewNodeKey = CurrentParentNode.Key & "|" & Proj.Collection(lngItem).FileName
               Exit For
            Else
               strNewNodeKey = CurrentParentNode.Key & "|" & Ref.Guid
            End If
         Next
      Else
         strNewNodeKey = CurrentParentNode.Key & "|" & Ref.FullPath
      End If
      '
      If AddNode(strNewNode, strNewNodeKey) Then
         Call SetNodeTag(Ref, otReference)
         Call SetNodeImage
         Call SetNodeColorsFont(otReference)
      End If
      '
ExitRoutine:
   Set CurrentParentNode = nPrevCurrentParentNode
   Set nPrevCurrentParentNode = Nothing
   Exit Sub
ErrorHandler:
   Set CurrentParentNode = nPrevCurrentParentNode
   Err.Raise Err.Number
End Sub

Private Sub AddComponentNodes(Comp As VBComponent)
'Add a Component Node
   Dim strNewNode As String
   Dim strNewNodeKey As String
   Dim nPrevCurrentParentNode As Node
   Dim Frm As VBForm
   '
      Set nPrevCurrentParentNode = CurrentParentNode
      strNewNode = Comp.Name
      If strNewNode = vbNullString Then
         strNewNode = StripPathFromFile(Comp.FileNames(1))
      End If
      '
      'Added this iif statement to handle unsaved components where filename = "")
      strNewNodeKey = IIf(Comp.FileNames(1) = "", Comp.Name, Comp.FileNames(1))  'Left this in but I think this should modified slightly using maybe #i instead of | to seperate individual files
      If AddNode(strNewNode, strNewNodeKey) Then
         Call SetNodeTag(Comp, otComponent)
         Call SetNodeImage
         Call SetNodeColorsFont(otComponent)
         Set CurrentParentNode = LastNodeAdded
         Select Case Comp.Type
         Case vbext_ComponentType.vbext_ct_ActiveXDesigner
            'TODO
            'Designers can be of different types etc an because of the variety etc I don't quite know what to do here
            'or how to handle these so I have left these out. Bad excuse I know but the truth.
            'If anyone does sort this out I would Love to see how etc.
         Case vbext_ComponentType.vbext_ct_ClassModule
            'Do Nothing. All Members are handled at the end of this procedure.
         Case vbext_ComponentType.vbext_ct_DocObject
            'Check if there are any Controls, if so add a dummy node to so that tree node has a '+'
            Set Frm = Comp.Designer
            If Frm.VBControls.Count <> 0 Then
               Call AddDummyNode
            End If
         Case vbext_ComponentType.vbext_ct_MSForm
            'Check if there are any Controls, if so add a dummy node to so that tree node has a '+'
            Set Frm = Comp.Designer
            If Frm.VBControls.Count <> 0 Then
               Call AddDummyNode
            End If
         Case vbext_ComponentType.vbext_ct_PropPage
            'Check if there are any Controls, if so add a dummy node to so that tree node has a '+'
            Set Frm = Comp.Designer
            If Frm.VBControls.Count <> 0 Then
               Call AddDummyNode
            End If
         Case vbext_ComponentType.vbext_ct_RelatedDocument
            'TODO:
            'Not sure what to do here so... do nothing...bad copout I know
            'but how many people even know what a related document is never mind actually use them.
            'MSDN describes as being a text document that could be code or just plain text.
            'Well that was Helpful I know.
         Case vbext_ComponentType.vbext_ct_ResFile
            'TODO:
            'Not sure what to do here so... do nothing...bad copout I know
            'I don't think many people use resource files and that is why this doesn't concern me too much.
            'Well that was Helpful I know.
            'This requires another AddIn to read it and I haven't yet figured out how to do this.
            'Not important to me at the moment.
         Case vbext_ComponentType.vbext_ct_StdModule
            'Do Nothing. All Members are handled at the end of this procedure.
         Case vbext_ComponentType.vbext_ct_UserControl
            'Check if there are any Controls, if so add a dummy node to so that tree node has a '+'
            Set Frm = Comp.Designer
            If Frm.VBControls.Count <> 0 Then
               Call AddDummyNode
            End If
         Case vbext_ComponentType.vbext_ct_VBForm
            'Check if there are any Controls, if so add a dummy node to so that tree node has a '+'
            Set Frm = Comp.Designer
            If Frm.VBControls.Count <> 0 Then
               Call AddDummyNode
            End If
         Case vbext_ComponentType.vbext_ct_VBMDIForm
            'Check if there are any Controls, if so add a dummy node to so that tree node has a '+'
            Set Frm = Comp.Designer
            If Frm.VBControls.Count <> 0 Then
               Call AddDummyNode
            End If
         End Select
         '
         If Not Comp.CodeModule Is Nothing Then
            'Check if there are any Members, if so add a dummy node to so that tree node has a '+'
            If Comp.CodeModule.Members.Count <> 0 Then
               Call AddDummyNode
            End If
         End If
      End If
      Set CurrentParentNode = nPrevCurrentParentNode
      Set nPrevCurrentParentNode = Nothing
      Set Frm = Nothing
End Sub

Private Sub ControlAddControlNode(Comp As VBComponent, Ctrl As VBControl, Optional IsContained As Boolean = False)
'Add a Control Node - Controller procedure
'I originally had ControlAddControlNode and AddControlNode as one procedure but could not get the recursion to work
'properly By splitting the and adding the input property IsContained to the recursion controller procedure I was able to achieve the desired tree structure.
On Error GoTo ErrorHandler
   'If the Control is contained within a control then the tree must reflect this
   'We use two methods to achieve this one is a variable that comes is passed in 'IsContained'
   'The other is we test 'Ctrl.Container' to double check 'IsContained' if 'IsContained' = False
   If IsContained Then
      Call AddControlNode(Comp, Ctrl, True)
   Else
      On Error Resume Next
      'This is to ensure we do not add a control which is housed within another control
      '(Check that the control is not within another control)
      If Ctrl.Container = "" Then
         On Error GoTo ErrorHandler
         Call AddControlNode(Comp, Ctrl, True)
      End If
   End If
   '
ExitRoutine:
   Exit Sub
ErrorHandler:
   MsgBox Err.Description, , "ControlAddControlNode"
End Sub

Private Sub AddControlNode(Comp As VBComponent, Ctrl As VBControl, Optional IsContained As Boolean = False)
'Add a Control Node - This is the procedure that actually adds the Node
   Dim strNewNode As String
   Dim strNewNodeKey As String
   Dim SubCtrl As VBControl
   Dim Memb As Member
   Dim nPrevCurrentParentNode As Node
   Dim lngInStrPos As Long
   '
      Set nPrevCurrentParentNode = CurrentParentNode
      If Ctrl.Properties("Index").Value = -1 Then
         strNewNode = Ctrl.Properties("Name").Value
      Else
         strNewNode = Ctrl.Properties("Name").Value & "(" & Ctrl.Properties("Index").Value & ")"
      End If
      strNewNode = catControl & ": " & strNewNode
      '
      strNewNodeKey = nPrevCurrentParentNode.Key & "|" & Ctrl.Properties("Name").Value & "|" & Ctrl.Properties("Index").Value
      '
      If AddNode(strNewNode, strNewNodeKey) Then
         Call SetNodeTag(Ctrl, otControl)
         Call SetNodeImage
         Call SetNodeColorsFont(otControl)
         LastNodeAdded.Expanded = True
         Set CurrentParentNode = LastNodeAdded
         If Ctrl.ContainedVBControls.Count <> 0 Then
            For Each SubCtrl In Ctrl.ContainedVBControls
               Call ControlAddControlNode(Comp, SubCtrl, True)
            Next
         End If
         If Comp.CodeModule.Find(Ctrl.Properties("Name").Value & "_", 1, 1, -1, -1) Then
            'TODO: THIS LOOP IS SLOW ON A LARGE CODE MODULE
            For Each Memb In Comp.CodeModule.Members
               lngInStrPos = InStr(1, Memb.Name, Ctrl.Properties("Name").Value, vbTextCompare)
               If InStr(1, Memb.Name, Ctrl.Properties("Name").Value, vbTextCompare) <> 0 Then
                  Call AddMemberNode(Memb)
               End If
            Next
         End If
      End If
      Set CurrentParentNode = nPrevCurrentParentNode
      Set SubCtrl = Nothing
      Set Memb = Nothing
      Set nPrevCurrentParentNode = Nothing
End Sub

Private Sub AddMemberNode(Mem As Member)
'Add a Member Node.
On Error GoTo ExitRoutine
   Dim strNewNode As String
   Dim strNewNodeKey As String
   Dim strProjectName As String
   Dim strComponentName As String
   Dim nPrevCurrentParentNode As Node
   Dim lngCodeLocation As Long
   '
      Set nPrevCurrentParentNode = CurrentParentNode
      '
      If Mem.Type = vbext_mt_Property Then
         strProjectName = GetProjectNameFromTag(LastNodeAdded.Tag)
         strComponentName = GetComponentNameFromTag(LastNodeAdded.Tag)
         '
         On Error Resume Next
         lngCodeLocation = 0
         lngCodeLocation = VBInstance.VBProjects(strProjectName).VBComponents(strComponentName).CodeModule.ProcBodyLine(Mem.Name, vbext_pk_Get)
         On Error GoTo ExitRoutine
         If lngCodeLocation <> 0 Then
            strNewNode = MemberTypeAsString(Mem.Type) & " Get: " & Mem.Name & "  (" & ScopeAsString(Mem.Scope) & ")"
            strNewNodeKey = VBInstance.VBProjects(GetProjectNameFromTag(CurrentParentNode.Tag)).VBComponents(GetComponentNameFromTag(CurrentParentNode.Tag)).FileNames(1) & "|Get" & Mem.Name
            '
            If AddNode(strNewNode, strNewNodeKey) Then
               Call SetNodeTag(Mem, otProcedure)
               Call SetNodeImage
               Call SetNodeColorsFont(otProcedure)
               LastNodeAdded.Expanded = True
            End If
         End If
         '
         On Error Resume Next
         lngCodeLocation = 0
         lngCodeLocation = VBInstance.VBProjects(strProjectName).VBComponents(strComponentName).CodeModule.ProcBodyLine(Mem.Name, vbext_pk_Let)
         On Error GoTo ExitRoutine
         If lngCodeLocation <> 0 Then
            strNewNode = MemberTypeAsString(Mem.Type) & " Let: " & Mem.Name & "  (" & ScopeAsString(Mem.Scope) & ")"
            strNewNodeKey = VBInstance.VBProjects(GetProjectNameFromTag(CurrentParentNode.Tag)).VBComponents(GetComponentNameFromTag(CurrentParentNode.Tag)).FileNames(1) & "|Let" & Mem.Name
            '
            If AddNode(strNewNode, strNewNodeKey) Then
               Call SetNodeTag(Mem, otProcedure)
               Call SetNodeImage
               Call SetNodeColorsFont(otProcedure)
               LastNodeAdded.Expanded = True
            End If
         End If
         '
         On Error Resume Next
         lngCodeLocation = 0
         lngCodeLocation = VBInstance.VBProjects(strProjectName).VBComponents(strComponentName).CodeModule.ProcBodyLine(Mem.Name, vbext_pk_Set)
         On Error GoTo ExitRoutine
         If lngCodeLocation <> 0 Then
            strNewNode = MemberTypeAsString(Mem.Type) & " Set: " & Mem.Name & "  (" & ScopeAsString(Mem.Scope) & ")"
            strNewNodeKey = VBInstance.VBProjects(GetProjectNameFromTag(CurrentParentNode.Tag)).VBComponents(GetComponentNameFromTag(CurrentParentNode.Tag)).FileNames(1) & "|Set" & Mem.Name
            '
            If AddNode(strNewNode, strNewNodeKey) Then
               Call SetNodeTag(Mem, otProcedure)
               Call SetNodeImage
               Call SetNodeColorsFont(otProcedure)
               LastNodeAdded.Expanded = True
            End If
         End If
      Else
         strNewNode = MemberTypeAsString(Mem.Type) & ": " & Mem.Name & "  (" & ScopeAsString(Mem.Scope) & ")"
         'We give all members the same key format FileName|MemberName.
         'Because we add Control members first they will not be duplicated under the Component
         strNewNodeKey = VBInstance.VBProjects(GetProjectNameFromTag(CurrentParentNode.Tag)).VBComponents(GetComponentNameFromTag(CurrentParentNode.Tag)).FileNames(1) & "|" & Mem.Name
         '
         If AddNode(strNewNode, strNewNodeKey) Then
            Call SetNodeTag(Mem, otProcedure)
            Call SetNodeImage
            Call SetNodeColorsFont(otProcedure)
            LastNodeAdded.Expanded = True
         End If
      End If
      '
      Set CurrentParentNode = nPrevCurrentParentNode
      Set nPrevCurrentParentNode = Nothing
ExitRoutine:
End Sub

Private Sub AddReferenceCategoryNode(Proj As VBProject)
'Add Reference Category Node
   Dim nPrevCurrentParentNode As Node
   '
      Set nPrevCurrentParentNode = CurrentParentNode
      If AddNode(cstrReferences, CurrentParentNode.Key & "|" & cstrReferences) Then
         Call SetNodeTag(Nothing, otNone)
         Call SetNodeImage
         Call SetNodeColorsFont(otNone)
         Set CurrentParentNode = LastNodeAdded
         'Check if there are any References, if so add a dummy node to so that tree node has a '+'
         If Proj.References.Count <> 0 Then
            Call AddDummyNode
         End If
      End If
      Set CurrentParentNode = nPrevCurrentParentNode
      Set nPrevCurrentParentNode = Nothing
End Sub

Private Sub AddComponentCategoryNode(Proj As VBProject)
'Add Component Category Node
   Dim nPrevCurrentParentNode As Node
   '
      Set nPrevCurrentParentNode = CurrentParentNode
      If AddNode(cstrComponents, CurrentParentNode.Key & "|" & cstrComponents) Then
         Call SetNodeTag(Nothing, otNone)
         Call SetNodeImage
         Call SetNodeColorsFont(otNone)
         Set CurrentParentNode = LastNodeAdded
         'Check if there are any Components, if so add a dummy node to so that tree node has a '+'
         If Proj.VBComponents.Count <> 0 Then
            Call AddDummyNode
         End If
      End If
      Set CurrentParentNode = nPrevCurrentParentNode
      Set nPrevCurrentParentNode = Nothing
End Sub

Private Sub AddDummyNode()
'TODO
'Using a static here could potentially be a problem, at some point we should reset this to 0
'The problem is when, what trigger, one option is to split this into two parts.
'The first part will reset i the second part is this procedure.
'There will be a performance sacrifice here but I can't see another way at the moment.
   Static i As Long
      If Not AddNode("Dummy", CurrentParentNode.Key & "|" & "Dummy" & i) Then
         i = i + 1
         Call AddDummyNode
      End If
End Sub

Private Function AddNode(NodeText As String, NodeKey As String, Optional Root As Boolean = False) As Boolean
'Add a Node to the Tree, basically just a wrapper for Tree.Nodes.Add
On Error GoTo ErrorHandler
   If Root Then
      Set LastNodeAdded = Tree.Nodes.Add(, , NodeKey, NodeText)
   Else
      Set LastNodeAdded = Tree.Nodes.Add(CurrentParentNode.Key, tvwChild, NodeKey, NodeText)
   End If
   LastNodeAdded.Sorted = True
   AddNode = True
ExitRoutine:
   DoEvents
   Exit Function
ErrorHandler:
   AddNode = False
   Resume ExitRoutine
End Function

Private Sub ClearAllChildNodes(Node As MSComctlLib.Node)
   Do While Node.Children
      Call Tree.Nodes.Remove(Node.Child.Index)
   Loop
End Sub

Public Sub PopulateNode(Node As MSComctlLib.Node)
'Populate node.
On Error GoTo ErrorHandler
   Dim strCategory As String
   Dim strProjectName As String
   Dim strComponentName As String
   Dim strControlName As String
   Dim lngType As Long
   Dim Frm As VBIDE.VBForm
   Dim Proj As VBIDE.VBProject
   Dim Ref As VBIDE.Reference
   Dim Comp As VBIDE.VBComponent
   Dim Memb As VBIDE.Member
   Dim Ctrl As VBIDE.VBControl
   '
      Set CurrentParentNode = Node
      Call ClearAllChildNodes(Node)
      '
      strProjectName = GetProjectNameFromTag(Node.Tag)
      If strProjectName <> "" Then
         Set Proj = VBInstance.VBProjects(strProjectName)
      End If
      '
      strComponentName = GetComponentNameFromTag(Node.Tag)
      If strComponentName <> "" Then
         Set Comp = Proj.VBComponents(strComponentName)
      End If
      '
      strCategory = GetCategoryFromTag(Node.Tag)
      Select Case strCategory
      Case catComponent
         lngType = GetTypeFromTag(Node.Tag)
         Select Case lngType
         Case vbext_ComponentType.vbext_ct_ActiveXDesigner
            'TODO
            'Designers can be of different types etc an because of the variety etc I don't quite know what to do here
            'or how to handle these so I have left these out. Bad excuse I know but the truth.
            'If anyone does sort this out I would Love to see how etc.
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         Case vbext_ComponentType.vbext_ct_ClassModule
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         Case vbext_ComponentType.vbext_ct_DocObject
            'For each Control add a Node
            Set Frm = Comp.Designer
            For Each Ctrl In Frm.VBControls
               Call ControlAddControlNode(Comp, Ctrl)
            Next
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         Case vbext_ComponentType.vbext_ct_MSForm
            'For each Control add a Node
            Set Frm = Comp.Designer
            For Each Ctrl In Frm.VBControls
               Call ControlAddControlNode(Comp, Ctrl)
            Next
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         Case vbext_ComponentType.vbext_ct_PropPage
            'For each Control add a Node
            Set Frm = Comp.Designer
            For Each Ctrl In Frm.VBControls
               Call ControlAddControlNode(Comp, Ctrl)
            Next
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         Case vbext_ComponentType.vbext_ct_RelatedDocument
            'TODO
            'Not sure what to do here so... do nothing...bad copout I know
            'but how many people even know what a related document is never mind actually use them.
            'MSDN describes as being a text document that could be code or just plain text.
            'Well that was Helpful I know.
            'This requires another AddIn to read it and I haven't yet figured out how to do this.
            'Not important to me at the moment.
         Case vbext_ComponentType.vbext_ct_ResFile
            'TODO
            'Another addin needs to be loaded for this to be opened so I am leaving this out.
            'I amreally not concerned about this either
         Case vbext_ComponentType.vbext_ct_StdModule
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         Case vbext_ComponentType.vbext_ct_UserControl
            'For each Control add a Node
            Set Frm = Comp.Designer
            For Each Ctrl In Frm.VBControls
               Call ControlAddControlNode(Comp, Ctrl)
            Next
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         Case vbext_ComponentType.vbext_ct_VBForm
            'For each Control add a Node
            Set Frm = Comp.Designer
            For Each Ctrl In Frm.VBControls
               Call ControlAddControlNode(Comp, Ctrl)
            Next
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         Case vbext_ComponentType.vbext_ct_VBMDIForm
            'For each Control add a Node
            Set Frm = Comp.Designer
            For Each Ctrl In Frm.VBControls
               Call ControlAddControlNode(Comp, Ctrl)
            Next
            'For each Member add a Node
            For Each Memb In Comp.CodeModule.Members
               Call AddMemberNode(Memb)
            Next
         End Select
      Case catControl
         Set CurrentParentNode = Node.Parent
         Set Frm = Comp.Designer
         Call GetControlName(Node.Text, strControlName)
         Set Ctrl = Frm.VBControls(strControlName)
         strCategory = GetCategoryFromTag(Node.Parent.Tag)
         Call Tree.Nodes.Remove(Node.Index)
         Call ControlAddControlNode(Comp, Ctrl, IIf(strCategory = catControl, True, False))
      Case catProcedure
         'We cannot populate this node as it is the lowest level.
         'An Item may also have changed it's name so we cannot refresh the node, the best is to refresh the parent node.
         Call PopulateNode(Node.Parent)
      Case catProject
         'Add Component and Refernce Category Heading Nodes
         Call AddReferenceCategoryNode(Proj)
         Call AddComponentCategoryNode(Proj)
      Case catProjectGroup
         'Refresh all projects.
         Set CurrentParentNode = Node.Root
         Call AddProjectNodes
      Case catReference
         'We cannot populate this node as it is the lowest level.
         'An Item may also have changed it's name so we cannot refresh the node, the best is to refresh the parent node.
         Call PopulateNode(Node.Parent)
      Case catTitle
         Select Case Node.Text
         Case cstrReferences
            'For each Reference Add a Node
            For Each Ref In Proj.References
               Call AddReferenceNode(Proj, Ref)
            Next
         Case cstrComponents
            'For each Component Add a Node
            For Each Comp In Proj.VBComponents
               Call AddComponentNodes(Comp)
            Next
         End Select
      End Select
ExitRoutine:
   Set Frm = Nothing
   Set Proj = Nothing
   Set Ref = Nothing
   Set Comp = Nothing
   Set Memb = Nothing
   Set Ctrl = Nothing
   Exit Sub
ErrorHandler:
   'If an error has occurred it is probably because the node is no longer valid and needs to be refreshed
   Call PopulateNode(Node.Parent)
   Resume ExitRoutine
End Sub

Private Sub SetNodeTag(ProjectObject, ObjectType As enObjectType)
'Set the Tag for the Last Added Node
On Error GoTo ErrorHandler
   Dim strFileName As String
   Dim strProjectName As String
   Dim strComponentName As String
   Dim strObjectType As String
   Dim lngIndex As Long
   Dim strClassName As String
   Dim strCodeLocation As String
   '
      Select Case ObjectType
      Case enObjectType.otProjectGroup
         strProjectName = ""
         strComponentName = ""
         strFileName = VBInstance.VBProjects.FileName
         strObjectType = ""
         strCodeLocation = ""
         LastNodeAdded.Tag = "Category=" & catProjectGroup & "|Project=" & strProjectName & "|Component=" & strComponentName & "|FileName=" & strFileName & "|Type=" & strObjectType & "|CodeLocation=" & strCodeLocation & "|"
      Case enObjectType.otProject
         strProjectName = ProjectObject.Name
         strComponentName = ""
         strFileName = ProjectObject.FileName
         strObjectType = ProjectObject.Type
         strCodeLocation = ""
         LastNodeAdded.Tag = "Category=" & catProject & "|Project=" & strProjectName & "|Component=" & strComponentName & "|FileName=" & strFileName & "|Type=" & strObjectType & "|CodeLocation=" & strCodeLocation & "|"
      Case enObjectType.otReference
         strProjectName = LastNodeAdded.Parent.Parent.Text
         strComponentName = ProjectObject.Name
         On Error Resume Next
         strFileName = ProjectObject.FullPath
         On Error GoTo ErrorHandler
         strObjectType = ProjectObject.Type
         strCodeLocation = ""
         LastNodeAdded.Tag = "Category=" & catReference & "|Project=" & strProjectName & "|Component=" & strComponentName & "|FileName=" & strFileName & "|Type=" & strObjectType & "|CodeLocation=" & strCodeLocation & "|"
      Case enObjectType.otComponent
         strProjectName = LastNodeAdded.Parent.Parent.Text
         strComponentName = LastNodeAdded.Text
         For lngIndex = 1 To ProjectObject.FileCount
            strFileName = ProjectObject.FileNames(lngIndex) & ";"
         Next
         strObjectType = ProjectObject.Type
         strCodeLocation = ""
         LastNodeAdded.Tag = "Category=" & catComponent & "|Project=" & strProjectName & "|Component=" & strComponentName & "|FileName=" & strFileName & "|Type=" & strObjectType & "|CodeLocation=" & strCodeLocation & "|"
      Case enObjectType.otControl
         strProjectName = GetProjectNameFromTag(LastNodeAdded.Parent.Tag)
         strComponentName = GetComponentNameFromTag(LastNodeAdded.Parent.Tag)
         strFileName = GetFileNameFromTag(LastNodeAdded.Parent.Tag)
         strObjectType = ProjectObject.ControlType
         strCodeLocation = ""
         LastNodeAdded.Tag = "Category=" & catControl & "|Project=" & strProjectName & "|Component=" & strComponentName & "|FileName=" & strFileName & "|Type=" & strObjectType & "|CodeLocation=" & strCodeLocation & "|"
      Case enObjectType.otProcedure
         strProjectName = GetProjectNameFromTag(LastNodeAdded.Parent.Tag)
         strComponentName = GetComponentNameFromTag(LastNodeAdded.Parent.Tag)
         strFileName = GetFileNameFromTag(LastNodeAdded.Parent.Tag)
         strObjectType = ProjectObject.Type
         lngIndex = -1
         'No Longer store code location in tag as it goes out of date very quickly
         'call GetCodeLocation(strProjectName,strComponentName,ProjectObject.Name)
         LastNodeAdded.Tag = "Category=" & catProcedure & "|Project=" & strProjectName & "|Component=" & strComponentName & "|FileName=" & strFileName & "|Type=" & strObjectType & "|CodeLocation=" & strCodeLocation & "|"
      Case enObjectType.otNone
         strProjectName = LastNodeAdded.Parent.Text
         strComponentName = ""
         strFileName = ""
         strObjectType = ""
         strCodeLocation = ""
         LastNodeAdded.Tag = "Category=" & catTitle & "|Project=" & strProjectName & "|Component=" & strComponentName & "|FileName=" & strFileName & "|Type=" & strObjectType & "|CodeLocation=" & strCodeLocation & "|"
      End Select
      '
ExitRoutine:
   DoEvents
   Exit Sub
ErrorHandler:
   MsgBox Err.Description, , "SetNodeTag"
End Sub

Private Sub SetNodeImage()
'Set the image for the Node inside the LastAddedNode
   Dim strFileName As String
   Dim strExtension As String
   Dim strCategory As String
   Dim lstimg As ListImage
   Dim lngType As Long
   Dim strName As String
   Dim strClass As String
   '
      strCategory = GetCategoryFromTag(LastNodeAdded.Tag)
      Select Case strCategory
      Case catProjectGroup
         LastNodeAdded.Image = "Project_Group"
      Case catProject
         lngType = GetTypeFromTag(LastNodeAdded.Tag)
         Select Case lngType
         Case vbext_ProjectType.vbext_pt_StandardExe
            LastNodeAdded.Image = "Project_EXE"
         Case vbext_ProjectType.vbext_pt_ActiveXDll
            LastNodeAdded.Image = "Project_ActiveXDLL"
         Case vbext_ProjectType.vbext_pt_ActiveXExe
            LastNodeAdded.Image = "Project_ActiveXEXE"
         Case vbext_ProjectType.vbext_pt_ActiveXControl
            LastNodeAdded.Image = "Project_UserControl"
         End Select
      Case catReference
         lngType = GetTypeFromTag(LastNodeAdded.Tag)
         Select Case lngType
         Case vbext_RefKind.vbext_rk_TypeLib
            strFileName = GetFileNameFromTag(LastNodeAdded.Tag)
            strExtension = UCase(Right(strFileName, 3))
            If InStr(1, strExtension, "\", vbTextCompare) > 0 Then
               strExtension = Left(strFileName, Len(strFileName) - InStr(1, strExtension, "\", vbTextCompare))
               strExtension = UCase(Right(strExtension, 3))
            End If
            Select Case strExtension
            Case "DLL", "OLB"
               LastNodeAdded.Image = "Reference_DLL"
            Case "TLB"
               LastNodeAdded.Image = "Reference_TLB"
            Case "LIB"
               LastNodeAdded.Image = "Reference_LIB"
            Case "OCA", "OCX"
               LastNodeAdded.Image = "Reference_OCX"
            Case Else
               On Error Resume Next
               lngType = VBInstance.VBProjects(GetComponentNameFromTag(LastNodeAdded.Tag)).Type
               If Err.Number = 0 Then
                  Select Case lngType
                  Case vbext_pt_StandardExe
                     LastNodeAdded.Image = "Project_EXE"
                  Case vbext_pt_ActiveXDll
                     LastNodeAdded.Image = "Project_ActiveXDLL"
                  Case vbext_pt_ActiveXExe
                     LastNodeAdded.Image = "Project_ActiveXEXE"
                  Case vbext_pt_ActiveXControl
                     LastNodeAdded.Image = "Project_UserControl"
                  Case Else
                     LastNodeAdded.Image = "UnIdentified"
                  End Select
               Else
                  LastNodeAdded.Image = "UnIdentified"
               End If
            End Select
         Case vbext_RefKind.vbext_rk_Project
            lngType = VBInstance.VBProjects(GetComponentNameFromTag(LastNodeAdded.Tag)).Type
            Select Case lngType
            Case vbext_pt_StandardExe
               LastNodeAdded.Image = "Project_EXE"
            Case vbext_pt_ActiveXDll
               LastNodeAdded.Image = "Project_ActiveXDLL"
            Case vbext_pt_ActiveXExe
               LastNodeAdded.Image = "Project_ActiveXEXE"
            Case vbext_pt_ActiveXControl
               LastNodeAdded.Image = "Project_UserControl"
            Case Else
               LastNodeAdded.Image = "UnIdentified"
            End Select
         Case Else
            LastNodeAdded.Image = "UnIdentified"
         End Select
      Case catComponent
         lngType = GetTypeFromTag(LastNodeAdded.Tag)
         Select Case lngType
         Case vbext_ComponentType.vbext_ct_ActiveXDesigner
            LastNodeAdded.Image = "Component_Designer"
         Case vbext_ComponentType.vbext_ct_ClassModule
            LastNodeAdded.Image = "Component_ClassModule"
         Case vbext_ComponentType.vbext_ct_DocObject
            LastNodeAdded.Image = "Component_UserDocument"
         Case vbext_ComponentType.vbext_ct_MSForm
            LastNodeAdded.Image = "Component_Designer"
         Case vbext_ComponentType.vbext_ct_PropPage
            LastNodeAdded.Image = "Component_PropertyPage"
         Case vbext_ComponentType.vbext_ct_RelatedDocument
            LastNodeAdded.Image = "Component_Form"
         Case vbext_ComponentType.vbext_ct_ResFile
            LastNodeAdded.Image = "Component_Resource"
         Case vbext_ComponentType.vbext_ct_StdModule
            LastNodeAdded.Image = "Component_Module"
         Case vbext_ComponentType.vbext_ct_UserControl
            LastNodeAdded.Image = "Component_UserControl"
         Case vbext_ComponentType.vbext_ct_VBForm
            LastNodeAdded.Image = "Component_Form"
         Case vbext_ComponentType.vbext_ct_VBMDIForm
            LastNodeAdded.Image = "Component_MDIForm"
         End Select
      Case catControl
         LastNodeAdded.Image = "Component_UserControl"
      Case catProcedure
         LastNodeAdded.Image = "Procedure"
      Case catTitle
         Select Case LastNodeAdded.Text
         Case cstrReferences
            LastNodeAdded.Image = "Title_References"
         Case cstrComponents
            LastNodeAdded.Image = "Title_Components"
         End Select
      End Select
      DoEvents
   Set lstimg = Nothing
End Sub

Private Sub SetNodeColorsFont(ObjectType As enObjectType)
'Set the Font, Backcolor and Forecolor for the Last Added Node
On Error Resume Next
   Select Case ObjectType
   Case enObjectType.otProjectGroup
      LastNodeAdded.BackColor = RGB(100, 255, 100)
      LastNodeAdded.ForeColor = RGB(0, 0, 200)
      LastNodeAdded.Bold = False
   Case enObjectType.otProject
      LastNodeAdded.BackColor = RGB(200, 255, 200)
      LastNodeAdded.ForeColor = RGB(0, 0, 0)
      LastNodeAdded.Bold = False
   Case enObjectType.otReference
      If VBInstance.VBProjects(GetProjectNameFromTag(LastNodeAdded.Tag)).References(GetComponentNameFromTag(LastNodeAdded.Tag)).IsBroken Then
         LastNodeAdded.BackColor = RGB(255, 150, 150)
      Else
         LastNodeAdded.BackColor = RGB(255, 255, 255)
      End If
      If StrComp(Right(LastNodeAdded.Key, 3), "vbp", vbTextCompare) = 0 Then
         LastNodeAdded.ForeColor = RGB(0, 0, 0)
      Else
         LastNodeAdded.ForeColor = RGB(150, 150, 150)
      End If
      LastNodeAdded.Bold = False
   Case enObjectType.otComponent
      LastNodeAdded.BackColor = RGB(255, 255, 255)
      LastNodeAdded.ForeColor = RGB(0, 0, 0)
      LastNodeAdded.Bold = False
   Case enObjectType.otControl
      LastNodeAdded.BackColor = RGB(255, 255, 255)
      LastNodeAdded.ForeColor = RGB(0, 0, 0)
      LastNodeAdded.Bold = False
   Case enObjectType.otProcedure
      LastNodeAdded.BackColor = RGB(255, 255, 255)
      LastNodeAdded.ForeColor = RGB(0, 0, 0)
      LastNodeAdded.Bold = False
   Case enObjectType.otNone
      LastNodeAdded.BackColor = RGB(230, 255, 255)
      LastNodeAdded.ForeColor = RGB(0, 0, 0)
      LastNodeAdded.Bold = False
   End Select
   DoEvents
End Sub

Public Sub ActivateSelectedNodeItem(Node As MSComctlLib.Node)
On Error Resume Next
   Dim strCategory As String
   Dim uObjectType As enObjectType
   '
      'Reset the node color to it's original color.
      If Not CurrentSelectedNode Is Nothing Then
         Set LastNodeAdded = CurrentSelectedNode
         strCategory = GetCategoryFromTag(LastNodeAdded.Tag)
         uObjectType = CategoryStringToObjectType(strCategory)
         Call SetNodeColorsFont(uObjectType)
      End If
      '
      Set CurrentSelectedNode = Node
      strCategory = GetCategoryFromTag(Node.Tag)
      Select Case strCategory
      Case catProjectGroup
         'Do Nothing
      Case catProject
         'Do Nothing
      Case catReference
         'Do Nothing
      Case catComponent
         Call ActivateSelectedComponentNode(Node)
      Case catControl
         Call ActivateSelectedControlNode(Node)
      Case catProcedure
         Call ActivateAndSelectedProcedureNode(Node)
      Case catTitle
         'Do Nothing
      End Select
      Node.BackColor = RGB(230, 230, 230)
End Sub

Private Sub ActivateSelectedComponentNode(Node As MSComctlLib.Node)
'Opens the selected Project Component in it's design window
On Error GoTo ErrorHandler
   Dim strProject As String
   Dim strComponentName As String
   Dim lngNodeIndex As Long
   '
      strProject = GetProjectNameFromTag(Node.Tag)
      strComponentName = GetComponentNameFromTag(Node.Tag)
      VBInstance.VBProjects(strProject).VBComponents(strComponentName).Activate
ExitRoutine:
   Exit Sub
ErrorHandler:
   lngNodeIndex = Node.Index
   Call PopulateNode(Node.Parent)
End Sub

Private Sub ActivateSelectedControlNode(Node As MSComctlLib.Node)
'Opens the selected Controls Form in it's design window and selects the control on the form.
   Dim strProject As String
   Dim strComponentName As String
   Dim strControlName As String
   Dim strControlIndex As String
   Dim Activeform As VBForm
   Dim blFound As Boolean
   Dim i As Long
   Dim lngNodeIndex As Long
   '
      strProject = GetProjectNameFromTag(Node.Tag)
      strComponentName = GetComponentNameFromTag(Node.Tag)
      Call GetControlName(Node.Text, strControlName, strControlIndex)
      VBInstance.VBProjects(strProject).VBComponents(strComponentName).Activate
      Set Activeform = VBInstance.VBProjects(strProject).VBComponents(strComponentName).Designer
      If strControlIndex = "" Then
         For i = 1 To Activeform.VBControls.Count
            If Activeform.VBControls.Item(i).Properties("Name") = strControlName Then
               Activeform.SelectedVBControls.Clear
               Activeform.VBControls.Item(i).InSelection = True
               blFound = True
               Exit For
            End If
         Next
      Else
         For i = 1 To Activeform.VBControls.Count
            If Activeform.VBControls.Item(i).Properties("Name") = strControlName And Activeform.VBControls.Item(i).Properties("Index") = strControlIndex Then
               Activeform.SelectedVBControls.Clear
               Activeform.VBControls.Item(i).InSelection = True
               blFound = True
               Exit For
            End If
         Next
      End If
      'If the item represented by the node is not found it has probably been deleted or renamed so refresh the nodes parent.
      If Not blFound Then
         lngNodeIndex = Node.Index
         Call PopulateNode(Node.Parent)
      End If
      Set Activeform = Nothing
End Sub

Private Sub ActivateAndSelectedProcedureNode(Node As MSComctlLib.Node)
'Opens the selected Member Item's Code Module in it's design window and navigates to the Member Item.
   Dim strProject As String
   Dim strComponentName As String
   Dim lngCodeLocation As Long
   Dim strMemberName As String
   Dim strPropertyKind As String
   Dim cpActiveCodePane As CodePane
   '
      strProject = GetProjectNameFromTag(Node.Tag)
      strComponentName = GetComponentNameFromTag(Node.Tag)
      strMemberName = GetMemberName(Node.Text)
      strPropertyKind = GetPropertyKind(Node.Text)
      VBInstance.VBProjects(strProject).VBComponents(strComponentName).CodeModule.CodePane.Show
      Set cpActiveCodePane = VBInstance.VBProjects(strProject).VBComponents(strComponentName).CodeModule.CodePane
      '
      'No longer use this method as codeLoation stored in tag could be out of date
'      lngCodeLocation = GetCodeLocationFromTag(Node.Tag)
      lngCodeLocation = GetCodeLocation(strProject, strComponentName, strMemberName, strPropertyKind)
      Call cpActiveCodePane.SetSelection(lngCodeLocation, 1, lngCodeLocation, 1)
      cpActiveCodePane.TopLine = lngCodeLocation
      '
      'This line is actually not needed, I am just trying to set the focus to the window so that you an se the cursor but it doesn't work.
      'VBInstance.VBProjects(strProject).VBComponents(strComponentName).CodeModule.CodePane.Window.SetFocus
      '
      'This is the only way I can get this to work is using the following which to me just seems ugly because the window flashes.
'      cpActiveCodePane.Window.Visible = False
'      cpActiveCodePane.Window.Visible = True
'      VBInstance.VBProjects(strProject).VBComponents(strComponentName).CodeModule.CodePane.Window.SetFocus
      '
      Set cpActiveCodePane = Nothing
End Sub

Private Function GetCategoryFromTag(TagText As String) As String
   Dim strValue As String
   Dim lngStartPosition As Long
   Dim lngEndPosition As Long
   '
      lngStartPosition = InStr(1, TagText, "Category=", vbTextCompare)
      If lngStartPosition > 0 Then
         lngStartPosition = InStr(1, TagText, "Category=", vbTextCompare) + 9
         lngEndPosition = InStr(lngStartPosition, TagText, "|", vbTextCompare)
         strValue = Mid(TagText, lngStartPosition, lngEndPosition - lngStartPosition)
      Else
         strValue = -1
      End If
      GetCategoryFromTag = strValue
End Function

Private Function GetProjectNameFromTag(TagText As String) As String
   Dim strValue As String
   Dim lngStartPosition As Long
   Dim lngEndPosition As Long
   '
      lngStartPosition = InStr(1, TagText, "|Project=", vbTextCompare)
      If lngStartPosition > 0 Then
         lngStartPosition = InStr(1, TagText, "|Project=", vbTextCompare) + 9
         lngEndPosition = InStr(lngStartPosition, TagText, "|", vbTextCompare)
         strValue = Mid(TagText, lngStartPosition, lngEndPosition - lngStartPosition)
      Else
         strValue = -1
      End If
      GetProjectNameFromTag = strValue
End Function

Private Function GetComponentNameFromTag(TagText As String) As String
   Dim strValue As String
   Dim lngStartPosition As Long
   Dim lngEndPosition As Long
   '
      lngStartPosition = InStr(1, TagText, "|Component=", vbTextCompare)
      If lngStartPosition > 0 Then
         lngStartPosition = InStr(1, TagText, "|Component=", vbTextCompare) + 11
         lngEndPosition = InStr(lngStartPosition, TagText, "|", vbTextCompare)
         strValue = Mid(TagText, lngStartPosition, lngEndPosition - lngStartPosition)
      Else
         strValue = -1
      End If
      GetComponentNameFromTag = strValue
End Function

Private Function GetFileNameFromTag(TagText As String) As String
   Dim strValue As String
   Dim lngStartPosition As Long
   Dim lngEndPosition As Long
   '
      lngStartPosition = InStr(1, TagText, "|FileName=", vbTextCompare)
      If lngStartPosition > 0 Then
         lngStartPosition = InStr(1, TagText, "|FileName=", vbTextCompare) + 10
         lngEndPosition = InStr(lngStartPosition, TagText, "|", vbTextCompare)
         strValue = Mid(TagText, lngStartPosition, lngEndPosition - lngStartPosition)
      Else
         strValue = -1
      End If
      GetFileNameFromTag = strValue
End Function

Private Function GetTypeFromTag(TagText As String) As Long
   Dim lngValue As Long
   Dim lngStartPosition As Long
   Dim lngEndPosition As Long
   '
      lngStartPosition = InStr(1, TagText, "|Type=", vbTextCompare)
      If lngStartPosition > 0 Then
         lngStartPosition = InStr(1, TagText, "|Type=", vbTextCompare) + 6
         lngEndPosition = InStr(lngStartPosition, TagText, "|", vbTextCompare)
         lngValue = Mid(TagText, lngStartPosition, lngEndPosition - lngStartPosition)
      Else
         lngValue = -1
      End If
      GetTypeFromTag = lngValue
End Function

Private Function GetCodeLocationFromTag(TagText As String) As String
   Dim strValue As String
   Dim lngStartPosition As Long
   Dim lngEndPosition As Long
   '
      lngStartPosition = InStr(1, TagText, "|CodeLocation=", vbTextCompare)
      If lngStartPosition > 0 Then
         lngStartPosition = InStr(1, TagText, "|CodeLocation=", vbTextCompare) + 14
         lngEndPosition = InStr(lngStartPosition, TagText, "|", vbTextCompare)
         strValue = Mid(TagText, lngStartPosition, lngEndPosition - lngStartPosition)
      Else
         strValue = -1
      End If
      GetCodeLocationFromTag = strValue
End Function

Private Sub GetControlName(ByVal NodeText As String, outControlName As String, Optional outControlIndex As String)
   Dim strControlName As String
   Dim strControlIndex As String
   '
      If InStr(1, NodeText, "(", vbTextCompare) > 0 Then
         strControlIndex = Right(NodeText, Len(NodeText) - InStr(1, NodeText, "(", vbTextCompare))
         strControlIndex = Left(strControlIndex, InStr(1, strControlIndex, ")", vbTextCompare) - 1)
         strControlName = Left(NodeText, InStr(1, NodeText, "(", vbTextCompare) - 1)
      Else
         strControlName = NodeText
      End If
      strControlName = Right(strControlName, Len(strControlName) - 9)
      '
      outControlName = strControlName
      outControlIndex = strControlIndex
End Sub

Private Function GetMemberName(ByVal NodeText As String) As String
   Dim strMemberName As String
   '
      If InStr(1, NodeText, "(", vbTextCompare) > 0 Then
         strMemberName = Left(NodeText, InStr(1, NodeText, "(", vbTextCompare) - 1)
      Else
         strMemberName = NodeText
      End If
      strMemberName = Right(strMemberName, Len(strMemberName) - InStr(1, strMemberName, ":", vbTextCompare))
      GetMemberName = Trim(strMemberName)
End Function

Private Function ScopeAsString(Scope As vbext_Scope) As String
'Used to convert the Enumeration into a String
   Select Case Scope
   Case vbext_Scope.vbext_Friend
      ScopeAsString = "Friend"
   Case vbext_Scope.vbext_Private
      ScopeAsString = "Private"
   Case vbext_Scope.vbext_Public
      ScopeAsString = "Public"
   End Select
End Function

Private Function MemberTypeAsString(MemberType As vbext_MemberType) As String
'Used to convert the Enumeration into a String
   Select Case MemberType
   Case vbext_MemberType.vbext_mt_Const
      MemberTypeAsString = "Const"
   Case vbext_MemberType.vbext_mt_Event
      MemberTypeAsString = "Event"
   Case vbext_MemberType.vbext_mt_Method
      MemberTypeAsString = "Method"
   Case vbext_MemberType.vbext_mt_Property
      MemberTypeAsString = "Property"
   Case vbext_MemberType.vbext_mt_Variable
      MemberTypeAsString = "Variable"
   End Select
End Function

Private Function CategoryStringToObjectType(CategoryTypeString As String) As enObjectType
   Select Case CategoryTypeString
   Case catComponent
      CategoryStringToObjectType = otComponent
   Case catControl
      CategoryStringToObjectType = otControl
   Case catProcedure
      CategoryStringToObjectType = otProcedure
   Case catProject
      CategoryStringToObjectType = otProject
   Case catProjectGroup
      CategoryStringToObjectType = otProjectGroup
   Case catReference
      CategoryStringToObjectType = otReference
   Case catTitle
      CategoryStringToObjectType = otNone
   Case Else
      CategoryStringToObjectType = otNone
   End Select
End Function

Private Function StripPathFromFile(FilePath As String) As String
'Strips the path from the file and return the the Filename only
   Dim strFile As String
   '
      strFile = FilePath
      Do While InStr(1, strFile, "\", vbTextCompare)
         strFile = Right(strFile, Len(strFile) - InStr(1, strFile, "\", vbTextCompare))
      Loop
      StripPathFromFile = strFile
End Function

Private Function StripFileFromPath(FilePath As String) As String
'Strips the file from the path and return the the Path only
   Dim strFile As String
   Dim lngLastPos As Long
   Dim lngPos As Long
   '
      strFile = FilePath
      Do
         lngLastPos = lngPos
         lngPos = InStr(lngPos + 1, strFile, "\", vbTextCompare)
      Loop Until lngPos = 0
      StripFileFromPath = Left(strFile, lngLastPos)
End Function

Private Function GetCodeLocation(ProjectName As String, ComponentName As String, MemberName As String, PropertyKind As String) As Long
On Error Resume Next
Dim lngCodeLocation As Long
Dim lngIndex As Long
'
   lngCodeLocation = 0
   lngIndex = lngIndex + 1
   lngCodeLocation = VBInstance.VBProjects(ProjectName).VBComponents(ComponentName).CodeModule.ProcBodyLine(MemberName, vbext_pk_Proc)
   If lngCodeLocation <> 0 Then
      GetCodeLocation = lngCodeLocation
      Exit Function
   End If
   '
   Select Case PropertyKind
   Case "Let"
      lngCodeLocation = VBInstance.VBProjects(ProjectName).VBComponents(ComponentName).CodeModule.ProcBodyLine(MemberName, vbext_pk_Let)
      If lngCodeLocation <> 0 Then
         GetCodeLocation = lngCodeLocation
         Exit Function
      End If
   Case "Set"
      lngCodeLocation = VBInstance.VBProjects(ProjectName).VBComponents(ComponentName).CodeModule.ProcBodyLine(MemberName, vbext_pk_Set)
      If lngCodeLocation <> 0 Then
         GetCodeLocation = lngCodeLocation
         Exit Function
      End If
   Case "Get"
      lngCodeLocation = VBInstance.VBProjects(ProjectName).VBComponents(ComponentName).CodeModule.ProcBodyLine(MemberName, vbext_pk_Get)
      If lngCodeLocation <> 0 Then
         GetCodeLocation = lngCodeLocation
         Exit Function
      End If
   End Select
   'All of the above methods failed so we use the .CodeLocation property
   'Code location property is the last resort as it does not take line continuation characters (_) into consideration
   'but it also seems to be the only way of returning line numbers for variable declarations.
   lngCodeLocation = VBInstance.VBProjects(ProjectName).VBComponents(ComponentName).CodeModule.Members(MemberName).CodeLocation
   If lngCodeLocation <> 0 Then
      GetCodeLocation = lngCodeLocation
      Exit Function
   End If
End Function

Private Function CodeLocationFound(LineNumber As Long) As Boolean

End Function

Private Function GetPropertyKind(NodeText As String) As String
   Dim strPropertyKind As String
   '
   If Left(NodeText, 8) = "Property" Then
      strPropertyKind = Left(NodeText, InStr(1, NodeText, ":", vbTextCompare) - 1)
      strPropertyKind = Right(strPropertyKind, Len(strPropertyKind) - 9)
      GetPropertyKind = Trim(strPropertyKind)
   Else
      GetPropertyKind = ""
   End If
End Function


'-------------------------------------------------------------------------------------------------------------------
'*******************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------------
'The following are possible future features that could be added. All these predefined ones will do is remove the need to manually update/Refresh the tree. I chose not to implement these for 2 reasons:
'1
'When an item is added particularly a component or project they have not yet been named or saved and I just felt why waste resource time etc on something that is probably going to change soon.
'2
'Sometimes when you are coding and creating procedures you change the name of an item a few times before you decide on a name that satifies you and your naming convention/criteria.
'
'All that this means is a little extra effort on the developers side in refreshing the tree.
'
'Other potential additions include the ability to add item to the tree by having a PopUp menu for the tree.
'Items in this menu would include the ability to:
'Add/Remove Projects
'Add/Remove References
'Add/Remove Components (Forms, Modules, Designers, etc)
'Add/Remove Controls (Maybe)
'Add/Remove Code (Define a procedure)
'Again these have not been done yet because the effort relative to gain does not prompt me to do this.
'These things are also currently handled by the native Project Explorer, Toolbox, and Add Procedure menu option,
'I also know of no developer who uses the add procedure dialog box so why add this feature to my tree.
'
'PS: If anyone does add these features I would love to have a copy of the source code.
'e-mail: darrylha@dazzlingsoftware.com
'Website: http://www.dazzlingsoftware.com
'-------------------------------------------------------------------------------------------------------------------
'*******************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------------
'
'Public Sub ProjectActivated(Proj As VBProject)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ProjectAdded(Proj As VBProject)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ProjectRemoved(Proj As VBProject)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ProjectRenamed(Proj As VBProject)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ReferenceAdded(Ref As Reference)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ReferenceRemoved(Ref As Reference)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ComponentActivated(Comp As VBComponent)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
''Dim ProjectNode As MSComctlLib.Node
''   Tree.Nodes(Comp.Collection.Parent.FileName).Expanded = True
''   Tree.Nodes(Comp.Collection.Parent.FileName & "|" & cstrComponents).Expanded = True
''   Tree.Nodes(Comp.FileNames(1)).Selected = True
''   Call Tree_NodeClick(Tree.Nodes(Comp.FileNames(1)))
'End Sub
'
'Public Sub ComponentAdded(Comp As VBComponent)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ComponentRemoved(Comp As VBComponent)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ComponentRenamed(Comp As VBComponent)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ControlAdded(NewCtl As VBControl)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ControlRenamed(Ctrl As VBControl, OldName As String)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub ControlRemoved(Ctrl As VBControl)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
'
'Public Sub FileAdded(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
''Reserved for potential future feature.
''Currently I have not written any code to perform this action as the effort relative to gain does not prompt me to do this.
''
''This will automatically Update the tree.
'End Sub
