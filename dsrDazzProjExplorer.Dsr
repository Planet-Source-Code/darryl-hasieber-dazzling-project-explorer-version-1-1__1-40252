VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dsrDazzProjExplorer 
   ClientHeight    =   13365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14640
   _ExtentX        =   25823
   _ExtentY        =   23574
   _Version        =   393216
   DisplayName     =   "Dazzling Project Explorer"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   SatName         =   "DazzProjExplorer.dll"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "dsrDazzProjExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Const guidMYTOOL$ = "DazzlingSoftwareVBProjectExplorer"
'
Public WithEvents PrjHandler As VBIDE.VBProjectsEvents            'projects event handler
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents RefHandler As VBIDE.ReferencesEvents            'reference event handler
Attribute RefHandler.VB_VarHelpID = -1
Public WithEvents CmpHandler As VBIDE.VBComponentsEvents          'components event handler
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents CtlHandler As VBIDE.VBControlsEvents            'controls event handler
Attribute CtlHandler.VB_VarHelpID = -1
Public WithEvents SelCtlHandler As VBIDE.SelectedVBControlsEvents 'selected controls event handler
Attribute SelCtlHandler.VB_VarHelpID = -1
Public WithEvents FileHandler As VBIDE.FileControlEvents          'file event handler
Attribute FileHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler As VBIDE.CommandBarEvents           'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
'Private mcbMenuCommandBar As Office.CommandBarControl            'command bar object
Private mcbMenuCommandBar As Object                               'command bar object - Switch to this method to fix a deployment problem
'

Public Property Get NonModalApp() As Boolean
   'Predefined code from VB i.e. I did not write this
   NonModalApp = True  'used by addin toolbar
End Property

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
'This occurs when the addin is loaded in an instance of VB
On Error GoTo ErrorHandler
   Dim aiTmp As AddIn
   '
      Set gVBInstance = Application
      Set TreeFunctions = New clsTree
      Set TreeFunctions.VBIDEInstance = gVBInstance
      '
      'If AddIn is Running
      If Not gwinWindow Is Nothing Then
         Show
         If ConnectMode = ext_cm_AfterStartup Then
            'Started from AddIn Manager
             AddToCommandBar
         End If
         Exit Sub
      End If
      '
      'AddIn is Not Running
      'create the tool window
      If ConnectMode = ext_cm_External Then
         'Test if already running
         On Error Resume Next
         Set aiTmp = gVBInstance.Addins("TabOrder.Connect")
         On Error GoTo ErrorHandler
         If aiTmp Is Nothing Then
            'app is not in the VBADDIN.INI file so it is not in the collection so lets attempt to use the 1st addin in the
            'collection just to get this app running and if there are none, an error will occur and this app will not run
             Set gwinWindow = gVBInstance.Windows.CreateToolWindow(gVBInstance.Addins(1), "DazzProjExplorer.docDazzProjExplorer", "Dazzling VB Project Explorer", guidMYTOOL$, gdocDazzExplorer)
         Else
            If aiTmp.Connect = False Then
               Set gwinWindow = gVBInstance.Windows.CreateToolWindow(aiTmp, "DazzProjExplorer.docDazzProjExplorer", "Dazzling VB Project Explorer", guidMYTOOL$, gdocDazzExplorer)
            End If
         End If
      Else
         'Called from AddIn Manager
         Set gwinWindow = gVBInstance.Windows.CreateToolWindow(AddInInst, "DazzProjExplorer.docDazzProjExplorer", "Dazzling VB Project Explorer", guidMYTOOL$, gdocDazzExplorer)
      End If
      '
      'synchronise the project, components and controls event handlers
      Set Me.PrjHandler = gVBInstance.Events.VBProjectsEvents
      Set Me.CmpHandler = gVBInstance.Events.VBComponentsEvents(Nothing)
      Set Me.CtlHandler = gVBInstance.Events.VBControlsEvents(Nothing, Nothing)
      '
      'If started from AddIn toolbar then else if started from AddIn Manager
      If ConnectMode = vbext_cm_External Then
         Show
      ElseIf ConnectMode = vbext_cm_AfterStartup Then
         AddToCommandBar
      End If
      '
ExitRoutine:
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
   Resume
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
'This occurs when the addin is unloaded from an instance of VB
On Error GoTo ErrorHandler
'
   mcbMenuCommandBar.Delete
   SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", gwinWindow.Visible
   '
ExitRoutine:
   Set gwinWindow = Nothing
   Set TreeFunctions = Nothing
   Set gVBInstance = Nothing
   Set PrjHandler = Nothing
   Set RefHandler = Nothing
   Set CmpHandler = Nothing
   Set CtlHandler = Nothing
   Set SelCtlHandler = Nothing
   Set FileHandler = Nothing
   Set MenuHandler = Nothing
   Set mcbMenuCommandBar = Nothing
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
'When IDE is fully loaded
   AddToCommandBar
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'this event fires when the command bar control is clicked in the IDE
   Show
End Sub

Sub AddToCommandBar()
'Add my AddIn menu button to the VB toolbar
On Error GoTo ErrorHandler
'
   gVBInstance.CommandBars(2).Visible = True
   
   'Add to VB IDE Toolbar
   Set mcbMenuCommandBar = gVBInstance.CommandBars(2).Controls.Add(1, , , gVBInstance.CommandBars(2).Controls.Count)
   mcbMenuCommandBar.Caption = "Dazzling Project Explorer"
   Clipboard.SetData LoadResPicture("Tree", 0)
   mcbMenuCommandBar.PasteFace
   '
   'Sync the Menu Event handler
   Set Me.MenuHandler = gVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
   '
   'restore the last state
   If GetSetting(APP_CATEGORY, App.Title, "DisplayOnConnect", "0") Then
      Me.Show
   End If
   '
ExitRoutine:
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
End Sub

Sub Show()
On Error GoTo ErrorHandler
   '
   gwinWindow.Visible = True
   gdocDazzExplorer.BuildTree
   '
ExitRoutine:
   Exit Sub
ErrorHandler:
   MsgBox Err.Description
End Sub

'Private Sub PrjHandler_ItemActivated(ByVal VBProject As VBIDE.VBProject)
''When Project is Activated
'   If gwinWindow.Visible Then
'      TreeFunctions.ProjectActivated VBProject
'   End If
'End Sub
'
'Private Sub PrjHandler_ItemAdded(ByVal VBProject As VBIDE.VBProject)
''When Project is Added
'   If gwinWindow.Visible Then
'      TreeFunctions.ProjectAdded VBProject
'   End If
'End Sub
'
'Private Sub PrjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
''When Project is Removed
'  If gwinWindow.Visible Then
'      TreeFunctions.ProjectRemoved VBProject
'  End If
'End Sub
'
'Private Sub PrjHandler_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
''When Project is Renamed
'   If gwinWindow.Visible Then
'      TreeFunctions.ProjectRenamed VBProject
'   End If
'End Sub
'
'Private Sub RefHandler_ItemAdded(ByVal Reference As VBIDE.Reference)
''When Reference is Added
'   If gwinWindow.Visible Then
'      TreeFunctions.ReferenceAdded Reference
'   End If
'End Sub
'
'Private Sub RefHandler_ItemRemoved(ByVal Reference As VBIDE.Reference)
''When Reference is Removed
'   If gwinWindow.Visible Then
'      TreeFunctions.ReferenceRemoved Reference
'   End If
'End Sub
'
'Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
''When Component is Activated
'   If gwinWindow.Visible Then
'      TreeFunctions.ComponentActivated VBComponent
'   End If
'End Sub
'
'Private Sub CmpHandler_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
''When Component is Added
'   If gwinWindow.Visible Then
'      TreeFunctions.ComponentAdded VBComponent
'   End If
'End Sub
'
'Private Sub CmpHandler_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
''When Component is Removed
'   If gwinWindow.Visible Then
'      TreeFunctions.ComponentRemoved VBComponent
'   End If
'End Sub
'
'Private Sub CmpHandler_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
''When Component is Renamed
'   If gwinWindow.Visible Then
'      TreeFunctions.ComponentRenamed VBComponent
'   End If
'End Sub
'
'Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
''When Component is Selected
'   CmpHandler_ItemActivated VBComponent
'End Sub
'
'Private Sub CtlHandler_ItemAdded(ByVal VBControl As VBIDE.VBControl)
''When a Control is Added
'   If gwinWindow.Visible Then
'      TreeFunctions.ControlAdded VBControl
'   End If
'End Sub
'
'Private Sub CtlHandler_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
''When a Control is Renamed
'   If gwinWindow.Visible Then
'      TreeFunctions.ControlRenamed VBControl, OldName
'   End If
'End Sub
'
'Private Sub CtlHandler_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
''When a Control is Removed
'   If gwinWindow.Visible Then
'      TreeFunctions.ControlRemoved VBControl
'   End If
'End Sub
'
'Private Sub SelCtlHandler_ItemAdded(ByVal VBControl As VBIDE.VBControl)
''When a Control is Removed
'   If gwinWindow.Visible Then
'      TreeFunctions.ControlRemoved VBControl
'   End If
'End Sub
'
'Private Sub SelCtlHandler_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
''When a Control is Removed
'   If gwinWindow.Visible Then
'      TreeFunctions.ControlRemoved VBControl
'   End If
'End Sub
'
'Private Sub FileHandler_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
''When a Control is Removed
'   If gwinWindow.Visible Then
'      TreeFunctions.FileAdded VBProject, FileType, FileName
'   End If
'End Sub
