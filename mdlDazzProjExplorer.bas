Attribute VB_Name = "mdlDazzProjExplorer"
Option Explicit
'
Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"
'
Const WM_SYSKEYDOWN = &H104
Const WM_SYSKEYUP = &H105
Const WM_SYSCHAR = &H106
Const VK_F = 70                        ' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
'
Global TreeFunctions As clsTree
Global gVBInstance As VBIDE.VBE        'instance of VB IDE
Global gwinWindow As VBIDE.Window      'used to make sure we only run one instance
Global gdocDazzExplorer As Object      'user doc object
'
Private hwndMenu As Long               'needed to pass the menu keystrokes to VB
'
Private Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal msg&, ByVal wp&, ByVal lp&) 'The PostMessage function places (posts) a message in the message queue associated with the thread that created the specified window and then returns without waiting for the thread to process the message.
Private Declare Sub SetFocus Lib "user32" (ByVal hwnd&)
Private Declare Function GetParent Lib "user32" (ByVal hwnd&) As Long   'Retrieves the handle of the specified child window’s parent window.
'

Function InRunMode(VBInst As VBIDE.VBE) As Boolean
   'This enables/disables my AddIn menu button
   InRunMode = Not gVBInstance.CommandBars("File").Controls(1).Enabled
End Function

Sub HandleKeyDown(ud As Object, KeyCode As Integer, Shift As Integer)
   If Shift <> 4 Then Exit Sub
   If KeyCode < 65 Or KeyCode > 90 Then Exit Sub
   If gVBInstance.DisplayModel = vbext_dm_SDI Then Exit Sub
   
   If hwndMenu = 0 Then hwndMenu = FindHwndMenu(ud.hwnd)
   Call PostMessage(hwndMenu, WM_SYSKEYDOWN, KeyCode, &H20000000)
   KeyCode = 0
   SetFocus hwndMenu
End Sub

Function FindHwndMenu(ByVal hwnd As Long) As Long
   Dim h As Long
      Do
         h = GetParent(hwnd)
         If h = 0 Then
            FindHwndMenu = hwnd
            Exit Function
         End If
         hwnd = h
      Loop
End Function


