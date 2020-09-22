VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private strMyProp As String
Dim strMyVar As String
'

Public Property Let MyProp(Val As String)
   strMyProp = Val
End Property


Public Function EnterClass()
   'This is the entry point
End Function

Private Sub DoSomeStuff()
   'Do some stuff damnit.... Jannit
End Sub
