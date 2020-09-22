VERSION 5.00
Object = "*\AFour.vbp"
Begin VB.Form FormOne 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2475
      Left            =   1140
      TabIndex        =   7
      Top             =   2280
      Width           =   3195
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1755
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2955
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1560
            TabIndex        =   14
            Text            =   "Combo1"
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   13
            Top             =   780
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   375
            Index           =   1
            Left            =   180
            TabIndex        =   12
            Top             =   1140
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   11
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   315
            Left            =   1560
            TabIndex        =   15
            Top             =   540
            Width           =   1215
         End
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   315
         Left            =   660
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   1500
      TabIndex        =   3
      Top             =   900
      Width           =   1575
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   540
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   780
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1275
      End
   End
   Begin ProjectFour.UserControl4_1 UserControl4_11 
      Height          =   795
      Left            =   3180
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
      _extentx        =   2355
      _extenty        =   1402
   End
   Begin ProjectOne.UserControl1_1 UserControl1_11 
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   900
      Width           =   1335
      _extentx        =   2355
      _extenty        =   767
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3300
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   2760
      Top             =   120
   End
End
Attribute VB_Name = "FormOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'test
End Sub

Private Sub Form_Load()
   'Form Load Event
End Sub

Private Sub Timer1_Timer()
   'do something
End Sub

Private Function ReturnFormName() As String

End Function
