VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   555
      TabIndex        =   1
      Top             =   2400
      Width           =   1230
   End
   Begin Project1.BarChart BC1 
      Height          =   1620
      Left            =   390
      TabIndex        =   0
      Top             =   480
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   2858
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
BC1.Value1 = 10
BC1.Value2 = 20
BC1.Value3 = 30
BC1.Value4 = 40
BC1.Value5 = 50
BC1.Value6 = 60
End Sub
