VERSION 5.00
Begin VB.UserControl BarChart 
   BackColor       =   &H8000000A&
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   FillColor       =   &H8000000A&
   ScaleHeight     =   1710
   ScaleWidth      =   3240
   ToolboxBitmap   =   "BarChart.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   15
      ScaleHeight     =   1635
      ScaleWidth      =   3180
      TabIndex        =   0
      Top             =   15
      Width           =   3210
      Begin VB.Label lab1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   2655
         TabIndex        =   9
         Top             =   1290
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "75"
         Height          =   165
         Left            =   -15
         TabIndex        =   8
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "25"
         Height          =   180
         Left            =   -15
         TabIndex        =   7
         Top             =   705
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   165
         Left            =   -15
         TabIndex        =   6
         Top             =   1095
         Width           =   315
      End
      Begin VB.Label lab1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   240
         Index           =   4
         Left            =   2175
         TabIndex        =   5
         Top             =   1275
         Width           =   360
      End
      Begin VB.Label lab1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lab1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   240
         Index           =   2
         Left            =   1110
         TabIndex        =   3
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label lab1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   2
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label lab1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   120
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   1305
         Width           =   360
      End
      Begin VB.Line Line4 
         X1              =   225
         X2              =   3165
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Line Line3 
         X1              =   225
         X2              =   3210
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   3225
         Y1              =   1185
         Y2              =   1185
      End
      Begin VB.Line Line1 
         X1              =   225
         X2              =   225
         Y1              =   -15
         Y2              =   1575
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   0
         Left            =   315
         Top             =   1005
         Width           =   300
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   1
         Left            =   750
         Top             =   990
         Width           =   330
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   2
         Left            =   1245
         Top             =   990
         Width           =   330
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   3
         Left            =   1785
         Top             =   990
         Width           =   285
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   4
         Left            =   2295
         Top             =   960
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   5
         Left            =   2760
         Top             =   960
         Width           =   300
      End
   End
End
Attribute VB_Name = "BarChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************
'*  Simple and Quick Bar Chart
'*  By Ken Foster
'*   2004
'*  Free to use or change to
'*  Suit your needs
'*****************************

Const m_def_Value1 = 0
Const m_def_Value2 = 0
Const m_def_Value3 = 0
Const m_def_Value4 = 0
Const m_def_Value5 = 0
Const m_def_Value6 = 0
Const m_def_ScaleMax = 100
Const m_def_NOB = 6
Const m_def_ScaleVisible = True

Dim m_ScaleMax As Long
Dim m_ScaleVisible As Boolean
Dim m_Value1 As Long
Dim m_Value2 As Long
Dim m_Value3 As Long
Dim m_Value4 As Long
Dim m_Value5 As Long
Dim m_Value6 As Long

Dim m_NOB As Long

Private Sub Picture1_Click()

End Sub

Private Sub UserControl_Initialize()
Dim x As Long
Dim y As Long

   Shape1(0).Top = 1600
   Shape1(0).Height = 180
   Shape1(0).Left = 375
   Shape1(0).Width = 495

   lab1(0).Top = 1320
   lab1(0).Left = 420
   lab1(0).Height = 180
   lab1(0).Width = 360

   x = Shape1(0).Width + 70

   For y = 1 To 5
     Shape1(y).Top = 1600
     lab1(y).Top = 1320
     Shape1(y).Height = 180
     lab1(y).Height = 180
     Shape1(y).Left = Shape1(0).Left + (x * y)
     lab1(y).Left = lab1(0).Left + (x * y)
     Shape1(y).Width = 495
     lab1(1).Width = 360
   Next y

   Picture1.Top = 0
   Picture1.Left = 0
   Picture1.Height = 1600
   Picture1.Width = 3750


   Line1.X1 = 285
   Line1.Y2 = 1665
   Line1.X2 = 285
   Line1.Y2 = Picture1.Height

   Line2.X1 = 285
   Line2.Y1 = Picture1.Height / 2
   Line2.X2 = Picture1.Width
   Line2.Y2 = Picture1.Height / 2

   Line3.X1 = 285
   Line3.Y1 = Picture1.Height / 1.33
   Line3.X2 = Picture1.Width
   Line3.Y2 = Picture1.Height / 1.33

   Line4.X1 = 285
   Line4.Y1 = Picture1.Height / 4
   Line4.X2 = Picture1.Width
   Line4.Y2 = Picture1.Height / 4
   ScaleMax = m_ScaleMax
   NOB = 6
   ScaleMax = 100
   ScaleVisible = True
End Sub
Public Property Get NOB() As Long
     NOB = m_NOB
End Property
Public Property Let NOB(ByVal New_NOB As Long)
     m_NOB = New_NOB
     PropertyChanged "NOB"
     NumOf_Bars
End Property

Public Property Get ScaleMax() As Long
     ScaleMax = m_ScaleMax
End Property
Public Property Let ScaleMax(ByVal New_ScaleMax As Long)
     m_ScaleMax = New_ScaleMax
     PropertyChanged "ScaleMax"
      
       Label3.Caption = (m_ScaleMax / 4) * 3
       Label1.Caption = m_ScaleMax / 4
       Label2.Caption = m_ScaleMax / 2
     
    Call Bar_Value(Value1, 1)
    Call Bar_Value(Value2, 2)
    Call Bar_Value(Value3, 3)
    Call Bar_Value(Value4, 4)
    Call Bar_Value(Value5, 5)
    Call Bar_Value(Value6, 6)
    
End Property
Public Property Get ScaleVisible() As Boolean
    ScaleVisible = m_ScaleVisible
End Property
Public Property Let ScaleVisible(New_ScaleVisible As Boolean)
    m_ScaleVisible = New_ScaleVisible
    PropertyChanged "ScaleVisible"
    Call Scale_Visible
    
End Property
Public Property Get Value1() As Long
     Value1 = m_Value1
End Property
Public Property Let Value1(ByVal New_Value1 As Long)
     m_Value1 = New_Value1
     PropertyChanged "Value1"
     Bar_Value Value1, 1
End Property

Public Property Get Value2() As Long
     Value2 = m_Value2
End Property
Public Property Let Value2(ByVal New_Value2 As Long)
     m_Value2 = New_Value2
     PropertyChanged "Value2"
     Bar_Value Value2, 2
End Property
Public Property Get Value3() As Long
     Value3 = m_Value3
End Property
Public Property Let Value3(ByVal New_Value3 As Long)
     m_Value3 = New_Value3
     PropertyChanged "Value3"
     Bar_Value Value3, 3
End Property
Public Property Get Value4() As Long
     Value4 = m_Value4
End Property
Public Property Let Value4(ByVal New_Value4 As Long)
     m_Value4 = New_Value4
     PropertyChanged "Value4"
     Bar_Value Value4, 4
End Property
Public Property Get Value5() As Long
     Value5 = m_Value5
End Property
Public Property Let Value5(ByVal New_Value5 As Long)
     m_Value5 = New_Value5
     PropertyChanged "Value5"
     Bar_Value Value5, 5
End Property
Public Property Get Value6() As Long
     Value6 = m_Value6
End Property
Public Property Let Value6(ByVal New_Value6 As Long)
     m_Value6 = New_Value6
     PropertyChanged "Value6"
     Bar_Value Value6, 6
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ScaleMax = PropBag.ReadProperty("ScaleMax", m_def_ScaleMax)
    m_ScaleVisible = PropBag.ReadProperty("ScaleVisible", m_def_ScaleVisible)
    m_Value1 = PropBag.ReadProperty("Value1", m_def_Value1)
    m_Value2 = PropBag.ReadProperty("Value2", m_def_Value2)
    m_Value3 = PropBag.ReadProperty("Value3", m_def_Value3)
    m_Value4 = PropBag.ReadProperty("Value4", m_def_Value4)
    m_Value5 = PropBag.ReadProperty("Value5", m_def_Value5)
    m_Value6 = PropBag.ReadProperty("Value6", m_def_Value6)
    m_NOB = PropBag.ReadProperty("NOB", m_def_NOB)
    
       Label3.Caption = (m_ScaleMax / 4) * 3
       Label1.Caption = m_ScaleMax / 4
       Label2.Caption = m_ScaleMax / 2
    Call Scale_Visible
   ' NumOf_Bars
    Call Bar_Value(Value1, 1)
    Call Bar_Value(Value2, 2)
    Call Bar_Value(Value3, 3)
    Call Bar_Value(Value4, 4)
    Call Bar_Value(Value5, 5)
    Call Bar_Value(Value6, 6)
End Sub

Private Sub UserControl_Resize()
   Picture1.Width = UserControl.Width
   UserControl.Height = 1615
   If UserControl.Width > 3750 Then UserControl.Width = 3750
   NumOf_Bars
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("ScaleMax", m_ScaleMax, m_def_ScaleMax)
   Call PropBag.WriteProperty("ScaleVisible", m_ScaleVisible, m_def_ScaleVisible)
   Call PropBag.WriteProperty("Value1", m_Value1, m_def_Value1)
   Call PropBag.WriteProperty("Value2", m_Value2, m_def_Value2)
   Call PropBag.WriteProperty("Value3", m_Value3, m_def_Value3)
   Call PropBag.WriteProperty("Value4", m_Value4, m_def_Value4)
   Call PropBag.WriteProperty("Value5", m_Value5, m_def_Value5)
   Call PropBag.WriteProperty("Value6", m_Value6, m_def_Value6)
   Call PropBag.WriteProperty("NOB", m_NOB, m_def_NOB)
End Sub
Public Function Bar_Value(Value As Long, bar As Integer)
Dim y As String
Dim z As Long
   On Error Resume Next
   If Value < 0 Or Value > ScaleMax Then Exit Function
   bar = bar - 1
   y = (ScaleMax / Value)
   z = 1600 / y
   Shape1(bar).Top = (Picture1.Height - z)
   Shape1(bar).Height = z
   lab1(bar).Caption = Value
End Function
Public Sub NumOf_Bars()
    If NOB = 0 Then NOB = 1
    If NOB = 1 Then UserControl.Width = 930
    If NOB = 2 Then UserControl.Width = 1500
    If NOB = 3 Then UserControl.Width = 2055
    If NOB = 4 Then UserControl.Width = 2640
    If NOB = 5 Then UserControl.Width = 3225
    If NOB = 6 Then UserControl.Width = 3750
End Sub
Public Sub Scale_Visible()
    If ScaleVisible = False Then
       Label1.Visible = False
       Label2.Visible = False
       Label3.Visible = False
    Else
       Label1.Visible = True
       Label2.Visible = True
       Label3.Visible = True
    End If
End Sub
