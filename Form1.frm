VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rush Hour "
   ClientHeight    =   10740
   ClientLeft      =   3465
   ClientTop       =   2430
   ClientWidth     =   10290
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   39
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   -1  'True
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14.948
   ScaleMode       =   0  'User
   ScaleWidth      =   9.636
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   5280
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   240
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar hsLevel 
      Height          =   255
      Left            =   1320
      Max             =   40
      Min             =   1
      TabIndex        =   0
      Top             =   5160
      Value           =   1
      Width           =   5655
   End
   Begin VB.Image CarV2 
      Height          =   1485
      Index           =   3
      Left            =   9360
      Picture         =   "Form1.frx":08CA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   795
   End
   Begin VB.Image CarV2 
      Height          =   1485
      Index           =   2
      Left            =   9360
      Picture         =   "Form1.frx":5894
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   795
   End
   Begin VB.Image CarV2 
      Height          =   1485
      Index           =   1
      Left            =   9360
      Picture         =   "Form1.frx":67F4
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   795
   End
   Begin VB.Image CarV3 
      Height          =   2355
      Index           =   3
      Left            =   9360
      Picture         =   "Form1.frx":A191
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   780
   End
   Begin VB.Image CarV3 
      Height          =   2355
      Index           =   2
      Left            =   6840
      Picture         =   "Form1.frx":B1CA
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   780
   End
   Begin VB.Image CarV3 
      Height          =   2355
      Index           =   1
      Left            =   7680
      Picture         =   "Form1.frx":FA82
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   780
   End
   Begin VB.Image CarH2 
      Height          =   780
      Index           =   3
      Left            =   5160
      Picture         =   "Form1.frx":14A0D
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   1485
   End
   Begin VB.Image CarH3 
      Height          =   780
      Index           =   3
      Left            =   120
      Picture         =   "Form1.frx":1975B
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   2355
   End
   Begin VB.Image CarH3 
      Height          =   780
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":1AA69
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   2355
   End
   Begin VB.Image CarH3 
      Height          =   780
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":1F4F0
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   2355
   End
   Begin VB.Image CarH2 
      Height          =   780
      Index           =   2
      Left            =   5160
      Picture         =   "Form1.frx":207E9
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   1485
   End
   Begin VB.Image CarH2 
      Height          =   780
      Index           =   1
      Left            =   5160
      Picture         =   "Form1.frx":21157
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   1485
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   23
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   22
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   21
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   20
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   19
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   18
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   17
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   16
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   15
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   14
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   13
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   12
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   11
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   10
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   9
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   8
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   7
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   6
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   5
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   4
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   3
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   2
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   1
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgCar 
      Height          =   375
      Index           =   0
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image CarH2 
      Height          =   780
      Index           =   0
      Left            =   2760
      Picture         =   "Form1.frx":21E60
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   1485
   End
   Begin VB.Image CarV3 
      Height          =   2355
      Index           =   0
      Left            =   8520
      Picture         =   "Form1.frx":22C5C
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   780
   End
   Begin VB.Image CarV2 
      Height          =   1485
      Index           =   0
      Left            =   9360
      Picture         =   "Form1.frx":23F3C
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   795
   End
   Begin VB.Image CarH3 
      Height          =   780
      Index           =   0
      Left            =   2640
      Picture         =   "Form1.frx":28D23
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   2355
   End
   Begin VB.Image MainCar 
      Height          =   780
      Left            =   2760
      Picture         =   "Form1.frx":2D492
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   1485
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   3210
      TabIndex        =   3
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label lblStep 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Steps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   359
      Left            =   1068
      TabIndex        =   2
      Top             =   108
      Width           =   1602
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Level : 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   5520
      Width           =   6375
   End
   Begin VB.Line Line5 
      BorderWidth     =   5
      X1              =   7
      X2              =   7
      Y1              =   4
      Y2              =   7.001
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   7
      X2              =   7
      Y1              =   0.999
      Y2              =   3.001
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   1.011
      X2              =   7.012
      Y1              =   0.835
      Y2              =   0.835
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   1
      X2              =   7
      Y1              =   7.001
      Y2              =   7.001
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   1
      X2              =   1
      Y1              =   0.999
      Y2              =   7.001
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   9
      X1              =   0.337
      X2              =   2.809
      Y1              =   9.353
      Y2              =   9.353
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   8
      X1              =   0.337
      X2              =   2.809
      Y1              =   9.52
      Y2              =   9.52
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   7
      X1              =   0.337
      X2              =   2.809
      Y1              =   9.687
      Y2              =   9.687
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   6
      X1              =   0.337
      X2              =   2.809
      Y1              =   9.854
      Y2              =   9.854
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   5
      X1              =   2.584
      X2              =   5.056
      Y1              =   9.186
      Y2              =   9.186
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   4
      X1              =   2.584
      X2              =   5.056
      Y1              =   9.353
      Y2              =   9.353
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   3
      X1              =   2.584
      X2              =   5.056
      Y1              =   9.52
      Y2              =   9.52
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   2
      X1              =   2.584
      X2              =   5.056
      Y1              =   9.687
      Y2              =   9.687
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   1
      X1              =   2.584
      X2              =   5.056
      Y1              =   9.854
      Y2              =   9.854
   End
   Begin VB.Line linGrid 
      BorderColor     =   &H80000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   0
      X1              =   0.337
      X2              =   2.809
      Y1              =   9.186
      Y2              =   9.186
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "&Open User Defined Level"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x0 As Single, y0 As Single, n As Integer, s As Integer, T As Integer

Private Sub Form_Load()
Form1.Width = Screen.Width / 2.5
Form1.Height = Form1.Width * 1.15

Form1.ScaleWidth = 8
Form1.ScaleHeight = 8

hsLevel.Left = 1
hsLevel.Top = 7.2
hsLevel.Height = 0.3
hsLevel.Width = 6
hsLevel.Min = 1
hsLevel.Max = 40
hsLevel.LargeChange = 1
hsLevel.SmallChange = 1

lblLevel.Top = 7.5
lblLevel.Left = 0
lblLevel.Width = 8
lblLevel.Height = 0.4

Line1.X1 = 0.9
Line1.X2 = 7.05
Line1.Y1 = 0.9
Line1.Y2 = 0.9

Line2.X1 = 0.9
Line2.X2 = 7.05
Line2.Y1 = 7.1
Line2.Y2 = 7.1

Line3.X1 = 0.9
Line3.X2 = 0.9
Line3.Y1 = 0.9
Line3.Y2 = 7.1

Line4.X1 = 7.1
Line4.X2 = 7.1
Line4.Y1 = 0.9
Line4.Y2 = 3

Line5.X1 = 7.1
Line5.X2 = 7.1
Line5.Y1 = 4
Line5.Y2 = 7.1

' Horizontal Grid :
linGrid(0).X1 = 0.9
linGrid(0).X2 = 7.1
linGrid(0).Y1 = 2
linGrid(0).Y2 = 2

linGrid(1).X1 = 0.9
linGrid(1).X2 = 7.1
linGrid(1).Y1 = 3
linGrid(1).Y2 = 3

linGrid(2).X1 = 0.9
linGrid(2).X2 = 7.1
linGrid(2).Y1 = 4
linGrid(2).Y2 = 4

linGrid(3).X1 = 0.9
linGrid(3).X2 = 7.1
linGrid(3).Y1 = 5
linGrid(3).Y2 = 5

linGrid(4).X1 = 0.9
linGrid(4).X2 = 7.1
linGrid(4).Y1 = 6
linGrid(4).Y2 = 6
' Vertical Grid :
linGrid(5).X1 = 2
linGrid(5).X2 = 2
linGrid(5).Y1 = 0.9
linGrid(5).Y2 = 7.1

linGrid(6).X1 = 3
linGrid(6).X2 = 3
linGrid(6).Y1 = 0.9
linGrid(6).Y2 = 7.1

linGrid(7).X1 = 4
linGrid(7).X2 = 4
linGrid(7).Y1 = 0.9
linGrid(7).Y2 = 7.1

linGrid(8).X1 = 5
linGrid(8).X2 = 5
linGrid(8).Y1 = 0.9
linGrid(8).Y2 = 7.1

linGrid(9).X1 = 6
linGrid(9).X2 = 6
linGrid(9).Y1 = 0.9
linGrid(9).Y2 = 7.1

lblStep.Left = 1
lblTime.Left = 4.7
lblLevel.Left = 0
Call Arrange
End Sub

Private Sub Arrange()
Dim i As Integer
For i = 0 To 23
  imgCar(i).Visible = False
Next i
n = -1
s = 0 ' Steps
T = 0 ' Time
lblStep.Caption = "Steps"
lblTime.Caption = "Time"
' =========================================
' imgCar(0) is allways the Main Block
Select Case hsLevel.Value
  Case 1
    '           W  H  T  L ==> Width,Height,Top,Left
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 1)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(1, 2, 3, 5)
    Call SetIni(3, 1, 4, 1)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 1)
    Call SetIni(2, 1, 6, 4)
  Case 2
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 3, 1, 6)
    Call SetIni(1, 3, 2, 1)
    Call SetIni(1, 3, 2, 4)
    Call SetIni(1, 2, 5, 1)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(3, 1, 6, 3)
  Case 3
    Call SetIni(2, 1, 3, 2)
    Call SetIni(1, 3, 1, 1)
    Call SetIni(1, 3, 1, 4)
    Call SetIni(1, 2, 4, 3)
    Call SetIni(3, 1, 4, 4)
    Call SetIni(1, 2, 5, 6)
    Call SetIni(3, 1, 6, 3)
  Case 4
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 4, 2)
    Call SetIni(1, 3, 3, 4)
    Call SetIni(1, 3, 4, 5)
    Call SetIni(1, 2, 5, 2)
    Call SetIni(2, 1, 6, 3)
  Case 5
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 2, 1, 4)
    Call SetIni(2, 1, 2, 1)
    Call SetIni(1, 3, 2, 5)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(1, 3, 3, 4)
    Call SetIni(2, 1, 4, 1)
    Call SetIni(1, 2, 4, 3)
    Call SetIni(1, 2, 5, 1)
    Call SetIni(3, 1, 6, 4)
  Case 6
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 3, 1, 4)
    Call SetIni(1, 2, 1, 6)
    Call SetIni(1, 3, 2, 1)
    Call SetIni(1, 3, 2, 5)
    Call SetIni(1, 2, 3, 6)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(1, 2, 5, 1)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 5)
  Case 7
    Call SetIni(2, 1, 3, 2)
    Call SetIni(1, 2, 1, 2)
    Call SetIni(2, 1, 1, 3)
    Call SetIni(1, 2, 1, 5)
    Call SetIni(1, 2, 1, 6)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(1, 2, 3, 6)
    Call SetIni(2, 1, 4, 3)
    Call SetIni(1, 2, 5, 4)
  Case 8
    Call SetIni(2, 1, 3, 1)
    Call SetIni(2, 1, 1, 4)
    Call SetIni(1, 3, 1, 6)
    Call SetIni(2, 1, 2, 3)
    Call SetIni(1, 2, 2, 5)
    Call SetIni(1, 2, 3, 3)
    Call SetIni(1, 2, 3, 4)
    Call SetIni(2, 1, 4, 1)
    Call SetIni(2, 1, 4, 5)
    Call SetIni(2, 1, 5, 1)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(3, 1, 5, 4)
    Call SetIni(2, 1, 6, 1)
    Call SetIni(3, 1, 6, 4)
  Case 9
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 2)
    Call SetIni(2, 1, 1, 3)
    Call SetIni(2, 1, 1, 5)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(2, 1, 2, 5)
    Call SetIni(1, 3, 3, 5)
    Call SetIni(1, 2, 3, 6)
    Call SetIni(1, 3, 4, 1)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(1, 2, 5, 6)
  Case 10
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(2, 1, 1, 5)
    Call SetIni(2, 1, 2, 1)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(1, 3, 3, 1)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(1, 2, 5, 4)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 1)
    Call SetIni(2, 1, 6, 5)
  Case 11
    Call SetIni(2, 1, 3, 2)
    Call SetIni(1, 3, 1, 1)
    Call SetIni(2, 1, 1, 2)
    Call SetIni(1, 3, 1, 4)
    Call SetIni(1, 2, 4, 3)
    Call SetIni(3, 1, 4, 4)
    Call SetIni(1, 2, 5, 6)
    Call SetIni(3, 1, 6, 3)
  Case 12
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 1)
    Call SetIni(2, 1, 1, 2)
    Call SetIni(1, 3, 1, 6)
    Call SetIni(1, 3, 2, 3)
    Call SetIni(3, 1, 4, 4)
    Call SetIni(1, 2, 5, 5)
    Call SetIni(3, 1, 6, 1)
  Case 13
    Call SetIni(2, 1, 3, 4)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(2, 1, 1, 3)
    Call SetIni(1, 2, 1, 5)
    Call SetIni(1, 2, 2, 3)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(1, 2, 3, 2)
    Call SetIni(1, 3, 4, 1)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(1, 2, 5, 4)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 2)
    Call SetIni(2, 1, 6, 5)
  Case 14
    Call SetIni(2, 1, 3, 3)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(2, 1, 2, 5)
    Call SetIni(1, 2, 3, 1)
    Call SetIni(1, 2, 3, 2)
    Call SetIni(1, 2, 3, 5)
    Call SetIni(1, 2, 3, 6)
    Call SetIni(2, 1, 4, 3)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 1)
  Case 15
    Call SetIni(2, 1, 3, 3)
    Call SetIni(2, 1, 1, 2)
    Call SetIni(2, 1, 1, 4)
    Call SetIni(2, 1, 2, 1)
    Call SetIni(2, 1, 2, 3)
    Call SetIni(1, 3, 2, 5)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(1, 3, 3, 1)
    Call SetIni(1, 3, 3, 2)
    Call SetIni(1, 2, 4, 3)
    Call SetIni(1, 2, 4, 4)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 2)
    Call SetIni(2, 1, 6, 4)
  Case 16
    Call SetIni(2, 1, 3, 4)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(2, 1, 1, 3)
    Call SetIni(1, 2, 1, 5)
    Call SetIni(1, 3, 1, 6)
    Call SetIni(1, 2, 2, 1)
    Call SetIni(2, 1, 2, 3)
    Call SetIni(1, 2, 3, 2)
    Call SetIni(1, 3, 3, 3)
    Call SetIni(3, 1, 4, 4)
    Call SetIni(2, 1, 6, 1)
  Case 17
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 1)
    Call SetIni(3, 1, 1, 2)
    Call SetIni(2, 1, 2, 3)
    Call SetIni(2, 1, 2, 5)
    Call SetIni(1, 2, 3, 3)
    Call SetIni(2, 1, 4, 1)
    Call SetIni(1, 3, 4, 4)
    Call SetIni(3, 1, 5, 1)
    Call SetIni(1, 2, 5, 5)
    Call SetIni(1, 2, 5, 6)
    Call SetIni(3, 1, 6, 1)
  Case 18
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(1, 3, 1, 4)
    Call SetIni(2, 1, 2, 1)
    Call SetIni(1, 3, 3, 1)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(2, 1, 5, 2)
    Call SetIni(3, 1, 6, 1)
  Case 19
    Call SetIni(2, 1, 3, 3)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(2, 1, 1, 4)
    Call SetIni(1, 2, 2, 5)
    Call SetIni(1, 2, 3, 2)
    Call SetIni(2, 1, 4, 3)
    Call SetIni(1, 2, 4, 5)
    Call SetIni(3, 1, 5, 2)
  Case 20
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 1)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(2, 1, 2, 2)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(1, 2, 3, 3)
    Call SetIni(1, 3, 3, 6)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(2, 1, 5, 4)
    Call SetIni(3, 1, 6, 4)
  Case 21
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(1, 3, 1, 4)
    Call SetIni(1, 3, 2, 1)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(3, 1, 6, 4)
  Case 22
    Call SetIni(2, 1, 3, 2)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(1, 2, 2, 1)
    Call SetIni(1, 3, 2, 4)
    Call SetIni(2, 1, 2, 5)
    Call SetIni(1, 2, 4, 2)
    Call SetIni(2, 1, 4, 5)
    Call SetIni(1, 2, 5, 1)
    Call SetIni(2, 1, 5, 3)
    Call SetIni(1, 2, 5, 6)
    Call SetIni(3, 1, 6, 2)
  Case 23
    Call SetIni(2, 1, 3, 4)
    Call SetIni(3, 1, 1, 3)
    Call SetIni(1, 3, 1, 6)
    Call SetIni(1, 2, 2, 3)
    Call SetIni(2, 1, 2, 4)
    Call SetIni(1, 2, 4, 3)
    Call SetIni(1, 2, 4, 4)
    Call SetIni(2, 1, 4, 5)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(3, 1, 6, 3)
  Case 24
    Call SetIni(2, 1, 3, 3)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(2, 1, 1, 4)
    Call SetIni(1, 2, 2, 2)
    Call SetIni(1, 2, 3, 1)
    Call SetIni(1, 2, 3, 5)
    Call SetIni(2, 1, 4, 2)
    Call SetIni(3, 1, 5, 1)
    Call SetIni(1, 2, 5, 5)
    Call SetIni(2, 1, 6, 1)
  Case 25
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(2, 1, 1, 5)
    Call SetIni(2, 1, 2, 1)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(1, 3, 3, 1)
    Call SetIni(1, 2, 3, 5)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(1, 2, 5, 2)
    Call SetIni(1, 2, 5, 4)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 5)
  Case 26
    Call SetIni(2, 1, 3, 2)
    Call SetIni(1, 2, 1, 2)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(1, 2, 2, 1)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(1, 3, 2, 5)
    Call SetIni(1, 2, 3, 6)
    Call SetIni(1, 2, 4, 1)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(1, 2, 5, 6)
    Call SetIni(2, 1, 6, 4)
  Case 27
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 1)
    Call SetIni(2, 1, 1, 2)
    Call SetIni(1, 3, 1, 4)
    Call SetIni(2, 1, 2, 2)
    Call SetIni(1, 2, 3, 3)
    Call SetIni(1, 3, 3, 6)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(3, 1, 6, 4)
  Case 28
    Call SetIni(2, 1, 3, 1)
    Call SetIni(3, 1, 1, 1)
    Call SetIni(1, 2, 1, 4)
    Call SetIni(1, 3, 2, 3)
    Call SetIni(2, 1, 2, 5)
    Call SetIni(1, 2, 4, 1)
    Call SetIni(1, 2, 4, 2)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(1, 3, 4, 6)
    Call SetIni(3, 1, 5, 3)
    Call SetIni(2, 1, 6, 1)
    Call SetIni(2, 1, 6, 3)
  Case 29
    Call SetIni(2, 1, 3, 1)
    Call SetIni(3, 1, 1, 1)
    Call SetIni(1, 3, 1, 5)
    Call SetIni(1, 2, 2, 3)
    Call SetIni(1, 2, 3, 6)
    Call SetIni(1, 2, 4, 1)
    Call SetIni(2, 1, 4, 2)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(2, 1, 5, 2)
    Call SetIni(1, 2, 5, 4)
    Call SetIni(1, 2, 5, 6)
    Call SetIni(3, 1, 6, 1)
  Case 30
    Call SetIni(2, 1, 3, 2)
    Call SetIni(1, 3, 1, 1)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(2, 1, 4, 1)
    Call SetIni(2, 1, 4, 3)
    Call SetIni(1, 3, 4, 6)
    Call SetIni(2, 1, 6, 1)
    Call SetIni(2, 1, 6, 3)
  Case 31
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(2, 1, 2, 5)
    Call SetIni(1, 2, 3, 1)
    Call SetIni(3, 1, 6, 4)
    Call SetIni(1, 3, 3, 6)
    Call SetIni(1, 3, 4, 3)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(2, 1, 5, 1)
  Case 32
    Call SetIni(2, 1, 3, 1)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 3, 1, 3)
    Call SetIni(1, 2, 1, 4)
    Call SetIni(2, 1, 1, 5)
    Call SetIni(1, 2, 4, 1)
    Call SetIni(2, 1, 4, 2)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(1, 3, 4, 6)
    Call SetIni(1, 2, 5, 4)
    Call SetIni(2, 1, 6, 1)
  Case 33
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 2)
    Call SetIni(1, 3, 1, 3)
    Call SetIni(2, 1, 1, 5)
    Call SetIni(3, 1, 6, 1)
    Call SetIni(1, 2, 4, 1)
    Call SetIni(2, 1, 4, 2)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(1, 3, 4, 6)
    Call SetIni(2, 1, 5, 2)
    Call SetIni(1, 2, 5, 4)
    Call SetIni(1, 2, 5, 5)
  Case 34
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 1)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(1, 2, 3, 5)
    Call SetIni(3, 1, 4, 1)
    Call SetIni(1, 2, 4, 4)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 1)
    Call SetIni(2, 1, 6, 4)
  Case 35
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 3, 1, 3)
    Call SetIni(2, 1, 1, 4)
    Call SetIni(1, 3, 1, 6)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(1, 2, 4, 1)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(2, 1, 5, 2)
    Call SetIni(1, 2, 5, 4)
    Call SetIni(1, 2, 5, 5)
    Call SetIni(2, 1, 6, 1)
  Case 36
    Call SetIni(2, 1, 3, 3)
    Call SetIni(1, 3, 1, 1)
    Call SetIni(3, 1, 1, 2)
    Call SetIni(2, 1, 1, 5)
    Call SetIni(1, 2, 2, 2)
    Call SetIni(2, 1, 2, 3)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(3, 1, 4, 1)
    Call SetIni(1, 2, 4, 4)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 1)
  Case 37
    Call SetIni(2, 1, 3, 2)
    Call SetIni(2, 1, 1, 1)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(2, 1, 1, 5)
    Call SetIni(2, 1, 2, 1)
    Call SetIni(1, 3, 2, 5)
    Call SetIni(1, 3, 2, 6)
    Call SetIni(1, 3, 3, 1)
    Call SetIni(2, 1, 6, 5)
    Call SetIni(3, 1, 4, 2)
    Call SetIni(1, 2, 5, 4)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 1)
  Case 38
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 1)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(2, 1, 2, 2)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(1, 2, 3, 3)
    Call SetIni(1, 3, 3, 6)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(2, 1, 5, 4)
    Call SetIni(3, 1, 6, 4)
  Case 39
    Call SetIni(2, 1, 3, 4)
    Call SetIni(1, 3, 1, 1)
    Call SetIni(2, 1, 1, 2)
    Call SetIni(1, 2, 1, 5)
    Call SetIni(1, 2, 2, 2)
    Call SetIni(1, 2, 2, 3)
    Call SetIni(2, 1, 6, 4)
    Call SetIni(3, 1, 4, 1)
    Call SetIni(1, 2, 4, 4)
    Call SetIni(1, 2, 5, 3)
    Call SetIni(2, 1, 5, 5)
    Call SetIni(2, 1, 6, 1)
    Call SetIni(1, 3, 2, 6)
  Case 40
    Call SetIni(2, 1, 3, 1)
    Call SetIni(1, 2, 1, 3)
    Call SetIni(3, 1, 1, 4)
    Call SetIni(1, 2, 2, 4)
    Call SetIni(1, 2, 3, 3)
    Call SetIni(1, 3, 3, 6)
    Call SetIni(2, 1, 4, 1)
    Call SetIni(2, 1, 4, 4)
    Call SetIni(1, 2, 5, 1)
    Call SetIni(1, 2, 5, 2)
    Call SetIni(2, 1, 5, 3)
    Call SetIni(2, 1, 6, 3)
End Select
End Sub
Private Sub SetIni(W As Integer, H As Integer, T As Integer, L As Integer)
Dim nn As Integer
n = n + 1
imgCar(n).Visible = True
imgCar(n).Width = W
imgCar(n).Height = H
imgCar(n).Top = T
imgCar(n).Left = L

Randomize Timer
nn = Int(Rnd * 4)
Select Case W
  Case 1
    Select Case H
      Case 2
        'V2
        imgCar(n).Picture = CarV2(nn).Picture
      Case 3
        'V3
        imgCar(n).Picture = CarV3(nn).Picture
    End Select
  Case 2
    'H2
    imgCar(n).Picture = CarH2(nn).Picture
  Case 3
    'H3
    imgCar(n).Picture = CarH3(nn).Picture
End Select
If n = 0 Then imgCar(n).Picture = MainCar.Picture
End Sub

Private Sub hsLevel_Change()
lblLevel.Caption = "Level : " + Str$(hsLevel.Value)
Call Arrange
End Sub

Private Sub imgCar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
x0 = X
y0 = Y

s = s + 1
lblStep.Caption = Str(s)
End Sub

Private Sub imgCar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xx, yy

If Button <> vbLeftButton Then Exit Sub

If imgCar(Index).Width < imgCar(Index).Height Then
    ' Vertical Block
    dy = ScaleY(Y - y0)
    If dy = 0 Then Exit Sub
    'Check Borders
    If imgCar(Index).Top + dy < 1 Then Exit Sub
    If imgCar(Index).Top + imgCar(Index).Height + dy > 7 Then Exit Sub
  
    If dy > 0 Then yy = imgCar(Index).Top + imgCar(Index).Height + 0.1
    If dy < 0 Then yy = imgCar(Index).Top + dy - 0.1
  
    xx = imgCar(Index).Left + imgCar(Index).Width / 2
  
    If Form1.Point(xx, yy) <> vbWhite Then
      y0 = Y
      Exit Sub
    End If
    
  imgCar(Index).Top = imgCar(Index).Top + dy
End If

If imgCar(Index).Width > imgCar(Index).Height Then
    ' Horizontal Block
    dx = ScaleX(X - x0)
    If dx = 0 Then Exit Sub
    ' Check Borders
    If imgCar(Index).Left + dx < 1 Then Exit Sub
    If imgCar(Index).Left + imgCar(Index).Width + dx > 7 Then Exit Sub
  
    If dx > 0 Then xx = imgCar(Index).Left + imgCar(Index).Width + 0.1
    If dx < 0 Then xx = imgCar(Index).Left + dx - 0.1
  
    yy = imgCar(Index).Top + imgCar(Index).Height / 2
  
    If Form1.Point(xx, yy) <> vbWhite Then
      x0 = X
      Exit Sub
    End If
    
    imgCar(Index).Left = imgCar(Index).Left + dx
End If

End Sub

Private Sub imgCar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCar(Index).Left = Int(imgCar(Index).Left + 0.5)
imgCar(Index).Top = Int(imgCar(Index).Top + 0.5)
If Abs(imgCar(0).Left - 5) < 0.1 Then
  Timer1.Enabled = False
  MsgBox "Congrdulations!", , "Level Finished !"
  Timer1.Enabled = True
  If hsLevel.Value = 40 Then
    MsgBox "Congradulations, You finished the GAME !!!", , "Congradulations !!!"
    Else
    hsLevel.Value = hsLevel.Value + 1
  End If
End If

End Sub

Private Sub mnuAbout_Click()
MsgBox "Program : Rush Hour" + vbCrLf + _
       "Programmer : Mahdi Shakouri Rad" + vbCrLf + _
       "email : Mahdi.ShakouriRad@GMail.com" + vbCrLf + _
       "Programming Language : Visual Basic 6" + vbCrLf + _
       "Program version : 1.00, 2005,Apr.", , "About ..."
       
End Sub

Private Sub mnuExit_Click()
Dim ans As Integer
ans = MsgBox("Do you realy want to exit ?", vbYesNo, "warning !")
If ans = vbYes Then End
End Sub

Private Sub mnuHelp_Click()
MsgBox "The purpose of the game is to bring out the Red car." + vbCrLf + _
       "You are only able to move each car straightly." + vbCrLf + _
       "Try to do each level in minimum Steps and Time !" + vbCrLf + _
       "================================================" + vbCrLf + _
       "================================================" + vbCrLf + _
       "For defining a user level, create a text file, for each car" + vbCrLf + _
       "enter its Width, Height, Top and Left ... which :" + vbCrLf + _
       "Width : Width of the car (1,2 or 3)" + vbCrLf + _
       "Height : Height of the car (1,2 or 3)" + vbCrLf + _
       "Top : Top coordinate of the car (1 to 6)" + vbCrLf + _
       "Left : Left coordinate of the car (1 to 6)" + vbCrLf + _
       "Note 1 : The Board is 6x6" + vbCrLf + _
       "Note 2 : The origin of the coordinate is located at Top and Left" + vbCrLf + _
       "Note 3 : The X coordinate increases from left to right" + vbCrLf + _
       "Note 4 : The Y coordinate increases from top to bottom" + vbCrLf + _
       "Note 5 : One of the Width and Height should be 1" + vbCrLf + _
       "Note 6 : The Red piece should be defined as first piece" + vbCrLf + _
       "================================================" + vbCrLf + _
       "Enjoy ..." + vbCrLf + _
       "Mahdi.ShakouriRad@gmail.com" + vbCrLf + _
       ">>> NOTE : All levels have been checked and are correct.", , "Help"
       
End Sub

Private Sub mnuOpen_Click()
' User Defined Level
Dim i As Integer, f As String
Dim W As Integer, H As Integer, T As Integer, L As Integer

For i = 0 To 23
  imgCar(i).Visible = False
Next i
n = -1
s = 0 ' Steps
T = 0 ' Time
lblStep.Caption = "Steps"
lblTime.Caption = "Time"
' =========================================
' imgCar(0) is allways the Main Block
CD.DialogTitle = "Open User Defined Level"
CD.InitDir = App.Path
CD.ShowOpen
f = CD.FileName
If f = "" Then
  Call Arrange
  Exit Sub
End If
lblLevel.Caption = "User Defined Level : " + Right(f, Len(f) - Len(App.Path) - 1)
Open f For Input As 1

Do While Not EOF(1)
  Input #1, W, H, T, L     ' ==> Width,Height,Top,Left
  Call SetIni(W, H, T, L)
Loop
Close 1
End Sub

Private Sub Timer1_Timer()
If s = 0 Then Exit Sub
T = T + 1
lblTime = Str(T)
End Sub
