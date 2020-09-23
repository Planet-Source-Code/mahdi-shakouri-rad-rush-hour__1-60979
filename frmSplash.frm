VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5985
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5715
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7185
      Begin VB.Image Image1 
         Height          =   1710
         Left            =   480
         Picture         =   "frmSplash.frx":08CA
         Top             =   3360
         Width           =   6285
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   " Press any key to start the game ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3000
         Width           =   6615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cars can move only in straight line."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   5
         Top             =   2400
         Width           =   5325
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Try to move the  red car  to the exit !"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   4
         Top             =   1680
         Width           =   6210
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Rush Hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   765
         Left            =   3240
         TabIndex        =   3
         Top             =   600
         Width           =   3285
      End
      Begin VB.Label lblCompany 
         Caption         =   "2005,April"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   2
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblCopyright 
         Caption         =   " Mahdi.ShakouriRad@gmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   5160
         Width           =   3855
      End
      Begin VB.Image imgLogo 
         Height          =   945
         Left            =   360
         Picture         =   "frmSplash.frx":7A8F
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Form1.Show
End Sub

Private Sub Frame1_Click()
    Unload Me
    Form1.Show
End Sub
