VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1_Splash_Screen 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Welcome To ABC College | Please Wait..."
   ClientHeight    =   6552
   ClientLeft      =   39
   ClientTop       =   299
   ClientWidth     =   8216
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   21.74
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "College-Management-Student-Portal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6552
   ScaleWidth      =   8216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   598
      Left            =   1755
      TabIndex        =   3
      Top             =   5382
      Width           =   4693
      _ExtentX        =   8650
      _ExtentY        =   1102
      _Version        =   327682
      Appearance      =   0
      Max             =   105
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   936
      Top             =   585
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0E0FF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   1768
      Left            =   6669
      Shape           =   3  'Circle
      Top             =   4095
      Width           =   2470
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   2353
      Left            =   6318
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   2821
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFC0&
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   1768
      Left            =   -468
      Shape           =   3  'Circle
      Top             =   4095
      Width           =   1885
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BorderColor     =   &H00C0E0FF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   2002
      Left            =   -468
      Shape           =   3  'Circle
      Top             =   4797
      Width           =   2236
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   16.3
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   4446
      TabIndex        =   6
      Top             =   4797
      Width           =   1885
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   16.3
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   2691
      TabIndex        =   5
      Top             =   4797
      Width           =   1768
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© all rights reserved by alok kumar (ytsrex media)"
      BeginProperty Font 
         Name            =   "Technic"
         Size            =   10.87
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   1170
      TabIndex        =   4
      Top             =   6201
      Width           =   5863
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12.23
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   2691
      TabIndex        =   2
      Top             =   936
      Width           =   2587
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MANAGEMENT/STUDENT PORTAL- 2021"
      BeginProperty Font 
         Name            =   "RomanC"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   481
      Left            =   1404
      TabIndex        =   1
      Top             =   585
      Width           =   5512
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ARCADE BUSINESS COLLEGE, PATNA (BIHAR)"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   18.34
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   117
      TabIndex        =   0
      Top             =   117
      Width           =   7969
   End
   Begin VB.Image Image1 
      Height          =   3172
      Left            =   2808
      Picture         =   "College-Management-Student-Portal.frx":2FACA
      Top             =   1404
      Width           =   2509
   End
End
Attribute VB_Name = "Form1_Splash_Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label5.Caption = "Loading..."
Label6.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
Form2_loginarea.Show
End If

End Sub
