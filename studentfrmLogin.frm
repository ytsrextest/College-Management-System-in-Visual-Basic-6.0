VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form studentfrmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Login Area"
   ClientHeight    =   2132
   ClientLeft      =   2834
   ClientTop       =   3484
   ClientWidth     =   4004
   Icon            =   "studentfrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1258.027
   ScaleMode       =   0  'User
   ScaleWidth      =   3752.534
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc studentloginado 
      Height          =   299
      Left            =   1404
      Top             =   1872
      Visible         =   0   'False
      Width           =   1768
      _ExtentX        =   3259
      _ExtentY        =   551
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"studentfrmLogin.frx":1084A
      OLEDBString     =   $"studentfrmLogin.frx":108D5
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from user_data"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox studenttxtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   117
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1170
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1404
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2574
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1404
      Width           =   1140
   End
   Begin VB.TextBox studenttxtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label registerlebel 
      BackStyle       =   0  'Transparent
      Caption         =   "Register Now"
      ForeColor       =   &H8000000D&
      Height          =   247
      Left            =   1404
      TabIndex        =   7
      Top             =   936
      Width           =   949
   End
   Begin VB.Image Image1 
      Height          =   1066
      Left            =   117
      Picture         =   "studentfrmLogin.frx":10960
      Stretch         =   -1  'True
      Top             =   1053
      Width           =   949
   End
   Begin VB.Label studentforgetpw 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Forget Password?"
      ForeColor       =   &H8000000D&
      Height          =   247
      Left            =   2457
      TabIndex        =   6
      Top             =   936
      Width           =   1300
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "studentfrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Public LoginUsername As String

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Form2_loginarea.Show
    
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()
    studentloginado.RecordSource = "select * from user_data where username='" + studenttxtUserName.Text + "' and pass='" + studenttxtPassword.Text + "'"
    studentloginado.Refresh
    LoginSucceeded = True
    LoginUsername = studentloginado.UserName
    
  
    
    If studentloginado.Recordset.EOF Then
    MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
    studenttxtPassword.SetFocus
    SendKeys "{Home}+{End}"
    Else
       'MsgBox "Login Successful.", vbInformation, "Thank You"
       Me.Hide
       MDIuserdashboard.Show
       
    End If
End Sub



Private Sub registerlebel_Click()
registerDialog.Show
End Sub

Private Sub studentforgetpw_Click()
Forgetpwstudent.Show

End Sub

