VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Userprofile 
   BackColor       =   &H00FFFFFF&
   Caption         =   "User Profile"
   ClientHeight    =   5135
   ClientLeft      =   104
   ClientTop       =   416
   ClientWidth     =   6253
   Icon            =   "Userprofile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5135
   ScaleWidth      =   6253
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc fetchuserdetailsado 
      Height          =   364
      Left            =   4212
      Top             =   2808
      Visible         =   0   'False
      Width           =   1768
      _ExtentX        =   3259
      _ExtentY        =   671
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
      Connect         =   $"Userprofile.frx":2FACA
      OLEDBString     =   $"Userprofile.frx":2FB55
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from user_data"
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
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      DataField       =   "mobile"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   19
      Top             =   3861
      Width           =   2470
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No.:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   18
      Top             =   3861
      Width           =   1183
   End
   Begin VB.Image Image1 
      Height          =   1768
      Left            =   4212
      Stretch         =   -1  'True
      Top             =   585
      Width           =   1885
   End
   Begin VB.Label showimglink 
      DataField       =   "image"
      DataSource      =   "fetchuserdetailsado"
      Height          =   247
      Left            =   4212
      TabIndex        =   17
      Top             =   3159
      Visible         =   0   'False
      Width           =   1885
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      DataField       =   "gender"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   16
      Top             =   3393
      Width           =   2236
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   15
      Top             =   3393
      Width           =   949
   End
   Begin VB.Image Image2 
      Height          =   715
      Left            =   234
      Picture         =   "Userprofile.frx":2FBE0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3406
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "session"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   247
      Left            =   4329
      TabIndex        =   14
      Top             =   2457
      Width           =   1417
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   13
      Top             =   2457
      Width           =   1066
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   12
      Top             =   4329
      Width           =   1066
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   11
      Top             =   2925
      Width           =   1066
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Roll. No.:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   247
      TabIndex        =   10
      Top             =   1989
      Width           =   1066
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   247
      Left            =   4446
      TabIndex        =   9
      Top             =   234
      Width           =   1066
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   247
      TabIndex        =   8
      Top             =   1053
      Width           =   1066
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   247
      TabIndex        =   7
      Top             =   1521
      Width           =   1066
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      DataField       =   "course"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Left            =   1638
      TabIndex        =   6
      Top             =   2457
      Width           =   1989
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      DataField       =   "address"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   1625
      TabIndex        =   5
      Top             =   4329
      Width           =   4342
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      DataField       =   "email"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Left            =   1638
      TabIndex        =   4
      Top             =   2925
      Width           =   3159
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      DataField       =   "rollno"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   3
      Top             =   1989
      Width           =   2002
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "ID"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   247
      Left            =   5382
      TabIndex        =   2
      Top             =   234
      Width           =   598
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   1
      Top             =   1053
      Width           =   2002
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      DataField       =   "namee"
      DataSource      =   "fetchuserdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   0
      Top             =   1521
      Width           =   2002
   End
End
Attribute VB_Name = "Userprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
'studentfrmLogin.studenttxtUserName
Label2.Caption = studentfrmLogin.studenttxtUserName
fetchuserdetailsado.RecordSource = "select * from user_data where username='" + Label2.Caption + "'"
fetchuserdetailsado.Refresh

End Sub


Private Sub Label2_Click()
Label2.Caption = studentfrmLogin.studenttxtUserName
End Sub




Private Sub showimglink_Change()
Image1.Picture = LoadPicture(showimglink.Caption)
End Sub

