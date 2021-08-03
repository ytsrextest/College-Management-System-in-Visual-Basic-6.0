VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form submirrojectform 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Submit Project"
   ClientHeight    =   3107
   ClientLeft      =   104
   ClientTop       =   416
   ClientWidth     =   5044
   Icon            =   "submirrojectform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3107
   ScaleWidth      =   5044
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc projectadb 
      Height          =   364
      Left            =   3042
      Top             =   468
      Visible         =   0   'False
      Width           =   1651
      _ExtentX        =   3043
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
      Connect         =   $"submirrojectform.frx":2FACA
      OLEDBString     =   $"submirrojectform.frx":2FB55
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from project"
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
   Begin VB.TextBox driveurl 
      DataField       =   "url"
      DataSource      =   "projectadb"
      Height          =   364
      Left            =   1053
      TabIndex        =   5
      Top             =   1521
      Width           =   2119
   End
   Begin VB.CommandButton submitbtnroject 
      BackColor       =   &H0000FF00&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   1404
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1989
      Width           =   1534
   End
   Begin VB.TextBox projecturl 
      DataField       =   "rollno"
      DataSource      =   "projectadb"
      Height          =   364
      Left            =   1053
      TabIndex        =   1
      Top             =   819
      Width           =   2119
   End
   Begin MSAdodcLib.Adodc matctdetailsado 
      Height          =   364
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1300
      _ExtentX        =   2396
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
      Connect         =   $"submirrojectform.frx":2FBE0
      OLEDBString     =   $"submirrojectform.frx":2FC6B
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
   Begin VB.Label Label5 
      Height          =   247
      Left            =   117
      TabIndex        =   8
      Top             =   468
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.Label Label4 
      DataField       =   "datee"
      DataSource      =   "projectadb"
      Height          =   247
      Left            =   117
      TabIndex        =   7
      Top             =   1872
      Visible         =   0   'False
      Width           =   1066
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE:  Upload your project in zip format on Google drive, Set Visiblity Public then coy URL and fill this form."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   598
      Left            =   410
      TabIndex        =   6
      Top             =   2691
      Width           =   4225
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Drive URL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   4
      Top             =   1287
      Width           =   1534
   End
   Begin VB.Label Label2rollno 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Roll. No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   3
      Top             =   585
      Width           =   1183
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SUBMIT PROJECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12.23
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   364
      Left            =   1112
      TabIndex        =   0
      Top             =   117
      Width           =   2821
   End
   Begin VB.Image Image1 
      Height          =   2821
      Left            =   2223
      Picture         =   "submirrojectform.frx":2FCF6
      Stretch         =   -1  'True
      Top             =   117
      Width           =   2821
   End
End
Attribute VB_Name = "submirrojectform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Form_Load()
Label4.Caption = Date
projectadb.Recordset.AddNew
Label4.Caption = Date
Label5.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub Label4_Click()
Label4.Caption = Date
End Sub

Private Sub Label5_Click()
Label5.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub submitbtnroject_Click()
matctdetailsado.RecordSource = "select * from user_data where username='" + Label5.Caption + "' and rollno='" + projecturl.Text + "'"
matctdetailsado.Refresh

If projecturl = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
projecturl.SetFocus
SendKeys "{Home}+{End}"

ElseIf driveurl = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
driveurl.SetFocus
SendKeys "{Home}+{End}"

ElseIf matctdetailsado.Recordset.EOF Then
 MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
    projecturl.SetFocus
    SendKeys "{Home}+{End}"
    
Else
projectadb.Recordset.AddNew
MsgBox "Successfully Submited!", vbOKOnly + vbInformation, "Success"

Label4.Caption = Date
End If
End Sub
