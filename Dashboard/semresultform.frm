VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form semresultform 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Semester Result"
   ClientHeight    =   3601
   ClientLeft      =   104
   ClientTop       =   416
   ClientWidth     =   5785
   Icon            =   "semresultform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3601
   ScaleWidth      =   5785
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc matctdetailsado 
      Height          =   364
      Left            =   3978
      Top             =   3042
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
      Connect         =   $"semresultform.frx":2FACA
      OLEDBString     =   $"semresultform.frx":2FB55
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
   Begin VB.CommandButton searchresult 
      BackColor       =   &H8000000D&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   1872
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2925
      Width           =   1651
   End
   Begin VB.TextBox semrollno 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   1521
      TabIndex        =   1
      Top             =   2223
      Width           =   2587
   End
   Begin VB.Label Label3 
      Height          =   247
      Left            =   4212
      TabIndex        =   4
      Top             =   1872
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label2 
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
      Left            =   2223
      TabIndex        =   2
      Top             =   1872
      Width           =   1183
   End
   Begin VB.Line Line1 
      X1              =   1170
      X2              =   4680
      Y1              =   1638
      Y2              =   1638
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEMESTER RESULT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12.23
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   364
      Left            =   1404
      TabIndex        =   0
      Top             =   1287
      Width           =   3289
   End
   Begin VB.Image Image1 
      Height          =   1521
      Left            =   234
      Picture         =   "semresultform.frx":2FBE0
      Stretch         =   -1  'True
      Top             =   -117
      Width           =   5278
   End
End
Attribute VB_Name = "semresultform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub studentusernamesem_Click()


End Sub

Private Sub Form_Load()
Label3.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub Label3_Click()
Label3.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub searchresult_Click()
matctdetailsado.RecordSource = "select * from user_data where username='" + Label3.Caption + "' and rollno='" + semrollno.Text + "'"
matctdetailsado.Refresh

DataEnvironment2semresult.rsresultgen.Open "select * from sem_result where rollno= '" + semrollno.Text + "'"
studentresult.Refresh
If DataEnvironment2semresult.rsresultgen.EOF Then
MsgBox "Roll. No. Doesn't Exist, Please Try Again", vbCritical, "Massage"
'semresultform.Visible = False
DataEnvironment2semresult.rsresultgen.Close

ElseIf matctdetailsado.Recordset.EOF Then
 MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
    semrollno.SetFocus
    SendKeys "{Home}+{End}"

Else
studentresult.Show
DataEnvironment2semresult.rsresultgen.Close
semresultform.Visible = False
End If

End Sub
