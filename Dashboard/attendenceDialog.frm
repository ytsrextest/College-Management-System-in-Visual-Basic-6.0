VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form attendenceDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Attendence"
   ClientHeight    =   3640
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6006
   Icon            =   "attendenceDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3640
   ScaleWidth      =   6006
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton searchattendence 
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
      Left            =   2223
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2808
      Width           =   1651
   End
   Begin VB.TextBox attendencerollno 
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
      Height          =   481
      Left            =   1755
      TabIndex        =   2
      Top             =   2223
      Width           =   2587
   End
   Begin MSAdodcLib.Adodc matctdetailsado 
      Height          =   364
      Left            =   4329
      Top             =   2808
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
      Connect         =   $"attendenceDialog.frx":2FACA
      OLEDBString     =   $"attendenceDialog.frx":2FB55
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
   Begin VB.Label Label3 
      Height          =   247
      Left            =   4329
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
      Left            =   2457
      TabIndex        =   1
      Top             =   1989
      Width           =   1183
   End
   Begin VB.Line Line1 
      X1              =   1287
      X2              =   4797
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT ATTENDENCE"
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
      Top             =   1404
      Width           =   3289
   End
   Begin VB.Image Image1 
      Height          =   1521
      Left            =   351
      Picture         =   "attendenceDialog.frx":2FBE0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5278
   End
End
Attribute VB_Name = "attendenceDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub searchresult_Click()

End Sub

Private Sub Form_Load()
Label3.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub Label3_Click()
Label3.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub searchattendence_Click()
matctdetailsado.RecordSource = "select * from user_data where username='" + Label3.Caption + "' and rollno='" + attendencerollno.Text + "'"
matctdetailsado.Refresh

DataEnvironment2semresult.rsresultgen.Open "select * from sem_result where rollno= '" + attendencerollno.Text + "'"
studentresult.Refresh
If DataEnvironment2semresult.rsresultgen.EOF Then
MsgBox "Roll. No. Doesn't Exist, Please Try Again", vbCritical, "Massage"
DataEnvironment2semresult.rsresultgen.Close
attendenceDialog.Show


ElseIf matctdetailsado.Recordset.EOF Then
 MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
 DataEnvironment2semresult.rsresultgen.Close
    attendencerollno.SetFocus
    SendKeys "{Home}+{End}"
Else
attendenceDialog.Visible = False
studentresult.Show
DataEnvironment2semresult.rsresultgen.Close

End If
End Sub
