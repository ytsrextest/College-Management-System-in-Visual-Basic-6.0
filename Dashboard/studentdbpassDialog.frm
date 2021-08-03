VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form studentdbpassDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Library Pass"
   ClientHeight    =   3224
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6032
   Icon            =   "studentdbpassDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "studentdbpassDialog.frx":1084A
   ScaleHeight     =   3224
   ScaleWidth      =   6032
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc matctdetailsado 
      Height          =   364
      Left            =   2457
      Top             =   2808
      Visible         =   0   'False
      Width           =   1092
      _ExtentX        =   2013
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
      Connect         =   $"studentdbpassDialog.frx":1AA6C
      OLEDBString     =   $"studentdbpassDialog.frx":1AAF7
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
   Begin VB.OptionButton acceptterm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Accept Term and Condition"
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
      Left            =   351
      TabIndex        =   12
      Top             =   2340
      Width           =   2353
   End
   Begin MSAdodcLib.Adodc adbresult 
      Height          =   299
      Left            =   3861
      Top             =   2574
      Visible         =   0   'False
      Width           =   1300
      _ExtentX        =   2396
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
      Connect         =   $"studentdbpassDialog.frx":1AB82
      OLEDBString     =   $"studentdbpassDialog.frx":1AC0D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from library_pass"
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
   Begin VB.CommandButton gosearch 
      BackColor       =   &H8000000D&
      Caption         =   "Generate"
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
      Left            =   1989
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1872
      Width           =   1066
   End
   Begin VB.TextBox entersearch 
      Alignment       =   2  'Center
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
      TabIndex        =   9
      Top             =   1287
      Width           =   2353
   End
   Begin MSAdodcLib.Adodc libraryadb 
      Height          =   299
      Left            =   3627
      Top             =   702
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
      CommandType     =   2
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
      Connect         =   $"studentdbpassDialog.frx":1AC98
      OLEDBString     =   $"studentdbpassDialog.frx":1AD23
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "library_pass"
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
   Begin VB.TextBox libraryroll 
      DataField       =   "rollno"
      DataSource      =   "libraryadb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   234
      TabIndex        =   7
      Top             =   1872
      Width           =   2587
   End
   Begin VB.ComboBox passCombo 
      DataField       =   "subject"
      DataSource      =   "libraryadb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   273
      Left            =   234
      TabIndex        =   5
      Text            =   "Select Subject"
      Top             =   1287
      Width           =   2587
   End
   Begin VB.TextBox passbookname 
      DataField       =   "bookname"
      DataSource      =   "libraryadb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   234
      TabIndex        =   3
      Top             =   585
      Width           =   2587
   End
   Begin VB.CommandButton generateassButton 
      BackColor       =   &H008080FF&
      Caption         =   "Generate Pass"
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
      Left            =   585
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2691
      Width           =   1651
   End
   Begin VB.Label Label6 
      Height          =   247
      Left            =   2808
      TabIndex        =   13
      Top             =   2574
      Visible         =   0   'False
      Width           =   949
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Roll. No."
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
      Left            =   1872
      TabIndex        =   11
      Top             =   936
      Width           =   1534
   End
   Begin VB.Label dateofissue 
      DataField       =   "datee"
      DataSource      =   "libraryadb"
      Height          =   247
      Left            =   3627
      TabIndex        =   8
      Top             =   351
      Visible         =   0   'False
      Width           =   1651
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Roll. No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   819
      TabIndex        =   6
      Top             =   1638
      Width           =   1300
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Subject"
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
      Left            =   819
      TabIndex        =   4
      Top             =   1053
      Width           =   1651
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Book Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   819
      TabIndex        =   2
      Top             =   351
      Width           =   1417
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please make sure every pass is valid for 3 days after the date of the issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   247
      Left            =   117
      TabIndex        =   1
      Top             =   117
      Width           =   5746
   End
   Begin VB.Image Image1 
      Height          =   3822
      Left            =   1872
      Picture         =   "studentdbpassDialog.frx":1ADAE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5161
   End
End
Attribute VB_Name = "studentdbpassDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub dateofissue_Click()
dateofissue.Caption = Date
End Sub

Private Sub Form_Load()
entersearch.Visible = False
gosearch.Visible = False
Label5.Visible = False
passCombo.AddItem "Maths"
passCombo.AddItem "Computer Fundamental"
passCombo.AddItem "English"
passCombo.AddItem "Hindi"
passCombo.AddItem "Programming"
libraryadb.Recordset.AddNew
dateofissue.Caption = Date
Label6.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub generateassButton_Click()
matctdetailsado.RecordSource = "select * from user_data where username='" + Label6.Caption + "' and rollno='" + libraryroll.Text + "'"
matctdetailsado.Refresh
If acceptterm = False Then
MsgBox "Please fill all details", vbCritical, "Massage"
passbookname.SetFocus
SendKeys "{Home}+{End}"

ElseIf matctdetailsado.Recordset.EOF Then
 MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
    libraryroll.SetFocus
    SendKeys "{Home}+{End}"

Else
libraryadb.Recordset.AddNew
MsgBox "Successfully Submited!", vbOKOnly + vbInformation, "Success"

Label2.Visible = False
passbookname.Visible = False
Label3.Visible = False
passCombo.Visible = False
Label4.Visible = False
libraryroll.Visible = False
acceptterm.Visible = False
generateassButton.Visible = False


entersearch.Visible = True
gosearch.Visible = True
Label5.Visible = True
End If

End Sub


Private Sub gosearch_Click()
DataEnvironment1.rspassgen.Open "select * from library_pass where rollno= '" + entersearch.Text + "'"
showpassDataReport.Refresh
If DataEnvironment1.rspassgen.EOF Then
MsgBox "Pass doesn't generated", vbCritical, "Massage"
Else
studentdbpassDialog.Visible = False
showpassDataReport.Show
DataEnvironment1.rspassgen.Close
entersearch.Visible = False
gosearch.Visible = False
Label5.Visible = False
Label5.Visible = False
End If

End Sub

Private Sub List1_Click()

End Sub

Private Sub Label6_Click()
Label6.Caption = studentfrmLogin.studenttxtUserName
End Sub
