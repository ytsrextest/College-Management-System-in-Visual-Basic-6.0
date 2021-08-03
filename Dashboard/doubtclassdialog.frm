VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form doubtclassdialog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doubt Class"
   ClientHeight    =   3263
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6032
   Icon            =   "doubtclassdialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3263
   ScaleWidth      =   6032
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc matchdetailsado 
      Height          =   364
      Left            =   3393
      Top             =   1170
      Visible         =   0   'False
      Width           =   1885
      _ExtentX        =   3475
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
      Connect         =   $"doubtclassdialog.frx":1084A
      OLEDBString     =   $"doubtclassdialog.frx":108D5
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
   Begin VB.ComboBox doubttiming 
      DataField       =   "timing"
      DataSource      =   "doubtadb"
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
      Left            =   117
      TabIndex        =   8
      Text            =   "Select Timing"
      Top             =   1053
      Width           =   2353
   End
   Begin MSAdodcLib.Adodc doubtadb 
      Height          =   364
      Left            =   3510
      Top             =   2808
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
      Connect         =   $"doubtclassdialog.frx":10960
      OLEDBString     =   $"doubtclassdialog.frx":109EB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "student_doubt"
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
   Begin VB.TextBox doubtroll 
      DataField       =   "rollno"
      DataSource      =   "doubtadb"
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
      Left            =   117
      TabIndex        =   7
      Top             =   2340
      Width           =   2353
   End
   Begin VB.TextBox doubttopic 
      DataField       =   "topic"
      DataSource      =   "doubtadb"
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
      Left            =   117
      TabIndex        =   5
      Top             =   1638
      Width           =   2353
   End
   Begin VB.ComboBox ChoosesubCombo 
      DataField       =   "subject"
      DataSource      =   "doubtadb"
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
      Left            =   117
      TabIndex        =   2
      Text            =   "Select Subject"
      Top             =   351
      Width           =   2353
   End
   Begin VB.CommandButton submitButton 
      BackColor       =   &H008080FF&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   585
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2808
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   247
      Left            =   3276
      TabIndex        =   10
      Top             =   351
      Visible         =   0   'False
      Width           =   1651
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "datee"
      DataSource      =   "doubtadb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   247
      Left            =   3510
      TabIndex        =   9
      Top             =   117
      Width           =   1768
   End
   Begin VB.Label Label4 
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
      Left            =   702
      TabIndex        =   6
      Top             =   2106
      Width           =   1534
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Topic"
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
      Left            =   468
      TabIndex        =   4
      Top             =   1404
      Width           =   1300
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Timing"
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
      Left            =   585
      TabIndex        =   3
      Top             =   702
      Width           =   1183
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   468
      TabIndex        =   1
      Top             =   117
      Width           =   1534
   End
   Begin VB.Image Image1 
      Height          =   3445
      Left            =   1521
      Picture         =   "doubtclassdialog.frx":10A76
      Stretch         =   -1  'True
      Top             =   -117
      Width           =   4459
   End
End
Attribute VB_Name = "doubtclassdialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ChoosesubCombo_Change()
ChoosesubCombo.Text = "Select Subject"
End Sub




Private Sub doubttiming_Change()
doubttiming.Text = "Select Timing"
End Sub

Private Sub Form_Load()
ChoosesubCombo.AddItem "Maths"
ChoosesubCombo.AddItem "Computer Fundamental"
ChoosesubCombo.AddItem "English"
ChoosesubCombo.AddItem "Hindi"
ChoosesubCombo.AddItem "Programming"
doubttiming.AddItem "10 AM"
doubttiming.AddItem "02 PM"
doubtadb.Recordset.AddNew
Label5.Caption = Date
Label6.Caption = studentfrmLogin.studenttxtUserName
End Sub



Private Sub Label5_Click()
Label5.Caption = Date
End Sub

Private Sub Label6_Click()
Label6.Caption = studentfrmLogin.studenttxtUserName
End Sub

Private Sub submitButton_Click()
matchdetailsado.RecordSource = "select * from user_data where username='" + Label6.Caption + "' and rollno='" + doubtroll.Text + "'"
matchdetailsado.Refresh

If matchdetailsado.Recordset.EOF Then
 MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
    doubtroll.SetFocus
    SendKeys "{Home}+{End}"
Else
doubtadb.Recordset.AddNew
MsgBox "Successfully Submited!", vbOKOnly + vbInformation, "Success"

'doubtclassdialog.Visible = False

End If

End Sub

