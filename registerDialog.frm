VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form registerDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register Now | Student"
   ClientHeight    =   7215
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6123
   Icon            =   "registerDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   6123
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4797
      Top             =   6669
      _ExtentX        =   839
      _ExtentY        =   839
      _Version        =   393216
   End
   Begin VB.TextBox pwText 
      DataField       =   "pass"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      IMEMode         =   3  'DISABLE
      Left            =   3510
      PasswordChar    =   "*"
      TabIndex        =   22
      Top             =   5967
      Width           =   2353
   End
   Begin MSAdodcLib.Adodc regado 
      Height          =   299
      Left            =   351
      Top             =   6318
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
      Connect         =   $"registerDialog.frx":2FACA
      OLEDBString     =   $"registerDialog.frx":2FB55
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from register_user"
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
   Begin VB.CommandButton uploadCommand 
      BackColor       =   &H8000000D&
      Caption         =   "Upload Image"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3042
      Width           =   1300
   End
   Begin VB.TextBox addressText 
      DataField       =   "address"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   832
      Left            =   117
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   5382
      Width           =   2353
   End
   Begin VB.ComboBox sessionCombo 
      DataField       =   "session"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   117
      TabIndex        =   18
      Text            =   "Select Session"
      Top             =   4680
      Width           =   2353
   End
   Begin VB.TextBox phoneText 
      DataField       =   "mobile"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   117
      TabIndex        =   16
      Top             =   3861
      Width           =   2353
   End
   Begin VB.ComboBox genderCombo 
      DataField       =   "gender"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3510
      TabIndex        =   15
      Text            =   "Choose Gender"
      Top             =   5382
      Width           =   2470
   End
   Begin VB.TextBox emailText 
      DataField       =   "email"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   3510
      TabIndex        =   14
      Top             =   3978
      Width           =   2470
   End
   Begin VB.ComboBox courseCombo 
      DataField       =   "course"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   117
      TabIndex        =   13
      Text            =   "Select Course"
      Top             =   3159
      Width           =   2353
   End
   Begin VB.TextBox rollText 
      DataField       =   "rollno"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   3510
      TabIndex        =   12
      Top             =   4680
      Width           =   2470
   End
   Begin VB.TextBox nameText 
      DataField       =   "namee"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   117
      TabIndex        =   11
      Top             =   2457
      Width           =   2353
   End
   Begin VB.TextBox usernameText 
      DataField       =   "username"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   117
      TabIndex        =   9
      Top             =   1755
      Width           =   2353
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H8000000D&
      Caption         =   "SUBMIT FOR REVIEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   1989
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6552
      Width           =   2002
   End
   Begin VB.Label showimglink 
      DataField       =   "image"
      DataSource      =   "regado"
      Height          =   247
      Left            =   2340
      TabIndex        =   24
      Top             =   1404
      Visible         =   0   'False
      Width           =   1183
   End
   Begin VB.Label datelabel 
      DataField       =   "datee"
      DataSource      =   "regado"
      Height          =   247
      Left            =   351
      TabIndex        =   23
      Top             =   6786
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   4329
      TabIndex        =   21
      Top             =   5733
      Width           =   949
   End
   Begin VB.Image Image1 
      Height          =   1534
      Left            =   3861
      Stretch         =   -1  'True
      Top             =   1404
      Width           =   1768
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Session:"
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
      Left            =   585
      TabIndex        =   17
      Top             =   4329
      Width           =   1300
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Students Registration Form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.23
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   364
      Left            =   1521
      TabIndex        =   10
      Top             =   1053
      Width           =   3406
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
      Left            =   585
      TabIndex        =   8
      Top             =   2223
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
      Left            =   585
      TabIndex        =   7
      Top             =   1404
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
      Left            =   4329
      TabIndex        =   6
      Top             =   4446
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
      Left            =   4329
      TabIndex        =   5
      Top             =   3627
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
      Left            =   585
      TabIndex        =   4
      Top             =   5031
      Width           =   1066
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
      Left            =   585
      TabIndex        =   3
      Top             =   2925
      Width           =   1066
   End
   Begin VB.Label Label152 
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
      Left            =   4329
      TabIndex        =   2
      Top             =   5148
      Width           =   949
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
      Left            =   585
      TabIndex        =   1
      Top             =   3510
      Width           =   1183
   End
   Begin VB.Image Image2 
      Height          =   1183
      Left            =   0
      Picture         =   "registerDialog.frx":2FBE0
      Stretch         =   -1  'True
      Top             =   -117
      Width           =   6097
   End
End
Attribute VB_Name = "registerDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub datelabel_Click()
datelabel.Caption = Date
End Sub

Private Sub Form_Load()
courseCombo.AddItem "BCA"
courseCombo.AddItem "BBM"
courseCombo.AddItem "BCA IT"

sessionCombo.AddItem "2018-2021"
sessionCombo.AddItem "2019-2022"
sessionCombo.AddItem "2020-2023"
sessionCombo.AddItem "2021-2024"

genderCombo.AddItem "Male"
genderCombo.AddItem "Female"

regado.Recordset.AddNew
datelabel.Caption = Date
End Sub



Private Sub OKButton_Click()
If usernameText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
usernameText.SetFocus
SendKeys "{Home}+{End}"

ElseIf nameText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
nameText.SetFocus
SendKeys "{Home}+{End}"

ElseIf courseCombo = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
courseCombo.SetFocus
SendKeys "{Home}+{End}"

ElseIf phoneText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
phoneText.SetFocus
SendKeys "{Home}+{End}"

ElseIf sessionCombo = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
sessionCombo.SetFocus
SendKeys "{Home}+{End}"

ElseIf addressText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
addressText.SetFocus
SendKeys "{Home}+{End}"

ElseIf emailText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
emailText.SetFocus
SendKeys "{Home}+{End}"

ElseIf rollText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
rollText.SetFocus
SendKeys "{Home}+{End}"

ElseIf genderCombo = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
genderCombo.SetFocus
SendKeys "{Home}+{End}"

ElseIf pwText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
pwText.SetFocus
SendKeys "{Home}+{End}"

ElseIf showimglink.Caption = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
uploadCommand.SetFocus
SendKeys "{Home}+{End}"

Else
regado.Recordset.AddNew
MsgBox "Successfully Submited!", vbOKOnly + vbInformation, "Success"
End If
End Sub

Private Sub showimglink_Change()
Image1.Picture = LoadPicture(showimglink.Caption)
End Sub

Private Sub uploadCommand_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
showimglink.Caption = CommonDialog1.FileName
Image1.Picture = LoadPicture(showimglink.Caption)
MsgBox "Image Selected, Now please click on Submit For Review button to save!", vbOKOnly + vbInformation, "Success"
End Sub
